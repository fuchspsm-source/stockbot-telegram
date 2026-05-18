"""
StockBot Telegram v5.1 — Full AI Agent with Tool Calling
Powered by Gemini Flash (google-genai SDK).
AI decides when & how to search data — no manual keyword filtering.
"""
import os
import re
import io
import json
import logging
import pandas as pd
import requests
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from google import genai
from google.genai import types

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
TELEGRAM_TOKEN = os.environ.get('TELEGRAM_TOKEN', '')
GDRIVE_FILE_ID = os.environ.get('GDRIVE_FILE_ID', '')
GDRIVE_COGS_FILE_ID = os.environ.get('GDRIVE_COGS_FILE_ID', '')
ALLOWED_USER_IDS = os.environ.get('ALLOWED_USER_IDS', '')
GEMINI_API_KEY = os.environ.get('GEMINI_API_KEY', '')

WAREHOUSE_COLS = ['ID30', 'ID40']
GEMINI_MODEL = 'gemini-2.5-flash'
MAX_TOOL_TURNS = 6

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

DATA = {
    'stock': None, 'stock_sheet': None,
    'cogs': None, 'cogs_sheet': None,
}
CHAT_HISTORY = {}
MAX_HISTORY = 10

gemini_client = None
if GEMINI_API_KEY:
    gemini_client = genai.Client(api_key=GEMINI_API_KEY)
    logger.info(f"✅ Gemini client initialized ({GEMINI_MODEL})")
else:
    logger.warning("⚠️ GEMINI_API_KEY belum di-set!")


# ─────────────────────────────────────────────
# GOOGLE DRIVE LOADER
# ─────────────────────────────────────────────
def download_from_gdrive(file_id, as_gsheet=False):
    if as_gsheet:
        url = f'https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx'
    else:
        url = f'https://drive.google.com/uc?export=download&id={file_id}'
    response = requests.get(url, allow_redirects=True, timeout=60)
    if not as_gsheet and 'confirm=' in response.text:
        match = re.search(r'confirm=([0-9A-Za-z_]+)', response.text)
        if match:
            url = f'https://drive.google.com/uc?export=download&confirm={match.group(1)}&id={file_id}'
            response = requests.get(url, allow_redirects=True, timeout=60)
    return io.BytesIO(response.content)


def load_stock_data():
    try:
        logger.info("📥 Loading STOCK...")
        try:
            excel_data = download_from_gdrive(GDRIVE_FILE_ID)
            xl = pd.ExcelFile(excel_data)
        except Exception:
            excel_data = download_from_gdrive(GDRIVE_FILE_ID, as_gsheet=True)
            xl = pd.ExcelFile(excel_data)

        sheets = xl.sheet_names
        if not sheets:
            return False, "File stok tidak punya sheet."
        latest_sheet = sheets[-1]
        df = pd.read_excel(excel_data, sheet_name=latest_sheet)
        cols = list(df.columns)
        if str(cols[0]).strip() == '' or 'Unnamed' in str(cols[0]):
            df = df.rename(columns={cols[0]: 'Material'})

        code_col = None
        for col in df.columns:
            if 'CODE' in str(col).upper() or str(col).upper().strip() == 'MATERIAL':
                code_col = col
                break

        for col in WAREHOUSE_COLS:
            if col not in df.columns:
                df[col] = 0
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        df['Total'] = df[WAREHOUSE_COLS].sum(axis=1)
        df['desc_clean'] = df['Material Description'].astype(str).str.strip().str.upper()
        df['code_clean'] = df[code_col].astype(str).str.strip().str.upper() if code_col else ''

        DATA['stock'] = df
        DATA['stock_sheet'] = latest_sheet
        logger.info(f"✅ STOCK: {len(df)} products from '{latest_sheet}'")
        return True, f"Stok: sheet '{latest_sheet}', {len(df):,} produk"
    except Exception as e:
        logger.error(f"❌ Stock load error: {e}")
        return False, f"Error stock: {str(e)}"


def load_cogs_data():
    if not GDRIVE_COGS_FILE_ID:
        return False, "COGS file ID belum di-set."
    try:
        logger.info("📥 Loading COGS...")
        try:
            excel_data = download_from_gdrive(GDRIVE_COGS_FILE_ID)
            xl = pd.ExcelFile(excel_data)
        except Exception:
            excel_data = download_from_gdrive(GDRIVE_COGS_FILE_ID, as_gsheet=True)
            xl = pd.ExcelFile(excel_data)

        sheets = xl.sheet_names
        latest_sheet = sheets[-1]
        df = pd.read_excel(excel_data, sheet_name=latest_sheet)
        required = ['Material Code', 'Material Description', 'Source', 'COGS', 'Update', 'Rata2 NC Sebelumnya']
        missing = [c for c in required if c not in df.columns]
        if missing:
            return False, f"Kolom COGS missing: {missing}"

        df['COGS'] = pd.to_numeric(df['COGS'], errors='coerce').fillna(0)
        df['Material Code'] = df['Material Code'].astype(str).str.strip()
        df['desc_clean'] = df['Material Description'].astype(str).str.strip().str.upper()
        df['code_clean'] = df['Material Code'].str.upper()
        df['Rata2 NC Sebelumnya'] = df['Rata2 NC Sebelumnya'].astype(str).str.strip()
        df = df[df['Material Code'].notna() & (df['Material Code'] != '') & (df['Material Code'] != 'nan')].reset_index(drop=True)

        DATA['cogs'] = df
        DATA['cogs_sheet'] = latest_sheet
        logger.info(f"✅ COGS: {len(df)} entries from '{latest_sheet}'")
        return True, f"COGS: sheet '{latest_sheet}', {len(df):,} entries"
    except Exception as e:
        logger.error(f"❌ COGS load error: {e}")
        return False, f"Error COGS: {str(e)}"


def load_all_data():
    stock_ok, stock_msg = load_stock_data()
    cogs_ok, cogs_msg = load_cogs_data()
    return stock_ok, f"{'✅' if stock_ok else '❌'} {stock_msg}\n{'✅' if cogs_ok else '⚠️'} {cogs_msg}"


# ─────────────────────────────────────────────
# TOOLS — Fungsi yang dipanggil AI
# ─────────────────────────────────────────────
def _stock_row(row):
    return {
        'code': str(row.get('code_clean', '')),
        'description': str(row['Material Description']).strip(),
        'ID30': int(row['ID30']), 'ID40': int(row['ID40']),
        'total': int(row['Total'])
    }

def _cogs_row(row):
    nc = str(row.get('Rata2 NC Sebelumnya', '')).strip()
    return {
        'code': str(row['Material Code']),
        'description': str(row['Material Description']).strip(),
        'source': str(row['Source']).strip(),
        'cogs_rupiah': int(row['COGS']),
        'update_label': str(row['Update']).strip(),
        'rata2_nc_sebelumnya': nc if nc != 'nan' else ''
    }


def tool_search_stock(keywords: str, limit: int = 30) -> dict:
    """AND search di Material Description tabel STOK."""
    if DATA['stock'] is None:
        return {'error': 'Data stok belum ter-load', 'results': []}
    df = DATA['stock']
    kw_list = [k.strip().upper() for k in keywords.split() if k.strip()]
    if not kw_list:
        return {'error': 'Keywords kosong', 'results': []}
    mask = pd.Series([True] * len(df), index=df.index)
    for kw in kw_list:
        mask &= df['desc_clean'].str.contains(re.escape(kw), na=False)
    matched = df[mask]
    total = len(matched)
    return {
        'total_matches': total,
        'returned': min(total, limit),
        'truncated': total > limit,
        'results': [_stock_row(row) for _, row in matched.head(limit).iterrows()]
    }


def tool_search_cogs(keywords: str, limit: int = 30) -> dict:
    """AND search di Material Description tabel COGS."""
    if DATA['cogs'] is None:
        return {'error': 'Data COGS belum ter-load', 'results': []}
    df = DATA['cogs']
    kw_list = [k.strip().upper() for k in keywords.split() if k.strip()]
    if not kw_list:
        return {'error': 'Keywords kosong', 'results': []}
    mask = pd.Series([True] * len(df), index=df.index)
    for kw in kw_list:
        mask &= df['desc_clean'].str.contains(re.escape(kw), na=False)
    matched = df[mask]
    total = len(matched)
    return {
        'total_matches': total,
        'returned': min(total, limit),
        'truncated': total > limit,
        'results': [_cogs_row(row) for _, row in matched.head(limit).iterrows()]
    }


def tool_get_by_code(code: str) -> dict:
    """Lookup exact material code di STOK dan COGS sekaligus."""
    code = code.strip().upper()
    stock_results, cogs_results = [], []
    if DATA['stock'] is not None:
        stock_results = [_stock_row(r) for _, r in DATA['stock'][DATA['stock']['code_clean'] == code].iterrows()]
    if DATA['cogs'] is not None:
        cogs_results = [_cogs_row(r) for _, r in DATA['cogs'][DATA['cogs']['code_clean'] == code].iterrows()]
    return {'code': code, 'stock': stock_results, 'cogs': cogs_results,
            'stock_count': len(stock_results), 'cogs_count': len(cogs_results)}


def tool_get_analytics() -> dict:
    """Ringkasan analytics: total, kosong, top 10 terbanyak & tersedikit."""
    if DATA['stock'] is None:
        return {'error': 'Data stok belum ter-load'}
    df = DATA['stock']
    return {
        'total_produk': len(df),
        'stok_kosong': int((df['Total'] == 0).sum()),
        'ada_stok': int((df['Total'] > 0).sum()),
        'total_id30': int(df['ID30'].sum()),
        'total_id40': int(df['ID40'].sum()),
        'grand_total': int(df['Total'].sum()),
        'top_10_terbanyak': [_stock_row(r) for _, r in df[df['Total'] > 0].nlargest(10, 'Total').iterrows()],
        'top_10_tersedikit': [_stock_row(r) for _, r in df[df['Total'] > 0].nsmallest(10, 'Total').iterrows()],
    }


def tool_list_empty_stock(keywords: str = '', limit: int = 50) -> dict:
    """List produk stok kosong, optional filter keyword."""
    if DATA['stock'] is None:
        return {'error': 'Data stok belum ter-load', 'results': []}
    df = DATA['stock']
    empty = df[df['Total'] == 0]
    if keywords and keywords.strip():
        for kw in keywords.upper().split():
            empty = empty[empty['desc_clean'].str.contains(re.escape(kw), na=False)]
    total = len(empty)
    return {
        'total_matches': total,
        'returned': min(total, limit),
        'truncated': total > limit,
        'results': [_stock_row(r) for _, r in empty.head(limit).iterrows()]
    }


def tool_list_cogs_sources() -> dict:
    """List unique Source di COGS."""
    if DATA['cogs'] is None:
        return {'error': 'Data COGS belum ter-load'}
    sources = DATA['cogs']['Source'].value_counts().to_dict()
    return {'sources': {str(k): int(v) for k, v in sources.items()}}


TOOL_FUNCTIONS = {
    'search_stock': tool_search_stock,
    'search_cogs': tool_search_cogs,
    'get_by_code': tool_get_by_code,
    'get_analytics': tool_get_analytics,
    'list_empty_stock': tool_list_empty_stock,
    'list_cogs_sources': tool_list_cogs_sources,
}


# ─────────────────────────────────────────────
# GEMINI TOOL DECLARATIONS (google-genai SDK)
# ─────────────────────────────────────────────
def _schema(type_, description=None, **props):
    return types.Schema(type=type_, description=description, **props)

GEMINI_TOOLS = [types.Tool(function_declarations=[
    types.FunctionDeclaration(
        name='search_stock',
        description=(
            'Cari produk di tabel STOK berdasarkan keyword (AND search — semua kata harus ada). '
            'Return: code, description, ID30, ID40, total stok. '
            'Gunakan nama produk + kemasan, contoh: "RENOLIN B 68 PLUS 1000L IBC".'
        ),
        parameters=types.Schema(type='OBJECT', properties={
            'keywords': _schema('STRING', 'Keyword pencarian dipisah spasi. Contoh: "RENOLIN B 68 PLUS"'),
            'limit': _schema('INTEGER', 'Max hasil, default 30'),
        }, required=['keywords'])
    ),
    types.FunctionDeclaration(
        name='search_cogs',
        description=(
            'Cari di tabel COGS berdasarkan keyword (AND search). '
            'Return: code, description, source, cogs_rupiah, update_label, rata2_nc_sebelumnya. '
            'PENTING: 1 produk fisik bisa punya MULTI-SOURCE dengan code berbeda. '
            'Lihat SEMUA hasil, jangan ambil yang pertama saja.'
        ),
        parameters=types.Schema(type='OBJECT', properties={
            'keywords': _schema('STRING', 'Keyword pencarian. Contoh: "RENOLIN B 68 PLUS 1000L"'),
            'limit': _schema('INTEGER', 'Max hasil, default 30'),
        }, required=['keywords'])
    ),
    types.FunctionDeclaration(
        name='get_by_code',
        description=(
            'Lookup exact material code (6+ digit angka) di STOK dan COGS sekaligus. '
            'Pakai ini kalau user kasih kode langsung. '
            'PENTING: code di STOK dan COGS bisa berbeda untuk produk yang sama — '
            'kalau salah satu tidak ketemu, lanjut search by description.'
        ),
        parameters=types.Schema(type='OBJECT', properties={
            'code': _schema('STRING', 'Material code. Contoh: "602723550"'),
        }, required=['code'])
    ),
    types.FunctionDeclaration(
        name='get_analytics',
        description='Rekap analytics: total produk, stok kosong/ada, total per gudang, top 10 terbanyak & tersedikit.',
        parameters=types.Schema(type='OBJECT', properties={})
    ),
    types.FunctionDeclaration(
        name='list_empty_stock',
        description='List produk stok kosong (total=0), optional filter keyword.',
        parameters=types.Schema(type='OBJECT', properties={
            'keywords': _schema('STRING', 'Optional filter by keyword.'),
            'limit': _schema('INTEGER', 'Max hasil, default 50'),
        })
    ),
    types.FunctionDeclaration(
        name='list_cogs_sources',
        description='List semua unique Source (LOKAL, IMPORT, dll) di tabel COGS.',
        parameters=types.Schema(type='OBJECT', properties={})
    ),
])]


# ─────────────────────────────────────────────
# SYSTEM PROMPT
# ─────────────────────────────────────────────
def get_system_prompt():
    stock_info = f"{len(DATA['stock']):,} produk dari sheet '{DATA['stock_sheet']}'" if DATA['stock'] is not None else "belum ter-load"
    cogs_info = f"{len(DATA['cogs']):,} entries dari sheet '{DATA['cogs_sheet']}'" if DATA['cogs'] is not None else "belum ter-load"
    return f"""Kamu adalah StockBot, asisten AI untuk perusahaan pelumas. Chat di Telegram, santai tapi akurat.

DATA: Stok {stock_info} | COGS {cogs_info}
Gudang: ID30 dan ID40.

WAJIB PAKAI TOOLS — jangan ngarang data.

STRATEGI:
1. Untuk produk spesifik → panggil search_stock DAN search_cogs secara bersamaan.
2. Hasil tidak cocok? Coba lagi dengan keyword lebih/kurang spesifik atau variasi penulisan.
3. Kode 6+ digit → get_by_code dulu, lalu search_cogs by description kalau perlu.
4. truncated=true → narrow query.
5. 1 produk bisa punya multi-kemasan (20L, 205L, 1000L) dan multi-source di COGS.
6. Material code di STOK dan COGS bisa beda untuk produk yang sama.

FORMULA: NC% = (Harga Jual - COGS) / Harga Jual × 100 | Harga Jual = COGS / (1 - NC%/100)

OUTPUT:
- Bahasa Indonesia natural, ngobrol biasa.
- Rupiah: Rp 36.934.577 (pakai titik pemisah ribuan).
- Markdown Telegram: *bold*, `code`. Emoji 1-2, jangan lebay.
- Kalau multi varian/source, list semua.
- Jangan tutup dengan "Ada lagi yang bisa dibantu?" kecuali relevan.
- Jujur kalau tidak ketemu setelah beberapa kali coba."""


# ─────────────────────────────────────────────
# AGENT LOOP
# ─────────────────────────────────────────────
def call_tool(name: str, args: dict) -> dict:
    func = TOOL_FUNCTIONS.get(name)
    if not func:
        return {'error': f'Unknown tool: {name}'}
    try:
        result = func(**args)
        logger.info(f"🔧 {name}({args}) → {result.get('total_matches', result.get('stock_count', 'ok'))}")
        return result
    except Exception as e:
        logger.error(f"❌ Tool {name} error: {e}")
        return {'error': str(e)}


def ask_gemini(user_input: str, user_id: int) -> str:
    if not gemini_client:
        return '⚠️ Fitur AI belum aktif. Set GEMINI_API_KEY di environment.'

    try:
        # Build history dalam format google-genai
        history = CHAT_HISTORY.get(user_id, [])
        contents = []
        for msg in history:
            role = "model" if msg["role"] == "assistant" else "user"
            contents.append(types.Content(role=role, parts=[types.Part(text=msg["content"])]))

        # Tambah pesan user
        contents.append(types.Content(role="user", parts=[types.Part(text=user_input)]))

        # Config
        config = types.GenerateContentConfig(
            system_instruction=get_system_prompt(),
            tools=GEMINI_TOOLS,
            max_output_tokens=2000,
        )

        # Agent loop
        for turn in range(MAX_TOOL_TURNS):
            response = gemini_client.models.generate_content(
                model=GEMINI_MODEL,
                contents=contents,
                config=config,
            )

            candidate = response.candidates[0]

            # Cek apakah ada function calls
            function_calls = [
                p.function_call for p in candidate.content.parts
                if p.function_call is not None
            ]

            if not function_calls:
                # Tidak ada tool call lagi — ambil text final
                text_parts = [p.text for p in candidate.content.parts if p.text]
                reply = '\n'.join(text_parts).strip()
                break

            # Tambah response model ke history lokal (termasuk tool calls)
            contents.append(candidate.content)

            # Execute tool calls dan buat function responses
            tool_response_parts = []
            for fc in function_calls:
                args = dict(fc.args) if fc.args else {}
                result = call_tool(fc.name, args)
                tool_response_parts.append(
                    types.Part.from_function_response(name=fc.name, response=result)
                )

            # Tambah tool responses ke contents
            contents.append(types.Content(role="user", parts=tool_response_parts))

        else:
            reply = '⚠️ AI butuh terlalu banyak langkah. Coba pertanyaan yang lebih spesifik.'

        if not reply:
            reply = '⚠️ AI tidak menghasilkan jawaban. Coba ulang.'

        # Simpan ke history (hanya user text + assistant text)
        history.append({"role": "user", "content": user_input})
        history.append({"role": "assistant", "content": reply})
        if len(history) > MAX_HISTORY * 2:
            history = history[-(MAX_HISTORY * 2):]
        CHAT_HISTORY[user_id] = history

        logger.info(f"🤖 Reply ke {user_id}: {len(reply)} chars")
        return reply

    except Exception as e:
        logger.error(f"❌ Gemini error: {e}", exc_info=True)
        return f'⚠️ AI error: {str(e)[:200]}\n\nCoba ulangi atau /reload.'


def clear_history(user_id):
    if user_id in CHAT_HISTORY:
        del CHAT_HISTORY[user_id]
        return '🔄 Chat history di-reset.'
    return '✅ Tidak ada history aktif.'


# ─────────────────────────────────────────────
# TELEGRAM HANDLERS
# ─────────────────────────────────────────────
def is_allowed(user_id):
    if not ALLOWED_USER_IDS:
        return True
    return user_id in [int(x.strip()) for x in ALLOWED_USER_IDS.split(',') if x.strip()]


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id):
        await update.message.reply_text(f'❌ Akses ditolak. ID: `{update.effective_user.id}`', parse_mode='Markdown')
        return
    msg = (
        '👋 *StockBot v5.1 — Full AI Agent*\n\n'
        '🤖 Powered by Gemini + Tool Calling\n\n'
        'Ngobrol aja kayak biasa — saya cari datanya sendiri.\n\n'
        '*Contoh:*\n'
        '• renolin b 68 plus 1000 berapa costnya?\n'
        '• stok titan plus 205l?\n'
        '• 602723550 ini barang apa?\n'
        '• produk apa yang stoknya paling banyak?\n\n'
        '/reload /status /clear'
    )
    await update.message.reply_text(msg, parse_mode='Markdown')


async def reload_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id): return
    await update.message.reply_text('🔄 Loading dari Google Drive...')
    ok, msg = load_all_data()
    await update.message.reply_text(f"{'✅' if ok else '❌'} Hasil:\n\n{msg}")


async def status_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id): return
    lines = ['📊 *Status:*\n']
    if DATA['stock'] is not None:
        df = DATA['stock']
        lines += [f'✅ *STOK:* `{DATA["stock_sheet"]}` — {len(df):,} produk',
                  f'   Ada: {(df["Total"]>0).sum():,} | Kosong: {(df["Total"]==0).sum():,}']
    else:
        lines.append('❌ *STOK:* belum ter-load')
    lines.append('')
    if DATA['cogs'] is not None:
        df = DATA['cogs']
        lines += [f'✅ *COGS:* `{DATA["cogs_sheet"]}` — {len(df):,} entries',
                  f'   Sources: {df["Source"].nunique()}']
    else:
        lines.append('⚠️ *COGS:* belum ter-load')
    lines.append('')
    lines.append(f'🤖 *AI:* {GEMINI_MODEL} (Agent Mode)' if gemini_client else '⚠️ *AI:* off')
    uid = update.effective_user.id
    if uid in CHAT_HISTORY:
        lines.append(f'💬 *History:* {len(CHAT_HISTORY[uid])//2} turn')
    await update.message.reply_text('\n'.join(lines), parse_mode='Markdown')


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id): return
    await update.message.reply_text(
        '📖 *StockBot v5.1*\n\nTanya apa aja pakai bahasa natural. '
        'AI cari datanya sendiri.\n\n/reload /status /clear',
        parse_mode='Markdown'
    )


async def clear_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id): return
    await update.message.reply_text(clear_history(update.effective_user.id))


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        await update.message.reply_text(f'❌ Akses ditolak. ID: `{user_id}`', parse_mode='Markdown')
        return

    user_text = update.message.text
    logger.info(f"Query from {user_id}: {user_text}")

    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action='typing')

    if user_text.upper().strip() in ['STOP', 'CLEAR', 'RESET']:
        await update.message.reply_text(clear_history(user_id))
        return

    reply = ask_gemini(user_text, user_id)

    if len(reply) > 4000:
        for chunk in [reply[i:i+4000] for i in range(0, len(reply), 4000)]:
            try:
                await update.message.reply_text(chunk, parse_mode='Markdown')
            except Exception:
                await update.message.reply_text(chunk)
    else:
        try:
            await update.message.reply_text(reply, parse_mode='Markdown')
        except Exception:
            await update.message.reply_text(reply)


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    if not TELEGRAM_TOKEN:
        logger.error("❌ TELEGRAM_TOKEN tidak di-set!")
        return
    if not GDRIVE_FILE_ID:
        logger.error("❌ GDRIVE_FILE_ID tidak di-set!")
        return
    if not GEMINI_API_KEY:
        logger.error("❌ GEMINI_API_KEY tidak di-set!")
        return

    logger.info("🚀 Starting StockBot v5.1 (Agent Mode)...")
    _, msg = load_all_data()
    logger.info(f"Initial load:\n{msg}")

    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("reload", reload_cmd))
    app.add_handler(CommandHandler("status", status_cmd))
    app.add_handler(CommandHandler("clear", clear_cmd))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    logger.info("✅ Bot running...")
    app.run_polling(drop_pending_updates=True)


if __name__ == '__main__':
    main()
