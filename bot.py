"""
StockBot Telegram v5.2 — Actual Stock = Open Stock - Sales
Powered by Gemini Flash (google-genai SDK).
AI decides when & how to search data — no manual keyword filtering.

Perubahan v5.2:
- Load 2 tab: tab DDMMYY (open stock) + tab SALES (penjualan berjalan)
- Actual Stock = ID30 (open stock) - Total Sales per material
- Semua query stok pakai Actual Stock
- COGS logic tidak diubah sama sekali
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
    'sales': None, 'sales_sheet': None,
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


def _get_excel_file(file_id):
    """Download GDrive file, coba direct dulu lalu fallback ke gsheet export."""
    try:
        data = download_from_gdrive(file_id)
        xl = pd.ExcelFile(data)
        return data, xl
    except Exception:
        data = download_from_gdrive(file_id, as_gsheet=True)
        xl = pd.ExcelFile(data)
        return data, xl


def _find_sales_sheet(sheet_names):
    """Cari tab SALES (case-insensitive). Return nama sheet atau None."""
    for name in sheet_names:
        if 'SALES' in str(name).upper():
            return name
    return None


def _find_stock_sheet(sheet_names):
    """
    Cari tab stock: format DDMMYY (6 digit angka).
    Kalau tidak ada yang match format, ambil tab pertama yang bukan SALES.
    """
    ddmmyy = re.compile(r'^\d{6}$')
    candidates = [s for s in sheet_names if ddmmyy.match(str(s).strip())]
    if candidates:
        return candidates[-1]  # ambil yang paling kanan / terbaru
    # fallback: tab pertama yang bukan SALES
    for name in sheet_names:
        if 'SALES' not in str(name).upper():
            return name
    return sheet_names[0]


def load_stock_data():
    """
    Load open stock dari tab DDMMYY dan penjualan dari tab SALES.
    Hitung Actual Stock = ID30 - Total Sales per material code.
    """
    try:
        logger.info("📥 Loading STOCK + SALES...")
        excel_data, xl = _get_excel_file(GDRIVE_FILE_ID)
        sheets = xl.sheet_names

        if not sheets:
            return False, "File stok tidak punya sheet."

        # ── Identifikasi tab ──
        stock_sheet_name = _find_stock_sheet(sheets)
        sales_sheet_name = _find_sales_sheet(sheets)

        logger.info(f"📋 Tab stock: '{stock_sheet_name}' | Tab sales: '{sales_sheet_name}'")

        # ── Load Open Stock ──
        df_stock = pd.read_excel(excel_data, sheet_name=stock_sheet_name)
        cols = list(df_stock.columns)
        if str(cols[0]).strip() == '' or 'Unnamed' in str(cols[0]):
            df_stock = df_stock.rename(columns={cols[0]: 'Material'})

        # Deteksi kolom material code
        code_col = None
        for col in df_stock.columns:
            if 'CODE' in str(col).upper() or str(col).upper().strip() == 'MATERIAL':
                code_col = col
                break

        for col in WAREHOUSE_COLS:
            if col not in df_stock.columns:
                df_stock[col] = 0
            df_stock[col] = pd.to_numeric(df_stock[col], errors='coerce').fillna(0)

        df_stock['desc_clean'] = df_stock['Material Description'].astype(str).str.strip().str.upper()
        df_stock['code_clean'] = df_stock[code_col].astype(str).str.strip().str.upper() if code_col else ''

        # ── Load Sales & Hitung Total Sales per Material Code ──
        sales_by_code = {}  # code_clean → total qty terjual
        sales_sheet_loaded = None

        if sales_sheet_name:
            try:
                df_sales = pd.read_excel(excel_data, sheet_name=sales_sheet_name)

                # Deteksi kolom MATERIAL CODE dan QTY di tab SALES
                sales_code_col = None
                sales_qty_col = None

                for col in df_sales.columns:
                    col_up = str(col).upper().strip()
                    if 'CODE' in col_up or col_up == 'MATERIAL':
                        sales_code_col = col
                    if 'QTY' in col_up or 'QUANTITY' in col_up or 'JUMLAH' in col_up:
                        sales_qty_col = col

                if sales_code_col and sales_qty_col:
                    df_sales['_code'] = df_sales[sales_code_col].astype(str).str.strip().str.upper()
                    df_sales['_qty'] = pd.to_numeric(df_sales[sales_qty_col], errors='coerce').fillna(0)
                    sales_by_code = df_sales.groupby('_code')['_qty'].sum().to_dict()
                    sales_sheet_loaded = sales_sheet_name
                    DATA['sales'] = df_sales
                    DATA['sales_sheet'] = sales_sheet_name
                    logger.info(f"✅ SALES: {len(df_sales)} rows dari '{sales_sheet_name}', "
                                f"{len(sales_by_code)} material terjual")
                else:
                    logger.warning(f"⚠️ Tab SALES ditemukan tapi kolom tidak lengkap. "
                                   f"Kolom ada: {list(df_sales.columns)}")
            except Exception as e:
                logger.warning(f"⚠️ Gagal load tab SALES: {e}")
        else:
            logger.warning("⚠️ Tab SALES tidak ditemukan. Actual Stock = Open Stock.")

        # ── Hitung Actual Stock ──
        def get_actual(row):
            code = str(row['code_clean'])
            open_id30 = float(row['ID30'])
            sold = float(sales_by_code.get(code, 0))
            actual = open_id30 - sold
            return actual

        df_stock['Open_ID30'] = df_stock['ID30'].copy()
        df_stock['Sales_QTY'] = df_stock['code_clean'].map(lambda c: sales_by_code.get(c, 0))
        df_stock['Actual'] = (df_stock['Open_ID30'] - df_stock['Sales_QTY']).clip(lower=0)
        df_stock['Total'] = df_stock['Actual']  # Total sekarang = Actual untuk semua tools

        DATA['stock'] = df_stock
        DATA['stock_sheet'] = stock_sheet_name

        sales_info = f" | SALES: '{sales_sheet_loaded}'" if sales_sheet_loaded else " | SALES: tidak ditemukan"
        logger.info(f"✅ STOCK: {len(df_stock)} products dari '{stock_sheet_name}'{sales_info}")
        return True, (f"Stok: sheet '{stock_sheet_name}', {len(df_stock):,} produk"
                      + (f"\nSales: sheet '{sales_sheet_loaded}', {len(sales_by_code):,} material"
                         if sales_sheet_loaded else "\n⚠️ Tab SALES tidak ditemukan, pakai open stock."))

    except Exception as e:
        logger.error(f"❌ Stock load error: {e}", exc_info=True)
        return False, f"Error stock: {str(e)}"


def load_cogs_data():
    # ── TIDAK DIUBAH SAMA SEKALI ──
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
        logger.info(f"✅ COGS: {len(df)} entries dari '{latest_sheet}'")
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
    """
    Return actual stock. 'stock' yang ditampilkan = Actual (open - sales).
    ID30/ID40 raw tidak ditampilkan ke user.
    """
    return {
        'code': str(row.get('code_clean', '')),
        'description': str(row['Material Description']).strip(),
        'stock': int(row['Actual']),          # Actual = Open ID30 - Sales
        'open_stock': int(row['Open_ID30']),  # Info tambahan kalau AI butuh
        'sold': int(row['Sales_QTY']),        # Info tambahan kalau AI butuh
    }


def _cogs_row(row):
    # ── TIDAK DIUBAH ──
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
    """AND search di Material Description — return Actual Stock."""
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
    # ── TIDAK DIUBAH ──
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
    """Lookup exact material code — return Actual Stock + COGS."""
    code = code.strip().upper()
    stock_results, cogs_results = [], []
    if DATA['stock'] is not None:
        stock_results = [_stock_row(r) for _, r in DATA['stock'][DATA['stock']['code_clean'] == code].iterrows()]
    if DATA['cogs'] is not None:
        cogs_results = [_cogs_row(r) for _, r in DATA['cogs'][DATA['cogs']['code_clean'] == code].iterrows()]
    return {'code': code, 'stock': stock_results, 'cogs': cogs_results,
            'stock_count': len(stock_results), 'cogs_count': len(cogs_results)}


def tool_get_analytics() -> dict:
    """Ringkasan analytics pakai Actual Stock."""
    if DATA['stock'] is None:
        return {'error': 'Data stok belum ter-load'}
    df = DATA['stock']
    return {
        'total_produk': len(df),
        'stok_kosong': int((df['Actual'] == 0).sum()),
        'ada_stok': int((df['Actual'] > 0).sum()),
        'total_actual_stock': int(df['Actual'].sum()),
        'total_open_stock_id30': int(df['Open_ID30'].sum()),
        'total_terjual': int(df['Sales_QTY'].sum()),
        'top_10_terbanyak': [_stock_row(r) for _, r in df[df['Actual'] > 0].nlargest(10, 'Actual').iterrows()],
        'top_10_tersedikit': [_stock_row(r) for _, r in df[df['Actual'] > 0].nsmallest(10, 'Actual').iterrows()],
    }


def tool_list_empty_stock(keywords: str = '', limit: int = 50) -> dict:
    """List produk actual stock = 0, optional filter keyword."""
    if DATA['stock'] is None:
        return {'error': 'Data stok belum ter-load', 'results': []}
    df = DATA['stock']
    empty = df[df['Actual'] == 0]
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
    # ── TIDAK DIUBAH ──
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
# GEMINI TOOL DECLARATIONS
# ─────────────────────────────────────────────
def _schema(type_, description=None, **props):
    return types.Schema(type=type_, description=description, **props)

GEMINI_TOOLS = [types.Tool(function_declarations=[
    types.FunctionDeclaration(
        name='search_stock',
        description=(
            'Cari produk di tabel STOK berdasarkan keyword (AND search). '
            'Return: code, description, stock (ACTUAL = open stock - penjualan berjalan), '
            'open_stock, sold. '
            'Gunakan nama produk + kemasan. Contoh: "RENOLIN B 68 PLUS 1000L IBC".'
        ),
        parameters=types.Schema(type='OBJECT', properties={
            'keywords': _schema('STRING', 'Keyword pencarian dipisah spasi.'),
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
            'keywords': _schema('STRING', 'Keyword pencarian.'),
            'limit': _schema('INTEGER', 'Max hasil, default 30'),
        }, required=['keywords'])
    ),
    types.FunctionDeclaration(
        name='get_by_code',
        description=(
            'Lookup exact material code (6+ digit angka) di STOK dan COGS sekaligus. '
            'Pakai ini kalau user kasih kode langsung. '
            'PENTING: code di STOK dan COGS bisa berbeda untuk produk yang sama.'
        ),
        parameters=types.Schema(type='OBJECT', properties={
            'code': _schema('STRING', 'Material code. Contoh: "602723550"'),
        }, required=['code'])
    ),
    types.FunctionDeclaration(
        name='get_analytics',
        description='Rekap analytics: total produk, stok kosong/ada, total actual stock, top 10.',
        parameters=types.Schema(type='OBJECT', properties={})
    ),
    types.FunctionDeclaration(
        name='list_empty_stock',
        description='List produk actual stock = 0 (habis terjual atau memang tidak ada), optional filter keyword.',
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
    stock_info = (f"{len(DATA['stock']):,} produk dari sheet '{DATA['stock_sheet']}'"
                  if DATA['stock'] is not None else "belum ter-load")
    sales_info = (f"'{DATA['sales_sheet']}'" if DATA['sales_sheet'] else "tidak ditemukan")
    cogs_info = (f"{len(DATA['cogs']):,} entries dari sheet '{DATA['cogs_sheet']}'"
                 if DATA['cogs'] is not None else "belum ter-load")

    return f"""Kamu adalah StockBot, asisten AI untuk perusahaan pelumas. Chat di Telegram, santai tapi akurat.

DATA: Stok {stock_info} | Sales tab: {sales_info} | COGS {cogs_info}

WAJIB PAKAI TOOLS — jangan ngarang data.

STOK = ACTUAL STOCK:
- Field 'stock' di hasil search_stock adalah ACTUAL stock (open stock awal bulan dikurangi penjualan berjalan).
- Ini adalah stok fisik yang tersedia sekarang.
- Jangan sebut "ID30" ke user. Cukup bilang "stok" atau "stok tersedia".
- Field 'open_stock' dan 'sold' boleh disebutkan kalau user tanya breakdown-nya.

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
# AGENT LOOP — tidak diubah
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
        history = CHAT_HISTORY.get(user_id, [])
        contents = []
        for msg in history:
            role = "model" if msg["role"] == "assistant" else "user"
            contents.append(types.Content(role=role, parts=[types.Part(text=msg["content"])]))

        contents.append(types.Content(role="user", parts=[types.Part(text=user_input)]))

        config = types.GenerateContentConfig(
            system_instruction=get_system_prompt(),
            tools=GEMINI_TOOLS,
            max_output_tokens=2000,
        )

        for turn in range(MAX_TOOL_TURNS):
            response = gemini_client.models.generate_content(
                model=GEMINI_MODEL,
                contents=contents,
                config=config,
            )

            candidate = response.candidates[0]
            function_calls = [
                p.function_call for p in candidate.content.parts
                if p.function_call is not None
            ]

            if not function_calls:
                text_parts = [p.text for p in candidate.content.parts if p.text]
                reply = '\n'.join(text_parts).strip()
                break

            contents.append(candidate.content)

            tool_response_parts = []
            for fc in function_calls:
                args = dict(fc.args) if fc.args else {}
                result = call_tool(fc.name, args)
                tool_response_parts.append(
                    types.Part.from_function_response(name=fc.name, response=result)
                )

            contents.append(types.Content(role="user", parts=tool_response_parts))

        else:
            reply = '⚠️ AI butuh terlalu banyak langkah. Coba pertanyaan yang lebih spesifik.'

        if not reply:
            reply = '⚠️ AI tidak menghasilkan jawaban. Coba ulang.'

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
# TELEGRAM HANDLERS — tidak diubah
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
        '👋 *StockBot v5.2 — Actual Stock*\n\n'
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
        lines += [
            f'✅ *STOK:* `{DATA["stock_sheet"]}` — {len(df):,} produk',
            f'   Actual ada: {(df["Actual"]>0).sum():,} | Kosong: {(df["Actual"]==0).sum():,}',
            f'   Open stock: {int(df["Open_ID30"].sum()):,} | Terjual: {int(df["Sales_QTY"].sum()):,} | Actual: {int(df["Actual"].sum()):,}',
        ]
    else:
        lines.append('❌ *STOK:* belum ter-load')

    if DATA['sales_sheet']:
        lines.append(f'✅ *SALES:* `{DATA["sales_sheet"]}`')
    else:
        lines.append('⚠️ *SALES:* tab tidak ditemukan')

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
        '📖 *StockBot v5.2*\n\nTanya apa aja pakai bahasa natural. '
        'AI cari datanya sendiri.\n\nStok yang ditampilkan = *actual stock* '
        '(open stock awal bulan dikurangi penjualan berjalan).\n\n/reload /status /clear',
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

    logger.info("🚀 Starting StockBot v5.2 (Actual Stock Mode)...")
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
