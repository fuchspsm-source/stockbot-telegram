"""
StockBot Telegram v4.1 — Full AI Assistant
Powered by Gemini Flash. All queries handled by AI.
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
import google.generativeai as genai

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
TELEGRAM_TOKEN = os.environ.get('TELEGRAM_TOKEN', '')
GDRIVE_FILE_ID = os.environ.get('GDRIVE_FILE_ID', '')
GDRIVE_COGS_FILE_ID = os.environ.get('GDRIVE_COGS_FILE_ID', '')
ALLOWED_USER_IDS = os.environ.get('ALLOWED_USER_IDS', '')
GEMINI_API_KEY = os.environ.get('GEMINI_API_KEY', '')

WAREHOUSE_COLS = ['ID30', 'ID40']
GEMINI_MODEL = 'gemini-2.5-flash-lite'
MAX_TOKENS = 1500
MAX_CONTEXT_ROWS = 50  # Max data rows kasih ke AI

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

gemini_model = None
if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)
    gemini_model = genai.GenerativeModel(GEMINI_MODEL)
    logger.info("✅ Gemini AI client initialized")
else:
    logger.warning("⚠️ GEMINI_API_KEY belum di-set!")

DATA = {
    'stock': None,
    'stock_sheet': None,
    'cogs': None,
    'cogs_sheet': None,
}

# Chat history per user (max 6 turns)
CHAT_HISTORY = {}
MAX_HISTORY = 6


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
            if 'CODE' in str(col).upper() or 'MATERIAL' == str(col).upper().strip():
                code_col = col
                break

        for col in WAREHOUSE_COLS:
            if col not in df.columns:
                df[col] = 0
        for col in WAREHOUSE_COLS:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        df['Total'] = df[WAREHOUSE_COLS].sum(axis=1)
        df['desc_clean'] = df['Material Description'].astype(str).str.strip().str.upper()

        if code_col:
            df['code_clean'] = df[code_col].astype(str).str.strip().str.upper()
        else:
            df['code_clean'] = ''

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
        if not sheets:
            return False, "File COGS tidak punya sheet."
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
        df['source_clean'] = df['Source'].astype(str).str.strip().str.upper()
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
    messages = [f"{'✅' if stock_ok else '❌'} {stock_msg}",
                f"{'✅' if cogs_ok else '⚠️'} {cogs_msg}"]
    return stock_ok, '\n'.join(messages)


# ─────────────────────────────────────────────
# SMART DATA FETCHER (untuk AI context)
# ─────────────────────────────────────────────
COMMON_STOP_WORDS = {
    'STOK', 'STOCK', 'COGS', 'COST', 'HARGA', 'BERAPA', 'YA', 'DONG',
    'PRODUK', 'BARANG', 'UNTUK', 'DARI', 'KEMASAN', 'YANG', 'DI',
    'DAN', 'SEMUA', 'ADA', 'CEK', 'CARI', 'LIHAT', 'TUNJUKKAN',
    'MODAL', 'HPP', 'JUAL', 'NC', 'BAGAIMANA', 'GIMANA', 'APA',
    'SAJA', 'MANA', 'LU', 'GW', 'GUE', 'KAMU', 'SAYA', 'AKU',
    'KAYANYA', 'SALAH', 'DEH', 'COBA', 'TOLONG', 'BISA', 'MAU',
    'TIDAK', 'GAK', 'NGGAK', 'KOK', 'KAN', 'AJA', 'SIH',
    'INI', 'ITU', 'KE', 'PADA', 'OLEH', 'JUGA', 'TAPI',
    'NAMUN', 'JADI', 'ATAU', 'KALAU', 'KALO', 'MISAL', 'CONTOH',
    'BANYAK', 'SEDIKIT', 'BAGUS', 'JELEK', 'BAIK', 'BURUK',
    'JIKA', 'BILA', 'AGAR', 'SUPAYA', 'KARENA', 'SEBAB',
    'BIAR', 'SAMA', 'TERIMA', 'KASIH', 'HALO', 'HAI', 'OK', 'OKE',
}


def extract_keywords(user_input):
    """Extract keyword penting dari user input."""
    text = user_input.upper()
    codes = re.findall(r'\b\d{6,}\b', text)
    words = [w for w in text.split() if w not in COMMON_STOP_WORDS and len(w) > 2 and not w.isdigit()]
    words = sorted(set(words), key=lambda w: -len(w))
    return codes, words


def fetch_relevant_data(user_input):
    """Smart fetch data relevan dengan query user."""
    codes, words = extract_keywords(user_input)

    stock_data = []
    cogs_data = []
    stock_found_idx = set()
    cogs_found_codes = set()

    # ─── STOCK ───
    if DATA['stock'] is not None:
        df = DATA['stock']

        if codes and 'code_clean' in df.columns:
            for code in codes:
                code_mask = df['code_clean'] == code.upper()
                for idx, row in df[code_mask].iterrows():
                    if idx in stock_found_idx:
                        continue
                    stock_found_idx.add(idx)
                    stock_data.append({
                        'desc': str(row['Material Description']).strip(),
                        'code': str(row.get('code_clean', '')),
                        'ID30': int(row['ID30']),
                        'ID40': int(row['ID40']),
                        'Total': int(row['Total'])
                    })

        for kw in words[:5]:
            if len(stock_data) >= MAX_CONTEXT_ROWS:
                break
            kw_mask = df['desc_clean'].str.contains(kw, na=False, regex=False)
            for idx, row in df[kw_mask].head(15).iterrows():
                if idx in stock_found_idx:
                    continue
                stock_found_idx.add(idx)
                stock_data.append({
                    'desc': str(row['Material Description']).strip(),
                    'code': str(row.get('code_clean', '')),
                    'ID30': int(row['ID30']),
                    'ID40': int(row['ID40']),
                    'Total': int(row['Total'])
                })
                if len(stock_data) >= MAX_CONTEXT_ROWS:
                    break

    # ─── COGS ───
    if DATA['cogs'] is not None:
        df = DATA['cogs']

        if codes:
            for code in codes:
                code_mask = df['code_clean'] == code.upper()
                for _, row in df[code_mask].iterrows():
                    if row['Material Code'] in cogs_found_codes:
                        continue
                    cogs_found_codes.add(row['Material Code'])
                    nc_prev = str(row.get('Rata2 NC Sebelumnya', '')).strip()
                    cogs_data.append({
                        'code': row['Material Code'],
                        'desc': str(row['Material Description']).strip(),
                        'source': str(row['Source']).strip(),
                        'cogs': int(row['COGS']),
                        'update': str(row['Update']).strip(),
                        'rata2_nc_sebelumnya': nc_prev if nc_prev != 'nan' else ''
                    })

        for kw in words[:5]:
            if len(cogs_data) >= MAX_CONTEXT_ROWS:
                break
            kw_mask = df['desc_clean'].str.contains(kw, na=False, regex=False)
            for _, row in df[kw_mask].head(15).iterrows():
                if row['Material Code'] in cogs_found_codes:
                    continue
                cogs_found_codes.add(row['Material Code'])
                nc_prev = str(row.get('Rata2 NC Sebelumnya', '')).strip()
                cogs_data.append({
                    'code': row['Material Code'],
                    'desc': str(row['Material Description']).strip(),
                    'source': str(row['Source']).strip(),
                    'cogs': int(row['COGS']),
                    'update': str(row['Update']).strip(),
                    'rata2_nc_sebelumnya': nc_prev if nc_prev != 'nan' else ''
                })
                if len(cogs_data) >= MAX_CONTEXT_ROWS:
                    break

    return stock_data, cogs_data


# ─────────────────────────────────────────────
# GEMINI AI HANDLER
# ─────────────────────────────────────────────
def get_data_summary():
    """Summary singkat tentang data yang ada."""
    parts = []
    if DATA['stock'] is not None:
        df = DATA['stock']
        parts.append(f"STOK: {len(df):,} produk (sheet: {DATA['stock_sheet']}). Warehouse: ID30, ID40.")
    if DATA['cogs'] is not None:
        df = DATA['cogs']
        sources = df['Source'].unique().tolist()
        parts.append(f"COGS: {len(df):,} entries (sheet: {DATA['cogs_sheet']}). Sources: {', '.join(sources)}.")
    return ' | '.join(parts)


def get_analytics_data():
    """Pre-compute analytics untuk question analytics."""
    if DATA['stock'] is None:
        return {}
    df = DATA['stock']
    return {
        'total_produk': len(df),
        'stok_kosong': int(len(df[df['Total'] == 0])),
        'ada_stok': int(len(df[df['Total'] > 0])),
        'total_id30': int(df['ID30'].sum()),
        'total_id40': int(df['ID40'].sum()),
        'grand_total': int(df['Total'].sum()),
        'top_10_terbanyak': [
            {'desc': str(row['Material Description']).strip(), 'total': int(row['Total']),
             'id30': int(row['ID30']), 'id40': int(row['ID40'])}
            for _, row in df[df['Total'] > 0].nlargest(10, 'Total').iterrows()
        ],
        'top_10_tersedikit': [
            {'desc': str(row['Material Description']).strip(), 'total': int(row['Total']),
             'id30': int(row['ID30']), 'id40': int(row['ID40'])}
            for _, row in df[df['Total'] > 0].nsmallest(10, 'Total').iterrows()
        ],
    }


def build_system_prompt(data_context):
    return f"""Kamu adalah StockBot, asisten AI untuk perusahaan pelumas. Bantu user dengan inventory, COGS, harga jual, margin (NC), dan analisa bisnis.

ATURAN:
1. Jawab pakai Bahasa Indonesia kasual tapi profesional.
2. JANGAN mengarang data — hanya berdasarkan data yang dikasih.
3. Kalau data tidak match query user, bilang apa adanya & kasih saran cek ejaan/code.
4. Format angka: pakai pemisah ribuan, contoh Rp 50.000.
5. Output ringkas (max 250 kata). Pakai bullet point + emoji secukupnya.
6. Pakai Markdown Telegram: *bold*, `code`.

FORMULA:
- NC% = (Harga Jual - COGS) / Harga Jual × 100
- Harga Jual = COGS / (1 - NC%/100)

PENTING SOAL MATERIAL CODE:
- Material code di STOK dan COGS BISA BEDA untuk produk yang sama (karena beda source/supplier).
- Kalau user kasih kode stok tapi data COGS pakai kode lain, cari COGS by deskripsi produk.
- 1 produk fisik bisa punya multi-source dengan beda kode (Local, Jerman, China, dll).

CARA KERJA DATA:
- Kolom STOK: Material Description, code, ID30, ID40, Total.
- Kolom COGS: code, desc, source, cogs (Rupiah), update = label COGS update terakhir (bukan margin update), rata2_nc_sebelumnya.
- Sheet COGS paling kanan = paling baru.

DATA KONTEKS:
{data_context}"""


def ask_gemini(user_input, user_id):
    """Send query to Gemini with smart context."""
    if not gemini_model:
        return '⚠️ Fitur AI belum aktif. Hubungi admin untuk set GEMINI_API_KEY.'

    try:
        # Fetch relevant data
        stock_data, cogs_data = fetch_relevant_data(user_input)

        wants_analytics = any(kw in user_input.upper() for kw in [
            'PALING', 'TOP', 'REKAP', 'RINGKASAN', 'TOTAL SEMUA',
            'KOSONG', 'HABIS', 'TERBANYAK', 'TERSEDIKIT'
        ])

        analytics = get_analytics_data() if wants_analytics else None
        data_summary = get_data_summary()

        # Build context
        context_parts = [f"DATA SUMMARY: {data_summary}"]
        if analytics:
            context_parts.append(f"\nANALYTICS:\n{json.dumps(analytics, indent=2, ensure_ascii=False)}")
        if stock_data:
            context_parts.append(f"\nSTOK RELEVAN ({len(stock_data)} hasil):\n{json.dumps(stock_data, indent=2, ensure_ascii=False)}")
        if cogs_data:
            context_parts.append(f"\nCOGS RELEVAN ({len(cogs_data)} hasil):\n{json.dumps(cogs_data, indent=2, ensure_ascii=False)}")
        if not stock_data and not cogs_data and not analytics:
            context_parts.append("\n(Tidak ada data spesifik yang match. User mungkin tanya hal umum.)")

        data_context = '\n'.join(context_parts)
        system_prompt = build_system_prompt(data_context)

        # Build chat history untuk Gemini format
        # Gemini pakai format: [{"role": "user"/"model", "parts": [{"text": "..."}]}]
        history = CHAT_HISTORY.get(user_id, [])
        gemini_history = []
        for msg in history:
            role = "model" if msg["role"] == "assistant" else "user"
            gemini_history.append({"role": role, "parts": [{"text": msg["content"]}]})

        # Start chat session dengan history
        chat = gemini_model.start_chat(history=gemini_history)

        # Gabungkan system prompt + user input (Gemini Flash tidak punya system role terpisah)
        full_prompt = f"{system_prompt}\n\n---\nPertanyaan user: {user_input}"

        response = chat.send_message(
            full_prompt,
            generation_config=genai.GenerationConfig(max_output_tokens=MAX_TOKENS)
        )
        reply = response.text

        # Save history (simpan format asli, bukan Gemini format)
        history.append({"role": "user", "content": user_input})
        history.append({"role": "assistant", "content": reply})
        if len(history) > MAX_HISTORY * 2:
            history = history[-(MAX_HISTORY * 2):]
        CHAT_HISTORY[user_id] = history

        logger.info(f"🤖 Reply ke {user_id}: {len(reply)} chars (stock:{len(stock_data)}, cogs:{len(cogs_data)})")
        return reply

    except Exception as e:
        logger.error(f"❌ Gemini error: {e}")
        return f'⚠️ AI error: {str(e)[:150]}\n\nCoba ulangi pertanyaannya atau ketik /reload.'


def clear_history(user_id):
    if user_id in CHAT_HISTORY:
        del CHAT_HISTORY[user_id]
        return '🔄 Chat history di-reset. Mulai dari awal lagi!'
    return '✅ Tidak ada history aktif.'


# ─────────────────────────────────────────────
# TELEGRAM HANDLERS
# ─────────────────────────────────────────────
def is_allowed(user_id):
    if not ALLOWED_USER_IDS:
        return True
    allowed = [int(x.strip()) for x in ALLOWED_USER_IDS.split(',') if x.strip()]
    return user_id in allowed


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        await update.message.reply_text(f'❌ Akses ditolak.\nUser ID: `{user_id}`', parse_mode='Markdown')
        return
    ai_status = '🤖 Full AI Mode' if gemini_model else '⚠️ AI off'
    msg = (
        '👋 *Halo! StockBot v4.1 (Full AI)*\n\n'
        f'{ai_status}\n\n'
        'Saya dijalankan oleh Gemini Flash.\n'
        'Tanya saya apa aja — natural language, casual, complex — semua OK!\n\n'
        '*Contoh:*\n'
        '• Cek stok ceplattyn\n'
        '• Berapa cost titan plus 205L dari local?\n'
        '• Hitung harga jual kalau NC 30%\n'
        '• Bandingkan margin titan vs ceplattyn\n'
        '• Saran restock produk apa minggu depan?\n\n'
        '*Commands:*\n'
        '/reload — refresh data\n'
        '/status — info data\n'
        '/clear — reset chat history'
    )
    await update.message.reply_text(msg, parse_mode='Markdown')


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id):
        return
    msg = (
        '📖 *StockBot v4.1 — Bantuan*\n\n'
        '🤖 Bot ini powered by Gemini AI. Tanya apa aja!\n\n'
        '*Yang bisa ditanya:*\n'
        '• Cek stok produk\n'
        '• Cek COGS (semua source)\n'
        '• Hitung NC / harga jual\n'
        '• Analisa & rekomendasi\n'
        '• Bandingkan produk\n\n'
        '*Tips:*\n'
        '• Bot ingat percakapan sebelumnya\n'
        '• Ketik `/clear` untuk reset history\n'
        '• Material code di STOK & COGS bisa beda, bot otomatis match by nama\n\n'
        '*Commands:*\n'
        '/reload — refresh data dari Google Drive\n'
        '/status — info data ter-load\n'
        '/clear — reset chat history'
    )
    await update.message.reply_text(msg, parse_mode='Markdown')


async def reload_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id):
        return
    await update.message.reply_text('🔄 Memuat ulang data dari Google Drive...')
    success, message = load_all_data()
    icon = '✅' if success else '❌'
    await update.message.reply_text(f'{icon} Hasil reload:\n\n{message}')


async def status_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id):
        return
    lines = ['📊 *Status Data:*\n']
    if DATA['stock'] is None:
        lines.append('❌ *STOK:* belum ter-load')
    else:
        df = DATA['stock']
        lines.append(f'✅ *STOK:*')
        lines.append(f'   • Sheet: `{DATA["stock_sheet"]}`')
        lines.append(f'   • Total produk: {len(df):,}')
        lines.append(f'   • Ada stok: {len(df[df["Total"] > 0]):,}')
        lines.append(f'   • Kosong: {len(df[df["Total"] == 0]):,}')
    lines.append('')
    if DATA['cogs'] is None:
        lines.append('⚠️ *COGS:* belum ter-load')
    else:
        df = DATA['cogs']
        sources = df['Source'].nunique()
        lines.append(f'✅ *COGS:*')
        lines.append(f'   • Sheet: `{DATA["cogs_sheet"]}`')
        lines.append(f'   • Total entries: {len(df):,}')
        lines.append(f'   • Jumlah source: {sources}')
    lines.append('')
    if gemini_model:
        lines.append('🤖 *AI:* Gemini Flash (FULL MODE)')
    else:
        lines.append('⚠️ *AI:* off')
    user_id = update.effective_user.id
    if user_id in CHAT_HISTORY:
        lines.append(f'\n💬 *Chat history:* {len(CHAT_HISTORY[user_id])//2} turn')
    await update.message.reply_text('\n'.join(lines), parse_mode='Markdown')


async def clear_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id):
        return
    user_id = update.effective_user.id
    reply = clear_history(user_id)
    await update.message.reply_text(reply)


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        await update.message.reply_text(f'❌ Akses ditolak.\nUser ID: `{user_id}`', parse_mode='Markdown')
        return
    user_text = update.message.text
    logger.info(f"Query from {user_id}: {user_text}")

    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action='typing')

    if user_text.upper().strip() in ['STOP', 'CLEAR', 'RESET']:
        reply = clear_history(user_id)
        await update.message.reply_text(reply)
        return

    reply = ask_gemini(user_text, user_id)

    if len(reply) > 4000:
        chunks = [reply[i:i+4000] for i in range(0, len(reply), 4000)]
        for chunk in chunks:
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
    if not GDRIVE_COGS_FILE_ID:
        logger.warning("⚠️ GDRIVE_COGS_FILE_ID belum di-set.")
    if not GEMINI_API_KEY:
        logger.error("❌ GEMINI_API_KEY tidak di-set! Bot v4.1 butuh AI.")
        return

    logger.info("🚀 Starting StockBot v4.1 (Gemini Flash AI Mode)...")
    success, msg = load_all_data()
    logger.info(f"Initial load:\n{msg}")

    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("reload", reload_cmd))
    app.add_handler(CommandHandler("status", status_cmd))
    app.add_handler(CommandHandler("clear", clear_cmd))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    logger.info("✅ Bot running with Gemini Flash AI...")
    app.run_polling(drop_pending_updates=True)


if __name__ == '__main__':
    main()
