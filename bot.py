"""
StockBot Telegram — Inventory Assistant
Deploy ke Render, baca data dari Google Drive
"""
import os
import re
import io
import logging
import pandas as pd
import requests
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

# ─────────────────────────────────────────────
# CONFIG — diisi via Environment Variables di Render
# ─────────────────────────────────────────────
TELEGRAM_TOKEN = os.environ.get('TELEGRAM_TOKEN', '')
GDRIVE_FILE_ID = os.environ.get('GDRIVE_FILE_ID', '')  # ID file Excel di Google Drive
ALLOWED_USER_IDS = os.environ.get('ALLOWED_USER_IDS', '')  # comma-separated, e.g. "123456,789012"

WAREHOUSE_COLS = ['ID30', 'ID40']

# Setup logging
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Global data cache
DATA = {'stock': None, 'filename': None, 'sheet': None, 'last_loaded': None}


# ─────────────────────────────────────────────
# GOOGLE DRIVE LOADER
# ─────────────────────────────────────────────

def download_from_gdrive(file_id):
    """Download file dari Google Drive (public/anyone with link)."""
    url = f'https://drive.google.com/uc?export=download&id={file_id}'
    response = requests.get(url, allow_redirects=True, timeout=60)

    # Handle large files yang butuh konfirmasi
    if 'confirm=' in response.text:
        match = re.search(r'confirm=([0-9A-Za-z_]+)', response.text)
        if match:
            url = f'https://drive.google.com/uc?export=download&confirm={match.group(1)}&id={file_id}'
            response = requests.get(url, allow_redirects=True, timeout=60)

    return io.BytesIO(response.content)


def load_data():
    """Load Excel dari Google Drive, parse sheet terbaru."""
    try:
        logger.info("📥 Downloading Excel from Google Drive...")
        excel_data = download_from_gdrive(GDRIVE_FILE_ID)

        xl = pd.ExcelFile(excel_data)
        sheets = xl.sheet_names
        if not sheets:
            return False, "File tidak punya sheet."

        latest_sheet = sheets[-1]
        df = pd.read_excel(excel_data, sheet_name=latest_sheet)

        # Normalize kolom pertama
        cols = list(df.columns)
        if str(cols[0]).strip() == '' or 'Unnamed' in str(cols[0]):
            df = df.rename(columns={cols[0]: 'Material'})

        # Pastikan kolom warehouse ada
        for col in WAREHOUSE_COLS:
            if col not in df.columns:
                df[col] = 0

        # Convert ke numeric
        for col in WAREHOUSE_COLS:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        df['Total'] = df[WAREHOUSE_COLS].sum(axis=1)
        df['desc_clean'] = df['Material Description'].astype(str).str.strip().str.upper()

        DATA['stock'] = df
        DATA['sheet'] = latest_sheet
        DATA['filename'] = 'GDrive File'

        logger.info(f"✅ Loaded {len(df)} products from sheet '{latest_sheet}'")
        return True, f"Data loaded: sheet '{latest_sheet}', {len(df):,} produk"

    except Exception as e:
        logger.error(f"❌ Load error: {e}")
        return False, f"Error loading data: {str(e)}"


# ─────────────────────────────────────────────
# QUERY ENGINE
# ─────────────────────────────────────────────

def extract_packaging(text):
    text = text.upper()
    match = re.search(r'(\d+(?:\.\d+)?)\s*L\b', text)
    if match:
        return match.group(0).replace(' ', '')
    return None


def answer_analytics(q, df):
    if any(kw in q for kw in ['PALING BANYAK', 'TERBESAR', 'TERBANYAK', 'TOP', 'HIGHEST']):
        n = 10
        match_n = re.search(r'TOP\s*(\d+)', q)
        if match_n:
            n = int(match_n.group(1))
        top = df[df['Total'] > 0].nlargest(n, 'Total')[['Material Description', 'ID30', 'ID40', 'Total']]
        lines = [f'🏆 *Top {n} Produk Stok Terbanyak:*\n']
        for i, (_, row) in enumerate(top.iterrows(), 1):
            desc = str(row['Material Description']).strip()[:50]
            lines.append(f'{i}. {desc}')
            lines.append(f'   ID30: {int(row["ID30"]):,} | ID40: {int(row["ID40"]):,} | Total: {int(row["Total"]):,}')
        return '\n'.join(lines)

    if any(kw in q for kw in ['PALING SEDIKIT', 'TERKECIL', 'TERENDAH', 'LOWEST']):
        low = df[df['Total'] > 0].nsmallest(10, 'Total')[['Material Description', 'ID30', 'ID40', 'Total']]
        lines = ['⚠️ *10 Produk Stok Paling Sedikit (>0):*\n']
        for i, (_, row) in enumerate(low.iterrows(), 1):
            desc = str(row['Material Description']).strip()[:50]
            lines.append(f'{i}. {desc}')
            lines.append(f'   ID30: {int(row["ID30"]):,} | ID40: {int(row["ID40"]):,} | Total: {int(row["Total"]):,}')
        return '\n'.join(lines)

    if any(kw in q for kw in ['KOSONG', 'HABIS', 'ZERO', 'NOL']):
        zero = df[df['Total'] == 0]
        return f'🚨 Produk dengan stok 0: *{len(zero):,}* dari total {len(df):,} produk.'

    if any(kw in q for kw in ['SUMMARY', 'RINGKASAN', 'TOTAL SEMUA', 'REKAP']):
        total_id30 = int(df['ID30'].sum())
        total_id40 = int(df['ID40'].sum())
        nonzero = len(df[df['Total'] > 0])
        zero = len(df[df['Total'] == 0])
        return (
            f'📊 *Ringkasan Stok:*\n\n'
            f'• Total produk: {len(df):,}\n'
            f'• Ada stok: {nonzero:,}\n'
            f'• Kosong: {zero:,}\n\n'
            f'🏭 ID30 Total: *{total_id30:,}*\n'
            f'🏭 ID40 Total: *{total_id40:,}*\n'
            f'✅ GRAND TOTAL: *{total_id30 + total_id40:,}*'
        )

    if 'ID30' in q or 'GUDANG 30' in q:
        top = df[df['ID30'] > 0].nlargest(10, 'ID30')[['Material Description', 'ID30']]
        lines = ['🏭 *Top 10 Stok di ID30:*\n']
        for i, (_, row) in enumerate(top.iterrows(), 1):
            lines.append(f'{i}. {str(row["Material Description"]).strip()[:50]} — {int(row["ID30"]):,}')
        return '\n'.join(lines)

    if 'ID40' in q or 'GUDANG 40' in q:
        top = df[df['ID40'] > 0].nlargest(10, 'ID40')[['Material Description', 'ID40']]
        lines = ['🏭 *Top 10 Stok di ID40:*\n']
        for i, (_, row) in enumerate(top.iterrows(), 1):
            lines.append(f'{i}. {str(row["Material Description"]).strip()[:50]} — {int(row["ID40"]):,}')
        return '\n'.join(lines)

    return (
        '🤖 Saya belum mengerti pertanyaan tersebut.\n\n'
        '*Contoh pertanyaan:*\n'
        '• stok titan truck plus 205L\n'
        '• produk stok paling banyak\n'
        '• 10 produk stok paling sedikit\n'
        '• berapa produk stok kosong\n'
        '• rekap total stok\n'
        '• top 10 stok di ID30'
    )


def query_stock(user_input):
    df = DATA['stock']
    if df is None:
        return '⚠️ Data belum ter-load. Coba ketik /reload'

    q = user_input.upper().strip()
    pkg = extract_packaging(q)

    stop_words = {'STOK', 'STOCK', 'ADA', 'BERAPA', 'DI', 'SEMUA', 'GUDANG',
                  'DAN', 'TOTALNYA', 'TOTAL', 'PRODUK', 'SEKARANG', 'SAAT', 'INI',
                  'JUMLAH', 'CEK', 'CARI', 'TUNJUKKAN', 'LIHAT'}

    q_no_pkg = re.sub(r'\d+(?:\.\d+)?\s*L\b', '', q).strip()
    words = [w for w in q_no_pkg.split() if w not in stop_words and len(w) > 1]

    if not words:
        return answer_analytics(q, df)

    mask = pd.Series([True] * len(df))
    for kw in words:
        mask = mask & df['desc_clean'].str.contains(kw, na=False, regex=False)

    if pkg:
        pkg_num = re.search(r'\d+', pkg).group()
        mask = mask & df['desc_clean'].str.contains(pkg_num + 'L', na=False, regex=False)

    result = df[mask].copy()

    if result.empty and len(words) > 2:
        mask2 = pd.Series([True] * len(df))
        for kw in words[:2]:
            mask2 = mask2 & df['desc_clean'].str.contains(kw, na=False, regex=False)
        result = df[mask2].copy()

    if result.empty:
        analytics_result = answer_analytics(q, df)
        if not analytics_result.startswith('🤖'):
            return analytics_result
        return f'❌ Produk tidak ditemukan untuk: "{user_input}"\n\nCoba kata kunci lebih spesifik.'

    lines = [f'📦 *Ditemukan {len(result)} varian:*\n']
    grand_id30 = 0
    grand_id40 = 0

    for _, row in result.iterrows():
        desc = str(row['Material Description']).strip()
        id30 = int(row['ID30'])
        id40 = int(row['ID40'])
        total = id30 + id40
        grand_id30 += id30
        grand_id40 += id40

        if total > 0:
            lines.append(f'• {desc}')
            lines.append(f'  ID30: {id30:,} | ID40: {id40:,} | Sub: {total:,}')
        else:
            lines.append(f'• {desc}  _(0)_')

    lines.append('\n─────────────────')
    lines.append(f'🏭 TOTAL ID30: *{grand_id30:,}*')
    lines.append(f'🏭 TOTAL ID40: *{grand_id40:,}*')
    lines.append(f'✅ GRAND TOTAL: *{grand_id30 + grand_id40:,}*')
    return '\n'.join(lines)


# ─────────────────────────────────────────────
# TELEGRAM HANDLERS
# ─────────────────────────────────────────────

def is_allowed(user_id):
    """Check apakah user_id diizinkan akses bot."""
    if not ALLOWED_USER_IDS:
        return True  # Kalau kosong, semua boleh
    allowed = [int(x.strip()) for x in ALLOWED_USER_IDS.split(',') if x.strip()]
    return user_id in allowed


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        await update.message.reply_text(f'❌ Akses ditolak.\nUser ID Anda: `{user_id}`', parse_mode='Markdown')
        return

    msg = (
        '👋 *Halo! Selamat datang di StockBot.*\n\n'
        'Saya bisa bantu cek stok inventory.\n\n'
        '*Commands:*\n'
        '/reload — refresh data dari Google Drive\n'
        '/status — info data yang ter-load\n'
        '/help — bantuan\n\n'
        '*Contoh pertanyaan:*\n'
        '• stok titan truck plus 205L\n'
        '• produk stok paling banyak\n'
        '• rekap total stok'
    )
    await update.message.reply_text(msg, parse_mode='Markdown')


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id):
        return
    msg = (
        '📖 *Bantuan StockBot:*\n\n'
        '*Cek produk spesifik:*\n'
        '`stok titan truck plus 205L`\n'
        '`titan cargo flex 20L`\n\n'
        '*Analytics:*\n'
        '`produk stok paling banyak`\n'
        '`top 20 stok terbanyak`\n'
        '`10 produk stok paling sedikit`\n'
        '`berapa produk stok kosong`\n'
        '`rekap total stok`\n\n'
        '*Per gudang:*\n'
        '`top 10 stok di ID30`\n'
        '`top 10 stok di ID40`\n\n'
        '*Command:*\n'
        '/reload — update data Excel dari Drive\n'
        '/status — lihat info data'
    )
    await update.message.reply_text(msg, parse_mode='Markdown')


async def reload_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id):
        return
    await update.message.reply_text('🔄 Memuat ulang data dari Google Drive...')
    success, message = load_data()
    if success:
        await update.message.reply_text(f'✅ {message}')
    else:
        await update.message.reply_text(f'❌ {message}')


async def status_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id):
        return
    if DATA['stock'] is None:
        await update.message.reply_text('⚠️ Data belum ter-load. Ketik /reload')
    else:
        df = DATA['stock']
        msg = (
            f'📊 *Status Data:*\n\n'
            f'• Sheet: `{DATA["sheet"]}`\n'
            f'• Total produk: {len(df):,}\n'
            f'• Ada stok: {len(df[df["Total"] > 0]):,}\n'
            f'• Kosong: {len(df[df["Total"] == 0]):,}'
        )
        await update.message.reply_text(msg, parse_mode='Markdown')


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        await update.message.reply_text(f'❌ Akses ditolak.\nUser ID Anda: `{user_id}`', parse_mode='Markdown')
        return

    user_text = update.message.text
    logger.info(f"Query from {user_id}: {user_text}")

    reply = query_stock(user_text)

    # Split message kalau terlalu panjang (Telegram limit 4096 chars)
    if len(reply) > 4000:
        chunks = [reply[i:i+4000] for i in range(0, len(reply), 4000)]
        for chunk in chunks:
            await update.message.reply_text(chunk, parse_mode='Markdown')
    else:
        try:
            await update.message.reply_text(reply, parse_mode='Markdown')
        except Exception:
            # Fallback tanpa markdown kalau ada karakter aneh
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

    # Load data saat bot start
    logger.info("🚀 Starting StockBot...")
    success, msg = load_data()
    logger.info(f"Initial load: {msg}")

    # Build app
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    # Handlers
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("reload", reload_cmd))
    app.add_handler(CommandHandler("status", status_cmd))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    logger.info("✅ Bot is running...")
    app.run_polling(drop_pending_updates=True)


if __name__ == '__main__':
    main()
