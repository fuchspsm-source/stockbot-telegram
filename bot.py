"""
StockBot Telegram — Inventory & COGS Assistant
Deploy ke Railway, baca data dari Google Drive
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
# CONFIG — diisi via Environment Variables di Railway
# ─────────────────────────────────────────────
TELEGRAM_TOKEN = os.environ.get('TELEGRAM_TOKEN', '')
GDRIVE_FILE_ID = os.environ.get('GDRIVE_FILE_ID', '')             # File Excel STOK
GDRIVE_COGS_FILE_ID = os.environ.get('GDRIVE_COGS_FILE_ID', '')   # File Excel COGS — NEW
ALLOWED_USER_IDS = os.environ.get('ALLOWED_USER_IDS', '')

WAREHOUSE_COLS = ['ID30', 'ID40']

# Setup logging
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Global data cache
DATA = {
    'stock': None,
    'stock_sheet': None,
    'cogs': None,
    'cogs_sheet': None,
    'filename': None,
}

# Session per user untuk context COGS
SESSIONS = {}

# Keywords untuk routing
COGS_KEYWORDS = {'COGS', 'COST', 'HARGA MODAL', 'MODAL', 'HPP'}
CLEAR_KEYWORDS = {'STOP', 'CLEAR', 'RESET'}
NC_PATTERN = re.compile(r'NC\s*(\d+(?:\.\d+)?)\s*%?', re.IGNORECASE)
PRICE_PATTERN = re.compile(r'JUAL\s*(?:RP\s*)?([\d.,]+)', re.IGNORECASE)


# ─────────────────────────────────────────────
# GOOGLE DRIVE LOADER
# ─────────────────────────────────────────────
def download_from_gdrive(file_id):
    """Download file dari Google Drive (public/anyone with link)."""
    url = f'https://drive.google.com/uc?export=download&id={file_id}'
    response = requests.get(url, allow_redirects=True, timeout=60)
    
    if 'confirm=' in response.text:
        match = re.search(r'confirm=([0-9A-Za-z_]+)', response.text)
        if match:
            url = f'https://drive.google.com/uc?export=download&confirm={match.group(1)}&id={file_id}'
            response = requests.get(url, allow_redirects=True, timeout=60)
    
    return io.BytesIO(response.content)


def load_stock_data():
    """Load Excel STOK dari Google Drive, parse sheet terbaru."""
    try:
        logger.info("📥 Downloading STOCK Excel from Google Drive...")
        excel_data = download_from_gdrive(GDRIVE_FILE_ID)
        xl = pd.ExcelFile(excel_data)
        sheets = xl.sheet_names
        
        if not sheets:
            return False, "File stok tidak punya sheet."
        
        latest_sheet = sheets[-1]
        df = pd.read_excel(excel_data, sheet_name=latest_sheet)
        
        cols = list(df.columns)
        if str(cols[0]).strip() == '' or 'Unnamed' in str(cols[0]):
            df = df.rename(columns={cols[0]: 'Material'})
        
        for col in WAREHOUSE_COLS:
            if col not in df.columns:
                df[col] = 0
        
        for col in WAREHOUSE_COLS:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        df['Total'] = df[WAREHOUSE_COLS].sum(axis=1)
        df['desc_clean'] = df['Material Description'].astype(str).str.strip().str.upper()
        
        DATA['stock'] = df
        DATA['stock_sheet'] = latest_sheet
        
        logger.info(f"✅ Loaded STOCK: {len(df)} products from sheet '{latest_sheet}'")
        return True, f"Stok: sheet '{latest_sheet}', {len(df):,} produk"
    
    except Exception as e:
        logger.error(f"❌ Stock load error: {e}")
        return False, f"Error loading stock: {str(e)}"


def load_cogs_data():
    """Load Excel COGS dari Google Drive, parse sheet terbaru."""
    if not GDRIVE_COGS_FILE_ID:
        logger.warning("⚠️ GDRIVE_COGS_FILE_ID belum di-set, skip COGS load.")
        return False, "COGS file ID belum di-set di environment."
    
    try:
        logger.info("📥 Downloading COGS Excel from Google Drive...")
        excel_data = download_from_gdrive(GDRIVE_COGS_FILE_ID)
        xl = pd.ExcelFile(excel_data)
        sheets = xl.sheet_names
        
        if not sheets:
            return False, "File COGS tidak punya sheet."
        
        latest_sheet = sheets[-1]
        df = pd.read_excel(excel_data, sheet_name=latest_sheet)
        
        required = ['Material Code', 'Material Description', 'Source', 'COGS', 'Update']
        missing = [c for c in required if c not in df.columns]
        if missing:
            return False, f"Kolom COGS tidak lengkap. Missing: {missing}"
        
        df['COGS'] = pd.to_numeric(df['COGS'], errors='coerce').fillna(0)
        df['Material Code'] = df['Material Code'].astype(str).str.strip()
        
        df['desc_clean'] = df['Material Description'].astype(str).str.strip().str.upper()
        df['code_clean'] = df['Material Code'].str.upper()
        df['source_clean'] = df['Source'].astype(str).str.strip().str.upper()
        
        df = df[df['COGS'] > 0].reset_index(drop=True)
        
        DATA['cogs'] = df
        DATA['cogs_sheet'] = latest_sheet
        
        logger.info(f"✅ Loaded COGS: {len(df)} entries from sheet '{latest_sheet}'")
        return True, f"COGS: sheet '{latest_sheet}', {len(df):,} entries"
    
    except Exception as e:
        logger.error(f"❌ COGS load error: {e}")
        return False, f"Error loading COGS: {str(e)}"


def load_all_data():
    """Load stok + COGS sekaligus. Return summary message."""
    stock_ok, stock_msg = load_stock_data()
    cogs_ok, cogs_msg = load_cogs_data()
    
    messages = []
    messages.append(f"{'✅' if stock_ok else '❌'} {stock_msg}")
    messages.append(f"{'✅' if cogs_ok else '⚠️'} {cogs_msg}")
    
    return stock_ok, '\n'.join(messages)


# ─────────────────────────────────────────────
# QUERY ENGINE — STOCK (tidak berubah)
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
        '• top 10 stok di ID30\n'
        '• cogs ceplattyn sf 30'
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
            lines.append(f'   ID30: {id30:,} | ID40: {id40:,} | Sub: {total:,}')
        else:
            lines.append(f'• {desc} _(0)_')
    
    lines.append('\n─────────────────')
    lines.append(f'🏭 TOTAL ID30: *{grand_id30:,}*')
    lines.append(f'🏭 TOTAL ID40: *{grand_id40:,}*')
    lines.append(f'✅ GRAND TOTAL: *{grand_id30 + grand_id40:,}*')
    
    return '\n'.join(lines)


# ─────────────────────────────────────────────
# QUERY ENGINE — COGS
# ─────────────────────────────────────────────
def extract_cogs_query(user_input):
    """Bersihkan input dari keyword COGS, ambil nama produk saja."""
    text = user_input.upper().strip()
    for kw in COGS_KEYWORDS:
        text = text.replace(kw, ' ')
    stop_words = {'BERAPA', 'YA', 'DONG', 'PRODUK', 'BARANG', 'UNTUK'}
    words = [w for w in text.split() if w not in stop_words and len(w) > 0]
    return ' '.join(words).strip()


def query_cogs(user_input, user_id):
    """Cari produk di data COGS, tampilkan semua source."""
    df = DATA['cogs']
    
    if df is None or len(df) == 0:
        return '⚠️ Data COGS belum ter-load. Coba ketik /reload'
    
    query = extract_cogs_query(user_input)
    
    if not query:
        return '❌ Tolong sebutkan nama produk atau material code.\n\nContoh: `cogs ceplattyn sf 30` atau `cogs 10234`'
    
    # Tier 1: Exact match Material Code
    code_mask = df['code_clean'] == query
    if code_mask.any():
        result = df[code_mask].copy()
    else:
        # Tier 2: Match deskripsi (semua keyword)
        words = [w for w in query.split() if len(w) > 1]
        if not words:
            return f'❌ Query terlalu pendek: "{user_input}"'
        
        mask = pd.Series([True] * len(df))
        for kw in words:
            mask = mask & df['desc_clean'].str.contains(kw, na=False, regex=False)
        result = df[mask].copy()
        
        # Tier 3: Fallback 2 kata pertama
        if result.empty and len(words) > 2:
            mask2 = pd.Series([True] * len(df))
            for kw in words[:2]:
                mask2 = mask2 & df['desc_clean'].str.contains(kw, na=False, regex=False)
            result = df[mask2].copy()
    
    if result.empty:
        return (
            f'❌ Produk COGS tidak ditemukan untuk: "{user_input}"\n\n'
            f'Coba kata kunci lebih spesifik atau cek Material Code.'
        )
    
    # Simpan ke session
    SESSIONS[user_id] = {
        'product_query': query,
        'matches': result.copy()
    }
    
    # Format output
    lines = [f'💰 *COGS — {len(result)} varian ditemukan:*\n']
    
    for _, row in result.iterrows():
        code = row['Material Code']
        desc = str(row['Material Description']).strip()
        source = str(row['Source']).strip()
        cogs = int(row['COGS'])
        update = str(row['Update']).strip()
        
        lines.append(f'• `[{code}]` {desc}')
        lines.append(f'   Source: *{source}* | COGS: *Rp {cogs:,}*')
        lines.append(f'   📅 {update}\n')
    
    lines.append('─────────────────')
    lines.append('💡 Lanjut hitung margin? Ketik misal:')
    lines.append('   `china nc 30%` atau `local jual 100000`')
    lines.append('   Ketik `stop` untuk reset.')
    
    return '\n'.join(lines)


# ─────────────────────────────────────────────
# CALCULATOR — NC / Harga Jual
# ─────────────────────────────────────────────
def parse_number(text):
    """Parse angka dari string '100000', '100.000', '1,000,000', dst."""
    cleaned = re.sub(r'[^\d]', '', text)
    return int(cleaned) if cleaned else 0


def calculate_margin(user_input, user_id):
    """Kalkulasi NC atau Harga Jual berdasarkan session aktif."""
    session = SESSIONS.get(user_id)
    
    if not session:
        return None
    
    matches = session['matches']
    query_upper = user_input.upper()
    
    available_sources = matches['source_clean'].unique().tolist()
    
    target_source = None
    for src in available_sources:
        if src in query_upper:
            target_source = src
            break
    
    if not target_source:
        if len(available_sources) == 1:
            target_source = available_sources[0]
        else:
            sources_list = ', '.join([s.title() for s in available_sources])
            return (
                f'⚠️ Tolong sebutkan source-nya.\n'
                f'Pilihan: *{sources_list}*\n\n'
                f'Contoh: `china nc 30%` atau `local jual 100000`'
            )
    
    row = matches[matches['source_clean'] == target_source].iloc[0]
    cogs = int(row['COGS'])
    desc = str(row['Material Description']).strip()
    code = row['Material Code']
    update = str(row['Update']).strip()
    
    nc_match = NC_PATTERN.search(user_input)
    price_match = PRICE_PATTERN.search(user_input)
    
    if nc_match:
        nc_percent = float(nc_match.group(1))
        if nc_percent >= 100:
            return f'❌ NC harus < 100%. Anda input: {nc_percent}%'
        
        nc_decimal = nc_percent / 100
        harga_jual = cogs / (1 - nc_decimal)
        
        return (
            f'💰 *{desc}*\n'
            f'`[{code}]` Source: *{target_source.title()}*\n\n'
            f'• COGS:        Rp {cogs:,}\n'
            f'• NC Target:   *{nc_percent}%*\n'
            f'• Harga Jual:  *Rp {int(harga_jual):,}*\n\n'
            f'📐 Formula: HJ = COGS / (1 - NC%)\n'
            f'📅 COGS update: {update}'
        )
    
    elif price_match:
        harga_jual = parse_number(price_match.group(1))
        if harga_jual <= 0:
            return f'❌ Harga jual harus > 0.'
        
        nc_value = (harga_jual - cogs) / harga_jual * 100
        
        if harga_jual <= cogs:
            return (
                f'⚠️ *Harga jual di bawah/sama dengan COGS — RUGI!*\n\n'
                f'• {desc} ({target_source.title()})\n'
                f'• COGS:       Rp {cogs:,}\n'
                f'• Harga Jual: Rp {harga_jual:,}\n'
                f'• NC:         *{nc_value:.1f}%* (negatif/nol)\n'
            )
        
        return (
            f'💰 *{desc}*\n'
            f'`[{code}]` Source: *{target_source.title()}*\n\n'
            f'• COGS:        Rp {cogs:,}\n'
            f'• Harga Jual:  Rp {harga_jual:,}\n'
            f'• NC:          *{nc_value:.1f}%*\n\n'
            f'📐 Formula: NC = (HJ - COGS) / HJ\n'
            f'📅 COGS update: {update}'
        )
    
    else:
        return None


def clear_session(user_id):
    """Hapus session user."""
    if user_id in SESSIONS:
        del SESSIONS[user_id]
        return '🔄 Session di-reset. Silakan tanya produk baru.'
    return '✅ Tidak ada session aktif.'


# ─────────────────────────────────────────────
# ROUTER — Tentukan handler mana
# ─────────────────────────────────────────────
def route_message(user_input, user_id):
    """Router utama: tentukan handler berdasarkan input."""
    text_upper = user_input.upper().strip()
    
    # Prioritas 1: Clear session
    if text_upper in CLEAR_KEYWORDS:
        return clear_session(user_id)
    
    # Prioritas 2: Session aktif + kalkulasi
    if user_id in SESSIONS:
        has_nc = bool(NC_PATTERN.search(user_input))
        has_price = bool(PRICE_PATTERN.search(user_input))
        
        if has_nc or has_price:
            result = calculate_margin(user_input, user_id)
            if result:
                return result
    
    # Prioritas 3: COGS query
    if any(kw in text_upper for kw in COGS_KEYWORDS):
        return query_cogs(user_input, user_id)
    
    # Prioritas 4: Default ke query stok
    return query_stock(user_input)


# ─────────────────────────────────────────────
# TELEGRAM HANDLERS
# ─────────────────────────────────────────────
def is_allowed(user_id):
    """Check apakah user_id diizinkan akses bot."""
    if not ALLOWED_USER_IDS:
        return True
    allowed = [int(x.strip()) for x in ALLOWED_USER_IDS.split(',') if x.strip()]
    return user_id in allowed


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        await update.message.reply_text(f'❌ Akses ditolak.\nUser ID Anda: `{user_id}`', parse_mode='Markdown')
        return
    
    msg = (
        '👋 *Halo! Selamat datang di StockBot.*\n\n'
        'Saya bisa bantu cek stok inventory & COGS.\n\n'
        '*Commands:*\n'
        '/reload — refresh data dari Google Drive\n'
        '/status — info data yang ter-load\n'
        '/help — bantuan\n\n'
        '*Contoh pertanyaan:*\n'
        '• stok titan truck plus 205L\n'
        '• cogs ceplattyn sf 30\n'
        '• china nc 30% (setelah query COGS)\n'
        '• rekap total stok'
    )
    await update.message.reply_text(msg, parse_mode='Markdown')


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id):
        return
    
    msg = (
        '📖 *Bantuan StockBot:*\n\n'
        '━━━ *FITUR STOK* ━━━\n'
        '*Cek produk:*\n'
        '`stok titan truck plus 205L`\n\n'
        '*Analytics:*\n'
        '`produk stok paling banyak`\n'
        '`top 20 stok terbanyak`\n'
        '`berapa produk stok kosong`\n'
        '`rekap total stok`\n'
        '`top 10 stok di ID30`\n\n'
        '━━━ *FITUR COGS* ━━━\n'
        '*Cek COGS:*\n'
        '`cogs ceplattyn sf 30`\n'
        '`cost titan truck`\n'
        '`cogs 10234` (pakai material code)\n\n'
        '*Hitung margin (setelah cek COGS):*\n'
        '`china nc 30%` → hitung harga jual\n'
        '`local jual 100000` → hitung NC%\n\n'
        '*Reset context:*\n'
        '`stop` / `clear` / `reset`\n\n'
        '━━━ *COMMANDS* ━━━\n'
        '/reload — update data Excel dari Drive\n'
        '/status — lihat info data'
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
    
    user_id = update.effective_user.id
    if user_id in SESSIONS:
        lines.append('')
        lines.append(f'💬 *Session aktif:* `{SESSIONS[user_id]["product_query"]}`')
    
    await update.message.reply_text('\n'.join(lines), parse_mode='Markdown')


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        await update.message.reply_text(f'❌ Akses ditolak.\nUser ID Anda: `{user_id}`', parse_mode='Markdown')
        return
    
    user_text = update.message.text
    logger.info(f"Query from {user_id}: {user_text}")
    
    reply = route_message(user_text, user_id)
    
    if len(reply) > 4000:
        chunks = [reply[i:i+4000] for i in range(0, len(reply), 4000)]
        for chunk in chunks:
            await update.message.reply_text(chunk, parse_mode='Markdown')
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
        logger.warning("⚠️ GDRIVE_COGS_FILE_ID belum di-set — fitur COGS off.")
    
    logger.info("🚀 Starting StockBot v2 (with COGS)...")
    success, msg = load_all_data()
    logger.info(f"Initial load:\n{msg}")
    
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("reload", reload_cmd))
    app.add_handler(CommandHandler("status", status_cmd))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    
    logger.info("✅ Bot is running...")
    app.run_polling(drop_pending_updates=True)


if __name__ == '__main__':
    main()
