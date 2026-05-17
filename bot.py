"""
StockBot Telegram v3.1 — Inventory, COGS, & AI Assistant
Deploy ke Railway, baca data dari Google Drive, powered by Claude Sonnet 4.6
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
from anthropic import Anthropic

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
TELEGRAM_TOKEN = os.environ.get('TELEGRAM_TOKEN', '')
GDRIVE_FILE_ID = os.environ.get('GDRIVE_FILE_ID', '')
GDRIVE_COGS_FILE_ID = os.environ.get('GDRIVE_COGS_FILE_ID', '')
ALLOWED_USER_IDS = os.environ.get('ALLOWED_USER_IDS', '')
ANTHROPIC_API_KEY = os.environ.get('ANTHROPIC_API_KEY', '')

WAREHOUSE_COLS = ['ID30', 'ID40']
CLAUDE_MODEL = 'claude-sonnet-4-5-20250929'

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

claude_client = None
if ANTHROPIC_API_KEY:
    claude_client = Anthropic(api_key=ANTHROPIC_API_KEY)
    logger.info("✅ Claude AI client initialized")
else:
    logger.warning("⚠️ ANTHROPIC_API_KEY belum di-set, AI features off")

DATA = {
    'stock': None,
    'stock_sheet': None,
    'cogs': None,
    'cogs_sheet': None,
    'filename': None,
}

SESSIONS = {}
CHAT_HISTORY = {}
MAX_HISTORY = 5

COGS_KEYWORDS = {'COGS', 'COST', 'HARGA MODAL', 'MODAL', 'HPP'}
CLEAR_KEYWORDS = {'STOP', 'CLEAR', 'RESET'}
NC_PATTERN = re.compile(r'NC\s*(\d+(?:\.\d+)?)\s*%?', re.IGNORECASE)
PRICE_PATTERN = re.compile(r'JUAL\s*(?:RP\s*)?([\d.,]+)', re.IGNORECASE)


# ─────────────────────────────────────────────
# GOOGLE DRIVE LOADER
# ─────────────────────────────────────────────
def download_from_gdrive(file_id, as_gsheet=False):
    if as_gsheet:
        # Google Sheets native → export as xlsx
        url = f'https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx'
    else:
        # Regular file (xlsx upload) → direct download
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
        logger.info("📥 Downloading STOCK Excel...")
        try:
            excel_data = download_from_gdrive(GDRIVE_FILE_ID)
            xl = pd.ExcelFile(excel_data)
        except Exception:
            logger.info("📥 Retry as Google Sheets export...")
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
        for col in WAREHOUSE_COLS:
            if col not in df.columns:
                df[col] = 0
        for col in WAREHOUSE_COLS:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        df['Total'] = df[WAREHOUSE_COLS].sum(axis=1)
        df['desc_clean'] = df['Material Description'].astype(str).str.strip().str.upper()
        DATA['stock'] = df
        DATA['stock_sheet'] = latest_sheet
        logger.info(f"✅ Loaded STOCK: {len(df)} products from '{latest_sheet}'")
        return True, f"Stok: sheet '{latest_sheet}', {len(df):,} produk"
    except Exception as e:
        logger.error(f"❌ Stock load error: {e}")
        return False, f"Error loading stock: {str(e)}"


def load_cogs_data():
    if not GDRIVE_COGS_FILE_ID:
        return False, "COGS file ID belum di-set."
    try:
        logger.info("📥 Downloading COGS Excel...")
        # Try regular Excel first, fallback to Google Sheets export
        try:
            excel_data = download_from_gdrive(GDRIVE_COGS_FILE_ID)
            xl = pd.ExcelFile(excel_data)
        except Exception:
            logger.info("📥 Retry as Google Sheets export...")
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
        # Drop hanya row tanpa Material Code (truly empty), simpan COGS=0 untuk produk baru
        df = df[df['Material Code'].notna() & (df['Material Code'] != '') & (df['Material Code'] != 'nan')].reset_index(drop=True)
        DATA['cogs'] = df
        DATA['cogs_sheet'] = latest_sheet
        logger.info(f"✅ Loaded COGS: {len(df)} entries from '{latest_sheet}'")
        return True, f"COGS: sheet '{latest_sheet}', {len(df):,} entries"
    except Exception as e:
        logger.error(f"❌ COGS load error: {e}")
        return False, f"Error loading COGS: {str(e)}"


def load_all_data():
    stock_ok, stock_msg = load_stock_data()
    cogs_ok, cogs_msg = load_cogs_data()
    messages = [f"{'✅' if stock_ok else '❌'} {stock_msg}",
                f"{'✅' if cogs_ok else '⚠️'} {cogs_msg}"]
    return stock_ok, '\n'.join(messages)


# ─────────────────────────────────────────────
# STOCK QUERY
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
    return None


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
        return analytics_result
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
# COGS QUERY
# ─────────────────────────────────────────────
def extract_cogs_query(user_input):
    text = user_input.upper().strip()
    for kw in COGS_KEYWORDS:
        text = text.replace(kw, ' ')
    stop_words = {'BERAPA', 'YA', 'DONG', 'PRODUK', 'BARANG', 'UNTUK'}
    words = [w for w in text.split() if w not in stop_words and len(w) > 0]
    return ' '.join(words).strip()


def query_cogs(user_input, user_id):
    df = DATA['cogs']
    if df is None or len(df) == 0:
        return '⚠️ Data COGS belum ter-load. Coba ketik /reload'
    query = extract_cogs_query(user_input)
    if not query:
        return '❌ Tolong sebutkan nama produk atau material code.'
    code_mask = df['code_clean'] == query
    if code_mask.any():
        result = df[code_mask].copy()
    else:
        words = [w for w in query.split() if len(w) > 1]
        if not words:
            return f'❌ Query terlalu pendek: "{user_input}"'
        mask = pd.Series([True] * len(df))
        for kw in words:
            mask = mask & df['desc_clean'].str.contains(kw, na=False, regex=False)
        result = df[mask].copy()
        if result.empty and len(words) > 2:
            mask2 = pd.Series([True] * len(df))
            for kw in words[:2]:
                mask2 = mask2 & df['desc_clean'].str.contains(kw, na=False, regex=False)
            result = df[mask2].copy()
    if result.empty:
        return None
    SESSIONS[user_id] = {
        'product_query': query,
        'matches': result.copy()
    }
    lines = [f'💰 *COGS — {len(result)} varian ditemukan:*\n']
    for _, row in result.iterrows():
        code = row['Material Code']
        desc = str(row['Material Description']).strip()
        source = str(row['Source']).strip()
        cogs = int(row['COGS'])
        update = str(row['Update']).strip()
        nc_prev = str(row.get('Rata2 NC Sebelumnya', '')).strip()
        
        lines.append(f'• `[{code}]` {desc}')
        if cogs > 0:
            lines.append(f'   Source: *{source}* | COGS: *Rp {cogs:,}*')
        else:
            lines.append(f'   Source: *{source}* | COGS: _belum ada_')
        if nc_prev and nc_prev != 'nan' and nc_prev != '' and nc_prev != 'None':
            lines.append(f'   📊 Rata2 NC Sebelumnya: *{nc_prev}*')
        lines.append(f'   📅 {update}\n')
    lines.append('─────────────────')
    lines.append('💡 Lanjut hitung margin? Ketik misal:')
    lines.append('   `china nc 30%` atau `local jual 100000`')
    lines.append('   Ketik `stop` untuk reset.')
    return '\n'.join(lines)


# ─────────────────────────────────────────────
# CALCULATOR
# ─────────────────────────────────────────────
def parse_number(text):
    cleaned = re.sub(r'[^\d]', '', text)
    return int(cleaned) if cleaned else 0


def calculate_margin(user_input, user_id):
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
                f'Pilihan: *{sources_list}*'
            )
    row = matches[matches['source_clean'] == target_source].iloc[0]
    cogs = int(row['COGS'])
    desc = str(row['Material Description']).strip()
    code = row['Material Code']
    update = str(row['Update']).strip()
    nc_prev = str(row.get('Rata2 NC Sebelumnya', '')).strip()
    
    # Guard: COGS belum ada
    if cogs <= 0:
        return (
            f'⚠️ *{desc}* (`[{code}]`)\n'
            f'Source: *{target_source.title()}*\n\n'
            f'COGS belum ada di database, tidak bisa hitung margin.\n'
            f'Update kolom COGS di Google Sheets dulu ya.'
        )
    
    nc_match = NC_PATTERN.search(user_input)
    price_match = PRICE_PATTERN.search(user_input)
    
    # Footer NC sebelumnya
    nc_footer = ''
    if nc_prev and nc_prev != 'nan' and nc_prev != '' and nc_prev != 'None':
        nc_footer = f'\n📊 Rata2 NC Sebelumnya: *{nc_prev}*'
    
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
            f'{nc_footer}'
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
                f'• NC:         *{nc_value:.1f}%*\n'
                f'{nc_footer}'
            )
        return (
            f'💰 *{desc}*\n'
            f'`[{code}]` Source: *{target_source.title()}*\n\n'
            f'• COGS:        Rp {cogs:,}\n'
            f'• Harga Jual:  Rp {harga_jual:,}\n'
            f'• NC:          *{nc_value:.1f}%*\n\n'
            f'📐 Formula: NC = (HJ - COGS) / HJ\n'
            f'📅 COGS update: {update}'
            f'{nc_footer}'
        )
    return None


def clear_session(user_id):
    cleared = False
    if user_id in SESSIONS:
        del SESSIONS[user_id]
        cleared = True
    if user_id in CHAT_HISTORY:
        del CHAT_HISTORY[user_id]
        cleared = True
    if cleared:
        return '🔄 Session & chat history di-reset. Silakan tanya yang baru.'
    return '✅ Tidak ada session aktif.'


# ─────────────────────────────────────────────
# CLAUDE AI HANDLER
# ─────────────────────────────────────────────
def build_data_context():
    context_parts = []
    if DATA['stock'] is not None:
        df = DATA['stock']
        context_parts.append(f"=== DATA STOK (sheet: {DATA['stock_sheet']}) ===")
        context_parts.append(f"Total produk: {len(df):,}")
        context_parts.append(f"Warehouse: ID30, ID40")
        context_parts.append(f"Kolom: Material Description, ID30, ID40, Total")
        context_parts.append(f"Produk dengan stok > 0: {len(df[df['Total'] > 0]):,}")
        context_parts.append(f"Produk dengan stok 0: {len(df[df['Total'] == 0]):,}")
    if DATA['cogs'] is not None:
        df = DATA['cogs']
        sources = df['Source'].unique().tolist()
        context_parts.append(f"\n=== DATA COGS (sheet: {DATA['cogs_sheet']}) ===")
        context_parts.append(f"Total entries: {len(df):,}")
        context_parts.append(f"Sources tersedia: {', '.join(sources)}")
        context_parts.append(f"Kolom: Material Code, Material Description, Source, COGS (IDR), Update, Rata2 NC Sebelumnya")
    return '\n'.join(context_parts)


def search_data_for_ai(user_input):
    results = {'stock_matches': [], 'cogs_matches': []}
    query_upper = user_input.upper()
    stop_words = {'STOK', 'STOCK', 'COGS', 'COST', 'HARGA', 'BERAPA', 'YA', 'DONG',
                  'PRODUK', 'BARANG', 'UNTUK', 'DARI', 'KEMASAN', 'YANG', 'DI',
                  'DAN', 'SEMUA', 'ADA', 'CEK', 'CARI', 'LIHAT', 'TUNJUKKAN',
                  'MODAL', 'HPP', 'JUAL', 'NC', 'BAGAIMANA', 'GIMANA', 'APA',
                  'SAJA', 'MANA', 'LU', 'GW', 'GUE', 'KAMU', 'SAYA'}
    words = [w for w in query_upper.split() if w not in stop_words and len(w) > 1]
    if not words:
        return results
    if DATA['stock'] is not None:
        df = DATA['stock']
        mask = pd.Series([True] * len(df))
        for kw in words[:3]:
            mask = mask & df['desc_clean'].str.contains(kw, na=False, regex=False)
        matches = df[mask].head(20)
        for _, row in matches.iterrows():
            results['stock_matches'].append({
                'desc': str(row['Material Description']).strip(),
                'ID30': int(row['ID30']),
                'ID40': int(row['ID40']),
                'Total': int(row['Total'])
            })
    if DATA['cogs'] is not None:
        df = DATA['cogs']
        mask = pd.Series([True] * len(df))
        for kw in words[:3]:
            mask = mask & df['desc_clean'].str.contains(kw, na=False, regex=False)
        matches = df[mask].head(20)
        for _, row in matches.iterrows():
            nc_prev = str(row.get('Rata2 NC Sebelumnya', '')).strip()
            results['cogs_matches'].append({
                'code': row['Material Code'],
                'desc': str(row['Material Description']).strip(),
                'source': str(row['Source']).strip(),
                'cogs': int(row['COGS']),
                'update': str(row['Update']).strip(),
                'rata2_nc_sebelumnya': nc_prev if nc_prev != 'nan' else ''
            })
    return results


def ask_claude(user_input, user_id):
    if not claude_client:
        return '⚠️ Fitur AI belum aktif. ANTHROPIC_API_KEY belum di-set.'
    try:
        data_context = build_data_context()
        relevant_data = search_data_for_ai(user_input)
        
        system_prompt = f"""Kamu adalah StockBot, asisten AI untuk perusahaan pelumas yang membantu cek inventory dan COGS.

PERAN KAMU:
- Jawab pertanyaan user tentang stok, COGS, harga jual, dan margin
- Berikan analisa dan rekomendasi bisnis kalau diminta
- Toleran terhadap typo dan natural language
- Gunakan Bahasa Indonesia kasual tapi profesional
- Output ringkas dan jelas (max 200 kata), pakai bullet point dan emoji secukupnya

ATURAN PENTING:
1. JANGAN mengarang data — hanya jawab berdasarkan data yang saya kasih
2. Kalau data tidak ada di context, bilang "data tidak ditemukan"
3. Format angka pakai pemisah ribuan (contoh: Rp 50.000)
4. Formula NC: (Harga Jual - COGS) / Harga Jual
5. Formula Harga Jual: COGS / (1 - NC%)
6. Pakai Markdown Telegram: *bold*, `code`
7. Kolom "Rata2 NC Sebelumnya" berisi rata-rata NC historis per produk. Gunakan untuk perbandingan atau saran pricing.

KONTEKS DATA:
{data_context}

DATA RELEVAN DENGAN QUERY USER:
{json.dumps(relevant_data, indent=2, ensure_ascii=False)}
"""
        history = CHAT_HISTORY.get(user_id, [])
        messages = history + [{"role": "user", "content": user_input}]
        
        response = claude_client.messages.create(
            model=CLAUDE_MODEL,
            max_tokens=1024,
            system=system_prompt,
            messages=messages
        )
        reply = response.content[0].text
        history.append({"role": "user", "content": user_input})
        history.append({"role": "assistant", "content": reply})
        if len(history) > MAX_HISTORY * 2:
            history = history[-(MAX_HISTORY * 2):]
        CHAT_HISTORY[user_id] = history
        logger.info(f"🤖 Claude reply for {user_id}: {len(reply)} chars")
        return reply
    except Exception as e:
        logger.error(f"❌ Claude AI error: {e}")
        return f'⚠️ AI error: {str(e)[:100]}\n\nCoba pertanyaan yang lebih spesifik atau ketik /help'


# ─────────────────────────────────────────────
# SMART ROUTER
# ─────────────────────────────────────────────
def is_simple_query(text_upper):
    simple_patterns = [
        r'^(STOK|STOCK)\s+\w',
        r'^(COGS|COST)\s+\w',
        r'^TOP\s+\d+',
        r'PALING\s+(BANYAK|SEDIKIT)',
        r'^REKAP',
        r'(STOK|PRODUK)\s+KOSONG',
    ]
    return any(re.search(p, text_upper) for p in simple_patterns)


def route_message(user_input, user_id):
    text_upper = user_input.upper().strip()
    
    if text_upper in CLEAR_KEYWORDS:
        return clear_session(user_id)
    
    if user_id in SESSIONS:
        has_nc = bool(NC_PATTERN.search(user_input))
        has_price = bool(PRICE_PATTERN.search(user_input))
        if has_nc or has_price:
            result = calculate_margin(user_input, user_id)
            if result:
                return result
    
    if is_simple_query(text_upper):
        if any(kw in text_upper for kw in COGS_KEYWORDS):
            result = query_cogs(user_input, user_id)
            if result:
                return result
        result = query_stock(user_input)
        if result:
            return result
    
    if claude_client:
        logger.info(f"🤖 Routing to Claude AI: {user_input[:80]}")
        return ask_claude(user_input, user_id)
    
    result = query_stock(user_input)
    if result:
        return result
    return (
        '🤖 Saya belum mengerti pertanyaan tersebut.\n\n'
        'Coba ketik /help untuk lihat contoh pertanyaan.'
    )


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
    ai_status = '🤖 AI aktif (Claude Sonnet)' if claude_client else '⚠️ AI off'
    msg = (
        '👋 *Halo! Selamat datang di StockBot v3.1*\n\n'
        f'{ai_status}\n\n'
        'Saya bisa bantu:\n'
        '• Cek stok inventory\n'
        '• Cek COGS produk\n'
        '• Hitung margin (NC)\n'
        '• Analisa & rekomendasi bisnis\n\n'
        '*Commands:*\n'
        '/reload — refresh data\n'
        '/status — info data\n'
        '/help — bantuan\n\n'
        'Coba tanya saya apa saja!'
    )
    await update.message.reply_text(msg, parse_mode='Markdown')


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id):
        return
    msg = (
        '📖 *Bantuan StockBot v3.1:*\n\n'
        '━━━ *FITUR STOK* ━━━\n'
        '`stok titan truck plus 205L`\n'
        '`top 20 stok terbanyak`\n'
        '`rekap total stok`\n'
        '`top 10 stok di ID30`\n\n'
        '━━━ *FITUR COGS* ━━━\n'
        '`cogs ceplattyn sf 30`\n'
        '`cogs 10234` (material code)\n'
        '`china nc 30%` (hitung harga jual)\n'
        '`local jual 100000` (hitung NC)\n\n'
        '━━━ *FITUR AI* 🤖 ━━━\n'
        'Tanya bebas dalam bahasa natural:\n'
        '• "harga ceplattyn sf 30 205L dari china?"\n'
        '• "produk mana paling profitable?"\n'
        '• "kasih saran restock minggu depan"\n'
        '• "bandingin titan vs ceplattyn"\n\n'
        '━━━ *RESET* ━━━\n'
        '`stop` / `clear` / `reset`'
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
    if claude_client:
        lines.append('🤖 *AI:* Claude Sonnet 4.6 aktif')
    else:
        lines.append('⚠️ *AI:* off (set ANTHROPIC_API_KEY)')
    user_id = update.effective_user.id
    if user_id in SESSIONS:
        lines.append(f'\n💬 *Session COGS:* `{SESSIONS[user_id]["product_query"]}`')
    if user_id in CHAT_HISTORY:
        lines.append(f'💬 *Chat history:* {len(CHAT_HISTORY[user_id])//2} turn')
    await update.message.reply_text('\n'.join(lines), parse_mode='Markdown')


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        await update.message.reply_text(f'❌ Akses ditolak.\nUser ID: `{user_id}`', parse_mode='Markdown')
        return
    user_text = update.message.text
    logger.info(f"Query from {user_id}: {user_text}")
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action='typing')
    reply = route_message(user_text, user_id)
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
        logger.warning("⚠️ GDRIVE_COGS_FILE_ID belum di-set — fitur COGS off.")
    if not ANTHROPIC_API_KEY:
        logger.warning("⚠️ ANTHROPIC_API_KEY belum di-set — fitur AI off.")
    
    logger.info("🚀 Starting StockBot v3.1 (with AI + NC History)...")
    success, msg = load_all_data()
    logger.info(f"Initial load:\n{msg}")
    
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("reload", reload_cmd))
    app.add_handler(CommandHandler("status", status_cmd))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    
    logger.info("✅ Bot is running with AI...")
    app.run_polling(drop_pending_updates=True)


if __name__ == '__main__':
    main()
