"""
StockBot Telegram v5.0 — Full AI Agent with Tool Calling
Powered by Gemini. AI decides when & how to search the data.
"""
import os
import re
import io
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
GEMINI_MODEL = 'gemini-2.5-flash'  # Optimal untuk reasoning & tool calling
MAX_TOOL_TURNS = 6  # Max berapa kali AI bisa panggil tool dalam 1 query

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

DATA = {
    'stock': None,
    'stock_sheet': None,
    'cogs': None,
    'cogs_sheet': None,
}

CHAT_HISTORY = {}
MAX_HISTORY = 10

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
# TOOLS — Fungsi yang bisa dipanggil AI
# ─────────────────────────────────────────────
def _stock_row_to_dict(row):
    return {
        'code': str(row.get('code_clean', '')),
        'description': str(row['Material Description']).strip(),
        'ID30': int(row['ID30']) if pd.notna(row['ID30']) else 0,
        'ID40': int(row['ID40']) if pd.notna(row['ID40']) else 0,
        'total': int(row['Total']) if pd.notna(row['Total']) else 0
    }

def _cogs_row_to_dict(row):
    nc_prev = str(row.get('Rata2 NC Sebelumnya', '')).strip()
    return {
        'code': str(row['Material Code']),
        'description': str(row['Material Description']).strip(),
        'source': str(row['Source']).strip(),
        'cogs_rupiah': int(row['COGS']) if pd.notna(row['COGS']) else 0,
        'update_label': str(row['Update']).strip(),
        'rata2_nc_sebelumnya': nc_prev if nc_prev != 'nan' else ''
    }

def tool_search_stock(keywords: str, limit: int = 30) -> dict:
    if DATA['stock'] is None:
        return {'error': 'Data stok belum ter-load', 'results': []}

    df = DATA['stock']
    if not keywords or not keywords.strip():
        return {'error': 'Keywords kosong', 'results': []}

    kw_list = [k.strip().upper() for k in keywords.split() if k.strip()]
    mask = pd.Series([True] * len(df), index=df.index)
    for kw in kw_list:
        mask &= df['desc_clean'].str.contains(re.escape(kw), na=False, regex=True)

    matched = df[mask]
    total = len(matched)
    results = [_stock_row_to_dict(row) for _, row in matched.head(limit).iterrows()]

    return {
        'total_matches': total,
        'returned': len(results),
        'truncated': total > limit,
        'results': results
    }

def tool_search_cogs(keywords: str, limit: int = 30) -> dict:
    if DATA['cogs'] is None:
        return {'error': 'Data COGS belum ter-load', 'results': []}

    df = DATA['cogs']
    if not keywords or not keywords.strip():
        return {'error': 'Keywords kosong', 'results': []}

    kw_list = [k.strip().upper() for k in keywords.split() if k.strip()]
    mask = pd.Series([True] * len(df), index=df.index)
    for kw in kw_list:
        mask &= df['desc_clean'].str.contains(re.escape(kw), na=False, regex=True)

    matched = df[mask]
    total = len(matched)
    results = [_cogs_row_to_dict(row) for _, row in matched.head(limit).iterrows()]

    return {
        'total_matches': total,
        'returned': len(results),
        'truncated': total > limit,
        'results': results
    }

def tool_get_by_code(code: str) -> dict:
    if not code or not code.strip():
        return {'error': 'Code kosong'}

    code = code.strip().upper()
    stock_results = []
    cogs_results = []

    if DATA['stock'] is not None:
        df = DATA['stock']
        matched = df[df['code_clean'] == code]
        stock_results = [_stock_row_to_dict(row) for _, row in matched.iterrows()]

    if DATA['cogs'] is not None:
        df = DATA['cogs']
        matched = df[df['code_clean'] == code]
        cogs_results = [_cogs_row_to_dict(row) for _, row in matched.iterrows()]

    return {
        'code': code,
        'stock': stock_results,
        'cogs': cogs_results,
        'stock_count': len(stock_results),
        'cogs_count': len(cogs_results)
    }

def tool_get_analytics() -> dict:
    if DATA['stock'] is None:
        return {'error': 'Data stok belum ter-load'}

    df = DATA['stock']
    return {
        'total_produk': len(df),
        'stok_kosong': int(len(df[df['Total'] == 0])),
        'ada_stok': int(len(df[df['Total'] > 0])),
        'total_id30': int(df['ID30'].sum() if pd.notna(df['ID30'].sum()) else 0),
        'total_id40': int(df['ID40'].sum() if pd.notna(df['ID40'].sum()) else 0),
        'grand_total': int(df['Total'].sum() if pd.notna(df['Total'].sum()) else 0),
        'top_10_terbanyak': [
            _stock_row_to_dict(row)
            for _, row in df[df['Total'] > 0].nlargest(10, 'Total').iterrows()
        ],
        'top_10_tersedikit': [
            _stock_row_to_dict(row)
            for _, row in df[df['Total'] > 0].nsmallest(10, 'Total').iterrows()
        ],
    }

def tool_list_empty_stock(keywords: str = '', limit: int = 50) -> dict:
    if DATA['stock'] is None:
        return {'error': 'Data stok belum ter-load', 'results': []}

    df = DATA['stock']
    empty = df[df['Total'] == 0]

    if keywords and keywords.strip():
        kw_list = [k.strip().upper() for k in keywords.split() if k.strip()]
        mask = pd.Series([True] * len(empty), index=empty.index)
        for kw in kw_list:
            mask &= empty['desc_clean'].str.contains(re.escape(kw), na=False, regex=True)
        empty = empty[mask]

    total = len(empty)
    results = [_stock_row_to_dict(row) for _, row in empty.head(limit).iterrows()]
    return {
        'total_matches': total,
        'returned': len(results),
        'truncated': total > limit,
        'results': results
    }

def tool_list_cogs_sources() -> dict:
    if DATA['cogs'] is None:
        return {'error': 'Data COGS belum ter-load'}
    df = DATA['cogs']
    sources = df['Source'].value_counts().to_dict()
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
TOOL_DECLARATIONS = [
    {
        'name': 'search_stock',
        'description': (
            'Cari produk di tabel STOK berdasarkan keyword di Material Description. '
            'Semua kata di parameter keywords harus muncul (AND search). '
            'Untuk produk pelumas, gunakan nama produk + spec/kemasan, contoh: '
            '"RENOLIN B 68 PLUS 1000L" atau "TITAN GT1 5W30". '
            'Return: kode material, description, ID30 (stok gudang 30), ID40 (stok gudang 40), total.'
        ),
        'parameters': {
            'type': 'object',
            'properties': {
                'keywords': {
                    'type': 'string',
                    'description': 'Keyword pencarian, dipisah spasi. Contoh: "RENOLIN B 68 PLUS"'
                },
                'limit': {
                    'type': 'integer',
                    'description': 'Max hasil, default 30'
                }
            },
            'required': ['keywords']
        }
    },
    {
        'name': 'search_cogs',
        'description': (
            'Cari di tabel COGS berdasarkan keyword nama produk (AND search). '
            'Return: COGS dalam Rupiah, Source (LOKAL/IMPORT), label update terakhir, '
            'dan rata-rata NC sebelumnya. '
            'PENTING: 1 produk fisik bisa punya MULTI-SOURCE dengan code berbeda. '
            'Selalu cek SEMUA hasil yang return, jangan ambil yang pertama saja.'
        ),
        'parameters': {
            'type': 'object',
            'properties': {
                'keywords': {
                    'type': 'string',
                    'description': 'Keyword pencarian. Contoh: "RENOLIN B 68 PLUS 1000L"'
                },
                'limit': {
                    'type': 'integer',
                    'description': 'Max hasil, default 30'
                }
            },
            'required': ['keywords']
        }
    },
    {
        'name': 'get_by_code',
        'description': (
            'Cari produk berdasarkan Material Code EXACT (kode 6+ digit angka). '
            'Cek di tabel STOK dan COGS sekaligus. '
            'PENTING: code di stok dan COGS bisa beda untuk produk yang sama — '
            'kalau salah satu tabel tidak match, lanjutkan dengan search_stock/search_cogs by description.'
        ),
        'parameters': {
            'type': 'object',
            'properties': {
                'code': {
                    'type': 'string',
                    'description': 'Material code, contoh: "602723550"'
                }
            },
            'required': ['code']
        }
    },
    {
        'name': 'get_analytics',
        'description': (
            'Ringkasan analytics inventory: total produk, stok kosong vs ada, '
            'total per warehouse, top 10 terbanyak & tersedikit. '
            'Pakai untuk pertanyaan umum seperti "rekap stok", "produk paling banyak", dll.'
        ),
        'parameters': {'type': 'object', 'properties': {}}
    },
    {
        'name': 'list_empty_stock',
        'description': (
            'List produk yang stoknya kosong (total = 0), optional filter dengan keyword. '
            'Pakai untuk pertanyaan "produk apa yang kosong" atau "stok kosong renolin".'
        ),
        'parameters': {
            'type': 'object',
            'properties': {
                'keywords': {
                    'type': 'string',
                    'description': 'Optional filter by keyword di description.'
                },
                'limit': {
                    'type': 'integer',
                    'description': 'Max hasil, default 50'
                }
            }
        }
    },
    {
        'name': 'list_cogs_sources',
        'description': 'List semua unique Source (LOKAL, IMPORT, dll) di tabel COGS.',
        'parameters': {'type': 'object', 'properties': {}}
    },
]


# ─────────────────────────────────────────────
# GEMINI CLIENT (UPDATED)
# ─────────────────────────────────────────────
gemini_initialized = False
if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)
    gemini_initialized = True
    logger.info(f"✅ Gemini configured ({GEMINI_MODEL}) with {len(TOOL_DECLARATIONS)} tools")
else:
    logger.warning("⚠️ GEMINI_API_KEY belum di-set!")

def get_system_prompt():
    stock_info = "belum ter-load"
    cogs_info = "belum ter-load"
    if DATA['stock'] is not None:
        stock_info = f"{len(DATA['stock']):,} produk dari sheet '{DATA['stock_sheet']}'"
    if DATA['cogs'] is not None:
        cogs_info = f"{len(DATA['cogs']):,} entries dari sheet '{DATA['cogs_sheet']}'"

    return f"""Kamu adalah StockBot, asisten AI cerdas untuk perusahaan pelumas. Kamu chat dengan user di Telegram seperti ngobrol biasa — santai, ramah, profesional, pakai Bahasa Indonesia natural.

DATA YANG TERSEDIA:
- Tabel STOK: {stock_info}. Kolom: Material Code, Material Description, ID30 (gudang 30), ID40 (gudang 40), Total.
- Tabel COGS: {cogs_info}. Kolom: Material Code, Material Description, Source, COGS (Rupiah), Update label, Rata2 NC Sebelumnya.

KAMU PUNYA TOOLS untuk akses data — JANGAN PERNAH ngarang data! Selalu pakai tools.

STRATEGI PENCARIAN (PENTING!):
1. Untuk produk spesifik, panggil search_stock DAN search_cogs.
2. Kalau hasil pertama tidak match yang user maksud, COBA LAGI dengan keyword berbeda:
   - Hilangkan kata yang tidak spesifik
   - Coba variasi penulisan ("1000L" vs "1000 L" vs "1000")
   - Coba dengan/tanpa spec (AU, MET, IBC, PLA, LG)
3. Kalau user kasih kode 6+ digit, pakai get_by_code dulu.
4. Kalau hasil truncated (truncated=true), narrow query lebih spesifik.
5. 1 produk fisik bisa punya MULTI VARIAN kemasan (20L, 205L, 1000L) DAN multi-source di COGS.
6. Material code di STOK dan COGS BISA BEDA untuk produk yang sama.

PENGETAHUAN PRODUK:
- Format nama: BRAND + GRADE/SERI + SPEC + KEMASAN
- Contoh: "RENOLIN B 68 PLUS 1000L IBC AU"
  - RENOLIN = brand, B 68 PLUS = grade, 1000L IBC = kemasan, AU = spec
- "renolin b 68 plus" tanpa kemasan = ada beberapa varian (20L, 205L, 1000L)

FORMULA:
- NC% = (Harga Jual - COGS) / Harga Jual × 100
- Harga Jual = COGS / (1 - NC%/100)

ATURAN OUTPUT:
- Jawab natural seperti ngobrol — JANGAN kaku/formal/bertele-tele.
- Format Rupiah pakai pemisah ribuan: Rp 36.934.577.
- Pakai Markdown Telegram: *bold*, `code`.
- Pakai emoji secukupnya (1-2), jangan berlebihan.
- Kalau data ada multi varian/source, LIST SEMUA — biar user yang pilih.
- Jangan tutup dengan basa-basi "Ada lagi yang bisa dibantu?".
- Kalau bener-bener tidak ketemu setelah beberapa kali coba, jujur bilang."""

def call_tool(tool_name: str, args: dict) -> dict:
    func = TOOL_FUNCTIONS.get(tool_name)
    if not func:
        return {'error': f'Unknown tool: {tool_name}'}
    try:
        result = func(**args)
        n = result.get('total_matches', result.get('stock_count', 'ok'))
        logger.info(f"🔧 {tool_name}({args}) → {n}")
        return result
    except Exception as e:
        logger.error(f"❌ Tool {tool_name} error: {e}")
        return {'error': str(e)}

def ask_gemini(user_input: str, user_id: int) -> str:
    if not gemini_initialized:
        return '⚠️ Fitur AI belum aktif. Hubungi admin untuk set GEMINI_API_KEY.'

    try:
        # Build history dengan format terstandarisasi untuk Gemini
        history = CHAT_HISTORY.get(user_id, [])
        gemini_history = []
        for msg in history:
            role = "model" if msg["role"] == "assistant" else "user"
            gemini_history.append({"role": role, "parts": [msg["content"]]})

        system_prompt = get_system_prompt()
        
        # Konfigurasi model: Masukkan tools langsung sebagai list dari deklarasi fungsi
        chat_model = genai.GenerativeModel(
            model_name=GEMINI_MODEL,
            tools=TOOL_DECLARATIONS,
            system_instruction=system_prompt
        )
        
        chat = chat_model.start_chat(history=gemini_history)
        response = chat.send_message(user_input)

        # Agent loop untuk menangani multi-turn tool calling
        for turn in range(MAX_TOOL_TURNS):
            function_calls = response.function_calls
            if not function_calls:
                break  # Berhenti jika tidak ada panggilan fungsi lagi

            tool_parts = []
            for fc in function_calls:
                tool_name = fc.name
                args = dict(fc.args) if fc.args else {}
                
                result = call_tool(tool_name, args)
                
                # Bungkus sesuai spesifikasi SDK untuk FunctionResponse
                tool_parts.append(
                    genai.types.Part(
                        function_response=genai.types.FunctionResponse(
                            name=tool_name,
                            response={'result': result}
                        )
                    )
                )

            # Kirim hasil eksekusi kembali ke model
            response = chat.send_message(tool_parts)

        # Ambil text respons akhir
        reply = response.text if response.text else ''
        
        if not reply:
            reply_parts = [part.text for part in response.candidates[0].content.parts if hasattr(part, 'text') and part.text]
            reply = '\n'.join(reply_parts).strip()

        if not reply:
            reply = '⚠️ AI tidak menghasilkan jawaban. Coba perkecil keyword pencarian Anda.'

        # Simpan pesan sukses ke history lokal
        history.append({"role": "user", "content": user_input})
        history.append({"role": "assistant", "content": reply})
        
        if len(history) > MAX_HISTORY * 2:
            history = history[-(MAX_HISTORY * 2):]
        CHAT_HISTORY[user_id] = history

        logger.info(f"🤖 Reply ke {user_id}: {len(reply)} chars")
        return reply

    except Exception as e:
        logger.error(f"❌ Gemini error: {e}", exc_info=True)
        return f'⚠️ AI error: {str(e)[:200]}\n\nCoba ulangi atau bersihkan sesi dengan /clear.'

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
    allowed = [int(x.strip()) for x in ALLOWED_USER_IDS.split(',') if x.strip()]
    return user_id in allowed

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        await update.message.reply_text(f'❌ Akses ditolak.\nUser ID: `{user_id}`', parse_mode='Markdown')
        return
    ai_status = '🤖 Agent Mode (Tool Calling)' if gemini_initialized else '⚠️ AI off'
    msg = (
        '👋 *StockBot v5.0 — Full AI Agent*\n\n'
        f'{ai_status}\n\n'
        'Ngobrol aja sama saya kayak chat biasa.\n'
        'Saya cari datanya sendiri & kasih jawaban akurat.\n\n'
        '*Contoh:*\n'
        '• cost renolin b 68 plus 1000\n'
        '• stok titan plus 205l ada berapa?\n'
        '• 602723550 ini barang apa?\n'
        '• kalau jual NC 25%, harganya berapa?\n'
        '• produk apa stoknya paling banyak?\n\n'
        '*Commands:*\n'
        '/reload — refresh data\n'
        '/status — info data\n'
        '/clear — reset chat'
    )
    await update.message.reply_text(msg, parse_mode='Markdown')

async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id):
        return
    msg = (
        '📖 *StockBot v5.0 — Help*\n\n'
        'Bot ini full AI agent. Tanya apa aja pakai bahasa natural — '
        'AI bakal mikir & cari datanya sendiri.\n\n'
        '*Commands:*\n'
        '/reload — refresh data dari Google Drive\n'
        '/status — info data ter-load\n'
        '/clear — reset chat history'
    )
    await update.message.reply_text(msg, parse_mode='Markdown')

async def reload_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id):
        return
    await update.message.reply_text('🔄 Loading data dari Google Drive...')
    success, message = load_all_data()
    icon = '✅' if success else '❌'
    await update.message.reply_text(f'{icon} Hasil:\n\n{message}')

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
        lines.append(f'   • Total: {len(df):,}')
        lines.append(f'   • Ada stok: {len(df[df["Total"] > 0]):,}')
        lines.append(f'   • Kosong: {len(df[df["Total"] == 0]):,}')
    lines.append('')
    if DATA['cogs'] is None:
        lines.append('⚠️ *COGS:* belum ter-load')
    else:
        df = DATA['cogs']
        lines.append(f'✅ *COGS:*')
        lines.append(f'   • Sheet: `{DATA["cogs_sheet"]}`')
        lines.append(f'   • Entries: {len(df):,}')
        lines.append(f'   • Sources: {df["Source"].nunique()}')
    lines.append('')
    if gemini_initialized:
        lines.append(f'🤖 *AI:* {GEMINI_MODEL} (Agent Mode)')
    else:
        lines.append('⚠️ *AI:* off')
    user_id = update.effective_user.id
    if user_id in CHAT_HISTORY:
        lines.append(f'\n💬 *History:* {len(CHAT_HISTORY[user_id])//2} turn')
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
        logger.error("❌ GEMINI_API_KEY tidak di-set!")
        return

    logger.info("🚀 Starting StockBot v5.0 (Agent Mode)...")
    success, msg = load_all_data()
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
