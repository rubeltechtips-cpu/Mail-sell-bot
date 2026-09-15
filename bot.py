import logging
import asyncio
import html

import os

import io

import time

from telegram import (

    Update, ReplyKeyboardMarkup, KeyboardButton, InputFile,

    ReplyKeyboardRemove, InlineKeyboardMarkup, InlineKeyboardButton

)

from telegram.ext import (

    ApplicationBuilder, CommandHandler, MessageHandler,

    filters, ContextTypes, ConversationHandler, CallbackQueryHandler

)

from telegram.constants import ChatMemberStatus

from openpyxl import Workbook, load_workbook

try:
    from google import genai
except ImportError:
    genai = None

import json
import re
from pathlib import Path

# ================ CONFIG ================

BOT_TOKEN = os.getenv("8349208659:AAEyJikjx1tUri_PztFGRca_lPT0WilJ0N0", "").strip()

# Gemini API: keep the key in an environment variable for safety.
# Windows CMD: set GEMINI_API_KEY=AQ.Ab8RN6KV-q795-satGMO8Ehdcbi7XXws2Eqhxfip8cGlE3rO8Q
# PowerShell: $env:GEMINI_API_KEY="AQ.Ab8RN6KV-q795-satGMO8Ehdcbi7XXws2Eqhxfip8cGlE3rO8Q"
GEMINI_API_KEY = os.getenv("AQ.Ab8RN6KV-q795-satGMO8Ehdcbi7XXws2Eqhxfip8cGlE3rO8Q", "").strip()
GEMINI_MODEL = "gemini-3.6-flash"

ADMIN_ID = 8061006207

ADMIN_USERNAME = "Rubel_QSB"

CHANNEL_USERNAME = "pocket_money_zone"

DATA_DIR = "categories"

# Low-stock alert settings
LOW_STOCK_LIMIT = 5          # Alert when stock is below 5 (0-4)
LOW_STOCK_CHECK_INTERVAL = 30  # Backup check every 30 seconds; alert only once per low-stock event
LOW_STOCK_FIRST_CHECK_DELAY = 10    # First check 10 seconds after bot starts

# Sold items waiting for admin download
SOLD_ITEMS_FILE = os.path.join(DATA_DIR, "sold_items.xlsx")

os.makedirs(DATA_DIR, exist_ok=True)

# ================ STATES ================

(

    BUY_MENU,

    BUY_SUB_MENU,

    ADMIN_PANEL,

    ADD_MAIN_CAT,

    REMOVE_MAIN_CAT,

    MANAGE_CATEGORY,

    MANAGE_SUB_CATEGORY,

    ADD_SUB_CAT,

    REMOVE_SUB_CAT,

    ADD_ITEMS,

    EDIT_PAYMENT,

    EDIT_PRICE_MAIN,

    EDIT_PRICE_SUB,

    RECEIVE_NEW_PRICE,

    GET_QUANTITY,

    WAIT_SCREENSHOT,

    DEPOSIT,

    GET_DEPOSIT_AMOUNT,

    DASHBOARD,

    SEND_NOTICE,

    VIEW_USER_PROFILE,

    SEARCH_USER_PROFILE,

    MANAGE_PAYMENT_CATEGORIES,

    # --- NEW STATES FOR BALANCE EDIT ---

    SEARCH_USER_FOR_BALANCE,

    BALANCE_EDIT_ACTION,

    RECEIVE_BALANCE_EDIT_AMOUNT,
    AI_ASSISTANT,
    AI_SUPPORT,
    AI_STOCK_SEARCH,
    AI_NL_ORDER,
    AI_FILE_ANALYZER,
    AI_REPORT

) = range(32)

# ================ LOGGER ================

logging.basicConfig(format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO)

logger = logging.getLogger(__name__)

# ================ DATA ================

# New: Hierarchical categories and prices

categories = {}

prices = {}

payment_info = "Bkash: 017XXXXXXXX\nNagad: 018XXXXXXXX\nBinance: yourmail@gmail.com"

# New: User balances dictionary

balances = {}

user_sales = {}

user_deposits = {}

user_info = {} # To store user ID and username mapping

# New: Dashboard metrics

total_deposits = 0

total_sales = 0

sales_count_per_category = {}

transaction_log = []

dashboard_message = "স্বাগতম! এটি আপনার বটের ড্যাশবোর্ড।"

# List of categories that require manual delivery

MANUAL_DELIVERY_CATEGORIES = []

# --- User data persistence ---

def load_user_data():

    global balances, user_sales, user_deposits, user_info, total_deposits, total_sales, transaction_log, categories, prices, sales_count_per_category, MANUAL_DELIVERY_CATEGORIES

    try:

        with open("user_data.json", "r", encoding='utf-8') as f:

            data = json.load(f)

            balances.update({int(k): v for k, v in data.get("balances", {}).items()})

            user_sales.update({int(k): v for k, v in data.get("user_sales", {}).items()})

            user_deposits.update({int(k): v for k, v in data.get("user_deposits", {}).items()})

            user_info.update({int(k): v for k, v in data.get("user_info", {}).items()})

            total_deposits = data.get("total_deposits", 0)

            total_sales = data.get("total_sales", 0)

            transaction_log = data.get("transaction_log", [])

            categories.update(data.get("categories", {}))

            prices.update(data.get("prices", {}))

            sales_count_per_category.update(data.get("sales_count_per_category", {}))

            MANUAL_DELIVERY_CATEGORIES = data.get("manual_delivery_categories", [])

    except FileNotFoundError:

        pass

def save_user_data():

    with open("user_data.json", "w", encoding='utf-8') as f:

        data = {

            "balances": balances,

            "user_sales": user_sales,

            "user_deposits": user_deposits,

            "user_info": user_info,

            "total_deposits": total_deposits,

            "total_sales": total_sales,

            "transaction_log": transaction_log,

            "categories": categories,

            "prices": prices,

            "sales_count_per_category": sales_count_per_category,

            "manual_delivery_categories": MANUAL_DELIVERY_CATEGORIES

        }

        json.dump(data, f, ensure_ascii=False, indent=4)

        

load_user_data()

# ================ HELPERS ================

def get_stock_path(main_cat: str, sub_cat: str) -> str:
    """TXT inventory path: one item per line."""
    file_name = f"{main_cat}_{sub_cat}.txt".replace(" ", "_").replace("-", "_")
    return os.path.join(DATA_DIR, file_name)


def get_excel_path(main_cat: str, sub_cat: str) -> str:
    """Backward-compatible alias; inventory is now stored in TXT."""
    return get_stock_path(main_cat, sub_cat)


def _legacy_xlsx_path(main_cat: str, sub_cat: str) -> str:
    file_name = f"{main_cat}_{sub_cat}.xlsx".replace(" ", "_").replace("-", "_")
    return os.path.join(DATA_DIR, file_name)


def _read_txt_items(path: str) -> list:
    if not os.path.exists(path):
        return []
    try:
        text = Path(path).read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        text = Path(path).read_text(encoding="utf-8", errors="replace")
    return [line.strip() for line in text.replace("\r", "").split("\n") if line.strip()]


def _write_txt_items(path: str, items: list):
    Path(path).write_text(
        "\n".join(str(item).strip() for item in items if str(item).strip()),
        encoding="utf-8"
    )


def _migrate_legacy_xlsx_to_txt(main_cat: str, sub_cat: str):
    """Migrate an old inventory XLSX to TXT once, if needed."""
    txt_path = get_stock_path(main_cat, sub_cat)
    xlsx_path = _legacy_xlsx_path(main_cat, sub_cat)
    if os.path.exists(txt_path) or not os.path.exists(xlsx_path):
        return
    try:
        wb = load_workbook(xlsx_path, read_only=True, data_only=True)
        ws = wb.active
        items = [str(row[0]).strip() for row in ws.iter_rows(values_only=True)
                 if row and row[0] not in (None, "")]
        wb.close()
        _write_txt_items(txt_path, items)
        logger.info("Migrated legacy XLSX to TXT: %s / %s (%s items)", main_cat, sub_cat, len(items))
    except Exception:
        logger.exception("Failed to migrate legacy XLSX: %s / %s", main_cat, sub_cat)


def ensure_excel(main_cat: str, sub_cat: str):
    """Backward-compatible name; ensures the TXT inventory exists."""
    path = get_stock_path(main_cat, sub_cat)
    _migrate_legacy_xlsx_to_txt(main_cat, sub_cat)
    if not os.path.exists(path):
        Path(path).write_text("", encoding="utf-8")


def add_item_to_excel(main_cat: str, sub_cat: str, item: str):
    """Backward-compatible name; append one item to TXT inventory."""
    ensure_excel(main_cat, sub_cat)
    with open(get_stock_path(main_cat, sub_cat), "a", encoding="utf-8") as f:
        f.write(str(item).strip() + "\n")


def pop_items_from_excel(main_cat: str, sub_cat: str, qty: int):
    """Backward-compatible name; remove the first qty items from TXT inventory."""
    _migrate_legacy_xlsx_to_txt(main_cat, sub_cat)
    path = get_stock_path(main_cat, sub_cat)
    if not os.path.exists(path):
        return []
    items = _read_txt_items(path)
    if len(items) < qty:
        return []
    result = items[:qty]
    _write_txt_items(path, items[qty:])
    return result


def ensure_sold_items_file():
    """Create the sold-items workbook if it does not exist."""
    if not os.path.exists(SOLD_ITEMS_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Sold Items"
        ws.append(["Main Category", "Sub Category", "Item", "Sold At", "User ID", "Price"])
        wb.save(SOLD_ITEMS_FILE)


def record_sold_items(items, main_cat: str, sub_cat: str, user_id: int, price_per_item: float = 0):
    """Append successfully delivered items to the admin sold-items workbook."""
    if not items:
        return
    ensure_sold_items_file()
    wb = load_workbook(SOLD_ITEMS_FILE)
    ws = wb.active
    sold_at = time.strftime("%Y-%m-%d %H:%M:%S")
    for item in items:
        ws.append([main_cat, sub_cat, str(item), sold_at, user_id, price_per_item])
    wb.save(SOLD_ITEMS_FILE)


def _load_sold_rows():
    """Load sold-item rows from the single master workbook."""
    ensure_sold_items_file()
    wb = load_workbook(SOLD_ITEMS_FILE, data_only=True)
    ws = wb.active
    rows = [list(row) for row in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


def get_sold_category_counts():
    """Return {main_category: sold_item_count} from the master sold-items file."""
    rows = _load_sold_rows()
    counts = {}
    for row in rows[1:]:
        if not row:
            continue
        main_cat = str(row[0] or '').strip()
        item = row[2] if len(row) > 2 else None
        if main_cat and item not in (None, ''):
            counts[main_cat] = counts.get(main_cat, 0) + 1
    return counts


def create_sold_category_xlsx(main_cat: str):
    """Create an XLSX containing only sold items of one main category."""
    rows = _load_sold_rows()
    if len(rows) <= 1:
        return None, 0

    header = rows[0]
    selected = [row for row in rows[1:] if str(row[0] or '').strip() == str(main_cat).strip()]
    if not selected:
        return None, 0

    wb = Workbook()
    ws = wb.active
    ws.title = "Sold Items"
    ws.append(header)
    for row in selected:
        ws.append(row)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for column in ws.columns:
        max_len = 0
        col_letter = column[0].column_letter
        for cell in column:
            max_len = max(max_len, len(str(cell.value or "")))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 45)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, len(selected)


def remove_sold_category_rows(main_cat: str):
    """Remove only one category from the master sold-items workbook."""
    rows = _load_sold_rows()
    if len(rows) <= 1:
        return 0

    header = rows[0]
    remaining = [row for row in rows[1:] if str(row[0] or '').strip() != str(main_cat).strip()]
    removed = len(rows) - 1 - len(remaining)

    wb = Workbook()
    ws = wb.active
    ws.title = "Sold Items"
    ws.append(header)
    for row in remaining:
        ws.append(row)
    wb.save(SOLD_ITEMS_FILE)
    return removed


def create_sold_items_xlsx():
    """Build an in-memory XLSX of all sold items waiting for admin download."""
    rows = _load_sold_rows()
    if len(rows) <= 1:
        return None, 0
    wb = Workbook()
    ws = wb.active
    ws.title = "Sold Items"
    for row in rows:
        ws.append(row)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for column in ws.columns:
        max_len = 0
        col_letter = column[0].column_letter
        for cell in column:
            max_len = max(max_len, len(str(cell.value or "")))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 45)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, len(rows) - 1


def clear_sold_items_file():
    """Clear the entire sold-items master workbook."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sold Items"
    ws.append(["Main Category", "Sub Category", "Item", "Sold At", "User ID", "Price"])
    wb.save(SOLD_ITEMS_FILE)


def get_sold_subcategory_counts(main_cat: str):
    """Return {sub_category: sold_item_count} for one main category."""
    rows = _load_sold_rows()
    counts = {}
    target_main = str(main_cat).strip()
    for row in rows[1:]:
        if not row:
            continue
        row_main = str(row[0] or '').strip()
        sub_cat = str(row[1] or '').strip() if len(row) > 1 else ''
        item = row[2] if len(row) > 2 else None
        if row_main == target_main and sub_cat and item not in (None, ''):
            counts[sub_cat] = counts.get(sub_cat, 0) + 1
    return counts


def create_sold_subcategory_xlsx(main_cat: str, sub_cat: str):
    """Create an XLSX containing only sold items of one main/sub category."""
    rows = _load_sold_rows()
    if len(rows) <= 1:
        return None, 0

    header = rows[0]
    target_main = str(main_cat).strip()
    target_sub = str(sub_cat).strip()
    selected = [
        row for row in rows[1:]
        if str(row[0] or '').strip() == target_main
        and str(row[1] or '').strip() == target_sub
    ]
    if not selected:
        return None, 0

    wb = Workbook()
    ws = wb.active
    ws.title = "Sold Items"
    ws.append(header)
    for row in selected:
        ws.append(row)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for column in ws.columns:
        max_len = 0
        col_letter = column[0].column_letter
        for cell in column:
            max_len = max(max_len, len(str(cell.value or "")))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 45)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, len(selected)


def remove_sold_subcategory_rows(main_cat: str, sub_cat: str):
    """Remove only one main/sub category from the master sold-items workbook."""
    rows = _load_sold_rows()
    if len(rows) <= 1:
        return 0

    header = rows[0]
    target_main = str(main_cat).strip()
    target_sub = str(sub_cat).strip()
    remaining = []
    removed = 0

    for row in rows[1:]:
        row_main = str(row[0] or '').strip() if row else ''
        row_sub = str(row[1] or '').strip() if len(row) > 1 else ''
        if row_main == target_main and row_sub == target_sub:
            removed += 1
        else:
            remaining.append(row)

    wb = Workbook()
    ws = wb.active
    ws.title = "Sold Items"
    ws.append(header)
    for row in remaining:
        ws.append(row)
    wb.save(SOLD_ITEMS_FILE)
    return removed


async def admin_sold_items_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Show main categories first; clicking one shows its sold sub-categories."""
    if update.effective_user.id != ADMIN_ID:
        return ADMIN_PANEL

    counts = get_sold_category_counts()
    if not counts:
        await update.message.reply_text(
            "📭 এখনো কোনো sold item জমা হয়নি।",
            reply_markup=ReplyKeyboardMarkup(
                [[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True
            )
        )
        return ADMIN_PANEL

    category_map = list(sorted(counts.keys(), key=lambda x: x.lower()))
    keyboard = []
    for index, cat in enumerate(category_map):
        keyboard.append([
            InlineKeyboardButton(
                f"📦 {cat} — {counts[cat]} পিস",
                callback_data=f"soldmain:{index}"
            )
        ])
    keyboard.append([InlineKeyboardButton("❌ Close", callback_data="soldclose")])
    context.user_data["sold_main_category_map"] = category_map

    await update.message.reply_text(
        "📦 <b>Sold Items</b>\n\n"
        "প্রথমে Main Category নির্বাচন করুন।\n"
        "তারপর সেই Category-এর sold Sub Category দেখাবে।\n\n"
        "👉 আপনি যে Sub Category-এর sold item নিতে চান, শুধু সেটিই নির্বাচন করতে পারবেন।",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode="HTML"
    )
    return ADMIN_PANEL


async def admin_sold_category_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Navigate main category -> sold sub-category -> download selected XLSX."""
    query = update.callback_query
    await query.answer()

    if update.effective_user.id != ADMIN_ID:
        await query.edit_message_text("❌ Unauthorized.")
        return

    data = query.data or ""

    if data == "soldclose":
        await query.edit_message_text("📦 Sold Items menu বন্ধ করা হয়েছে।")
        return

    if data == "soldback":
        counts = get_sold_category_counts()
        if not counts:
            await query.edit_message_text("📭 এখন কোনো sold item নেই।")
            return
        category_map = list(sorted(counts.keys(), key=lambda x: x.lower()))
        context.user_data["sold_main_category_map"] = category_map
        keyboard = [
            [InlineKeyboardButton(f"📦 {cat} — {counts[cat]} পিস", callback_data=f"soldmain:{i}")]
            for i, cat in enumerate(category_map)
        ]
        keyboard.append([InlineKeyboardButton("❌ Close", callback_data="soldclose")])
        await query.edit_message_text(
            "📦 <b>Sold Items</b>\n\nMain Category নির্বাচন করুন:",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode="HTML"
        )
        return

    if data.startswith("soldmain:"):
        try:
            index = int(data.split(":", 1)[1])
            category_map = context.user_data.get("sold_main_category_map", [])
            if index < 0 or index >= len(category_map):
                raise ValueError
            main_cat = category_map[index]
        except Exception:
            await query.edit_message_text("❌ Category selection expired. আবার Sold Items menu খুলুন।")
            return

        sub_counts = get_sold_subcategory_counts(main_cat)
        if not sub_counts:
            await query.edit_message_text(
                f"📭 <b>{html.escape(str(main_cat))}</b>-এ এখন কোনো sold sub-category নেই।",
                parse_mode="HTML"
            )
            return

        sub_map = list(sorted(sub_counts.keys(), key=lambda x: x.lower()))
        context.user_data["sold_subcategory_map"] = [(main_cat, sub) for sub in sub_map]
        keyboard = []
        for i, sub_cat in enumerate(sub_map):
            keyboard.append([
                InlineKeyboardButton(
                    f"📄 {sub_cat} — {sub_counts[sub_cat]} পিস",
                    callback_data=f"soldsub:{i}"
                )
            ])
        keyboard.append([InlineKeyboardButton("🔙 Main Category", callback_data="soldback")])
        keyboard.append([InlineKeyboardButton("❌ Close", callback_data="soldclose")])

        await query.edit_message_text(
            f"📦 <b>{html.escape(str(main_cat))}</b> — Sold Sub Categories\n\n"
            "যে Sub Category-এর sold item নিতে চান, সেটিতে চাপ দিন।",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode="HTML"
        )
        return

    if data.startswith("soldsub:"):
        try:
            index = int(data.split(":", 1)[1])
            sub_map = context.user_data.get("sold_subcategory_map", [])
            if index < 0 or index >= len(sub_map):
                raise ValueError
            main_cat, sub_cat = sub_map[index]
        except Exception:
            await query.edit_message_text("❌ Sub Category selection expired. আবার Sold Items menu খুলুন।")
            return

        output, count = create_sold_subcategory_xlsx(main_cat, sub_cat)
        if output is None or count == 0:
            await query.edit_message_text(
                f"📭 <b>{html.escape(str(sub_cat))}</b>-এ এখন কোনো sold item নেই।\n\n"
                "সম্ভবত অন্য জায়গা থেকে এই Sub Category-এর items ইতিমধ্যে নেওয়া হয়েছে।",
                parse_mode="HTML"
            )
            return

        try:
            safe_main = re.sub(r'[^A-Za-z0-9._-]+', '_', str(main_cat)).strip('_') or 'category'
            safe_sub = re.sub(r'[^A-Za-z0-9._-]+', '_', str(sub_cat)).strip('_') or 'subcategory'
            filename = f"sold_{safe_main}_{safe_sub}_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"

            await context.bot.send_document(
                chat_id=ADMIN_ID,
                document=InputFile(output, filename=filename),
                caption=(
                    f"📦 <b>Sold Items</b>\n\n"
                    f"Main Category: <b>{html.escape(str(main_cat))}</b>\n"
                    f"Sub Category: <b>{html.escape(str(sub_cat))}</b>\n"
                    f"মোট: <b>{count} পিস</b>\n\n"
                    "✅ নির্বাচিত Sub Category-এর XLSX file পাঠানো হয়েছে।"
                ),
                parse_mode="HTML"
            )

            removed = remove_sold_subcategory_rows(main_cat, sub_cat)
            remaining_counts = get_sold_category_counts()
            remaining_total = sum(remaining_counts.values())

            await query.edit_message_text(
                f"✅ <b>{html.escape(str(sub_cat))}</b> থেকে {count} পিস sold item আপনার কাছে পাঠানো হয়েছে।\n"
                f"🧹 এই Sub Category-এর {removed}টি row master sold file থেকে remove হয়েছে।\n\n"
                f"📦 অন্য sold item remaining: <b>{remaining_total} পিস</b>",
                parse_mode="HTML"
            )
        except Exception as e:
            logger.exception("Failed to send selected sold sub-category")
            await query.edit_message_text(
                "❌ XLSX file পাঠানো যায়নি। তাই sold item master file থেকে কোনো data remove করা হয়নি।\n\n"
                f"Error: {html.escape(str(e))}",
                parse_mode="HTML"
            )
        return

    await query.edit_message_text("❌ Invalid Sold Items selection।")


async def admin_download_sold_items(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Backward-compatible: open the category selection menu."""
    return await admin_sold_items_menu(update, context)


def count_items(main_cat: str, sub_cat: str) -> int:
    """Count stock from the TXT inventory."""
    _migrate_legacy_xlsx_to_txt(main_cat, sub_cat)
    return len(_read_txt_items(get_stock_path(main_cat, sub_cat)))


async def send_low_stock_alert(application):
    """Send one alert when a configured sub-category enters stock below 5."""
    alerted = application.bot_data.setdefault("low_stock_alerted", set())
    current_low_stock = set()
    newly_low_stock = []

    try:
        # Inventory TXT files are flat inside categories/.
        # categories dict is the authoritative main/sub-category list.
        for main_cat in sorted(categories.keys(), key=lambda x: str(x).lower()):
            for sub_cat in sorted(categories.get(main_cat, []), key=lambda x: str(x).lower()):
                stock_count = count_items(main_cat, sub_cat)
                key = (main_cat, sub_cat)

                if stock_count < LOW_STOCK_LIMIT:
                    current_low_stock.add(key)
                    if key not in alerted:
                        newly_low_stock.append((main_cat, sub_cat, stock_count))

        # Recovered stock can alert again if it later drops below 5.
        alerted.intersection_update(current_low_stock)

        if not newly_low_stock:
            return

        lines = [
            "⚠️ <b>Low Stock Alert</b>",
            "",
            f"নিচের Sub Category-গুলোতে {LOW_STOCK_LIMIT} পিসের কম item আছে:",
            "",
        ]
        for main_cat, sub_cat, stock_count in newly_low_stock:
            lines.append(
                f"🔸 <b>{html.escape(str(sub_cat))}</b>: "
                f"<b>{stock_count} পিস</b> বাকি "
                f"({html.escape(str(main_cat))})"
            )
        lines += [
            "",
            "📌 একই low-stock অবস্থায় alert শুধু ১ বার যাবে।",
            "📌 Stock ৫+ হলে পরে আবার ৫-এর নিচে নামলে নতুন alert যাবে।",
        ]

        await application.bot.send_message(
            chat_id=ADMIN_ID,
            text="\n".join(lines),
            parse_mode="HTML",
        )
        alerted.update((mc, sc) for mc, sc, _ in newly_low_stock)
        logger.info("Low-stock alert sent: %d sub-category(s)", len(newly_low_stock))
    except Exception:
        logger.exception("Failed to send low-stock alert")


async def check_low_stock_after_sale(application, main_cat, sub_cat):
    """Immediately notify admin if this sub-category falls below 5 after a sale."""
    try:
        stock_count = count_items(main_cat, sub_cat)
        if stock_count >= LOW_STOCK_LIMIT:
            return
        alerted = application.bot_data.setdefault("low_stock_alerted", set())
        key = (main_cat, sub_cat)
        if key in alerted:
            return
        await application.bot.send_message(
            chat_id=ADMIN_ID,
            text=(
                "⚠️ <b>Low Stock Alert</b>\n\n"
                f"🔸 <b>{html.escape(str(sub_cat))}</b>: <b>{stock_count} পিস</b> বাকি "
                f"({html.escape(str(main_cat))})\n\n"
                "📌 এই Sub Category-এর জন্য alert শুধু ১ বার পাঠানো হবে।"
            ),
            parse_mode="HTML",
        )
        alerted.add(key)
        logger.info("Immediate low-stock alert sent: %s / %s = %s", main_cat, sub_cat, stock_count)
    except Exception:
        logger.exception("Failed to send immediate low-stock alert")


async def low_stock_monitor(application):
    """Monitor stock periodically but notify only once per low-stock transition."""
    try:
        await asyncio.sleep(LOW_STOCK_FIRST_CHECK_DELAY)
        while True:
            await send_low_stock_alert(application)
            await asyncio.sleep(LOW_STOCK_CHECK_INTERVAL)
    except asyncio.CancelledError:
        raise
    except Exception:
        logger.exception("Low-stock monitor stopped unexpectedly")


async def post_init(application):
    old_task = application.bot_data.get("low_stock_monitor_task")
    if old_task and not old_task.done():
        old_task.cancel()

    application.bot_data["low_stock_monitor_task"] = asyncio.create_task(
        low_stock_monitor(application),
        name="low_stock_monitor"
    )
    logger.info("Low-stock monitor started: alert once when stock falls below %s", LOW_STOCK_LIMIT)


async def post_shutdown(application):
    task = application.bot_data.get("low_stock_monitor_task")
    if task and not task.done():
        task.cancel()
        try:
            await task
        except asyncio.CancelledError:
            pass


def get_total_stock(main_cat: str) -> int:

    total = 0

    if main_cat in categories:

        for sub_cat in categories[main_cat]:

            total += count_items(main_cat, sub_cat)

    return total


def normalize_text(value) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip().lower()


def search_stock_items(query: str, main_cat: str = "", sub_cat: str = "", limit: int = 20):
    """Search real stock rows from the bot's XLSX inventory."""
    q = normalize_text(query)
    results = []
    main_filter = normalize_text(main_cat)
    sub_filter = normalize_text(sub_cat)
    for mc, subs in categories.items():
        if main_filter and normalize_text(mc) != main_filter:
            continue
        for sc in subs:
            if sub_filter and normalize_text(sc) != sub_filter:
                continue
            path = get_excel_path(mc, sc)
            if not os.path.exists(path):
                continue
            try:
                wb = load_workbook(path, read_only=True, data_only=True)
                ws = wb.active
                for row in ws.iter_rows(values_only=True):
                    if not row or row[0] in (None, ""):
                        continue
                    item = str(row[0]).strip()
                    if not q or q in normalize_text(item) or q in normalize_text(sc) or q in normalize_text(mc):
                        results.append({
                            "main_cat": mc,
                            "sub_cat": sc,
                            "item": item,
                            "price": prices.get(mc, {}).get(sc, 0),
                        })
                        if len(results) >= limit:
                            wb.close()
                            return results
                wb.close()
            except Exception as e:
                logger.exception("Stock search error: %s", e)
    return results


def stock_summary(limit_per_sub=3):
    lines = []
    for mc, subs in categories.items():
        for sc in subs:
            count = count_items(mc, sc)
            price = prices.get(mc, {}).get(sc, 0)
            sample = search_stock_items("", mc, sc, limit_per_sub)
            sample_text = ", ".join(x["item"] for x in sample[:limit_per_sub]) if sample else ""
            lines.append(f"- {mc} / {sc}: stock={count}, price={price}, sample={sample_text}")
    return "\n".join(lines) if lines else "- কোনো stock/category নেই"


def build_sales_report(days: int = 7) -> dict:
    cutoff = time.time() - max(1, int(days)) * 86400
    sales = [t for t in transaction_log if len(t) >= 4 and t[0] == 'sale' and float(t[3]) >= cutoff]
    deposits = [t for t in transaction_log if len(t) >= 4 and t[0] == 'deposit' and float(t[3]) >= cutoff]
    sale_total = sum(float(t[2]) for t in sales)
    deposit_total = sum(float(t[2]) for t in deposits)
    category_sales = {}
    for cat, amount in sales_count_per_category.items():
        category_sales[cat] = amount
    return {
        "days": days,
        "sales_transactions": len(sales),
        "sales_amount": sale_total,
        "deposit_transactions": len(deposits),
        "deposit_amount": deposit_total,
        "users": len(user_info),
        "total_balance": sum(float(v) for v in balances.values()),
        "category_units_sold": category_sales,
    }


def extract_file_text(filename: str, data: bytes, max_chars: int = 50000) -> str:
    """Extract TXT/XLSX content locally before sending a bounded excerpt to Gemini."""
    lower = filename.lower()
    if lower.endswith('.txt'):
        for enc in ('utf-8', 'utf-8-sig', 'cp1252'):
            try:
                return data.decode(enc, errors='replace')[:max_chars]
            except Exception:
                pass
        return data.decode('utf-8', errors='replace')[:max_chars]
    if lower.endswith('.xlsx'):
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        chunks = []
        for ws in wb.worksheets:
            chunks.append(f"[SHEET: {ws.title}]")
            row_count = 0
            for row in ws.iter_rows(values_only=True):
                vals = [str(v) for v in row if v is not None]
                if vals:
                    chunks.append(" | ".join(vals))
                    row_count += 1
                if sum(len(x) + 1 for x in chunks) >= max_chars or row_count >= 5000:
                    break
            if sum(len(x) + 1 for x in chunks) >= max_chars:
                break
        wb.close()
        return "\n".join(chunks)[:max_chars]
    raise ValueError("শুধু TXT বা XLSX ফাইল সাপোর্ট করা হয়েছে।")


def gemini_text_sync(prompt: str) -> str:
    if not GEMINI_API_KEY:
        return "❌ Gemini API key সেট করা নেই। GEMINI_API_KEY environment variable-এ নতুন key দিন।"
    if genai is None:
        return "❌ google-genai package ইনস্টল নেই। চালান: pip install -U google-genai"
    try:
        client = genai.Client(api_key=GEMINI_API_KEY)
        response = client.models.generate_content(model=GEMINI_MODEL, contents=prompt)
        text = getattr(response, "text", None)
        return text.strip() if text else "❌ Gemini কোনো উত্তর দেয়নি।"
    except Exception as e:
        logger.exception("Gemini API error")
        return f"❌ Gemini API error: {e}"


def parse_order_with_gemini_sync(text: str) -> dict:
    """Use Gemini structured JSON output to turn natural language into an order request."""
    if not GEMINI_API_KEY or genai is None:
        return {"ok": False, "message": "Gemini API/package সেট করা নেই।"}
    schema = {
        "type": "object",
        "properties": {
            "main_category": {"type": "string"},
            "sub_category": {"type": "string"},
            "quantity": {"type": "integer"},
            "ok": {"type": "boolean"},
            "message": {"type": "string"},
        },
        "required": ["main_category", "sub_category", "quantity", "ok", "message"],
    }
    prompt = (
        "You are an order parser for a Telegram shop. Parse the customer's Bangla/English request. "
        "Choose only from the supplied categories. Never invent a category. Quantity must be a positive integer. "
        "If the request is ambiguous, set ok=false and explain what is missing.\n\n"
        f"CATEGORIES:\n{json.dumps(categories, ensure_ascii=False)}\n\n"
        f"CUSTOMER REQUEST:\n{text}"
    )
    try:
        client = genai.Client(api_key=GEMINI_API_KEY)
        response = client.models.generate_content(
            model=GEMINI_MODEL,
            contents=prompt,
            config={"response_mime_type": "application/json", "response_schema": schema},
        )
        raw = getattr(response, "text", "") or "{}"
        return json.loads(raw)
    except Exception as e:
        logger.exception("Order parser error")
        return {"ok": False, "message": f"AI order parser error: {e}"}


def match_category(name: str, options) -> str | None:
    n = normalize_text(name)
    for opt in options:
        if normalize_text(opt) == n:
            return opt
    for opt in options:
        if n and (n in normalize_text(opt) or normalize_text(opt) in n):
            return opt
    return None

def create_xlsx_file(items: list, file_name: str) -> io.BytesIO:

    """Creates an in-memory XLSX file from a list of items."""

    wb = Workbook()

    ws = wb.active

    ws.title = "Items"

    for item in items:

        ws.append([item])

    

    buffer = io.BytesIO()

    wb.save(buffer)

    buffer.seek(0)

    return buffer

def make_safe_filename(text: str) -> str:
    """Make a safe filename while keeping Unicode characters."""
    text = re.sub(r'[\\/:*?"<>|]+', '_', str(text))
    text = text.strip().strip('.')
    return text or "items"


async def send_items_xlsx(context, user_id: int, items: list, main_cat: str, sub_cat: str, order_qty: int):
    """Send purchased items to the user as an XLSX file."""
    if not items:
        return

    file_name = (
        f"order_{make_safe_filename(main_cat)}_"
        f"{make_safe_filename(sub_cat)}_{user_id}.xlsx"
    )
    xlsx_buffer = create_xlsx_file(items, file_name)
    await context.bot.send_document(
        chat_id=user_id,
        document=InputFile(xlsx_buffer, filename=file_name),
        caption=f"✅ আপনার অর্ডারের {len(items)}টি item XLSX ফাইলে দেওয়া হলো।"
    )


def add_items_from_txt_bytes(data: bytes) -> list:
    """Read one item per line from a TXT file."""
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = data.decode("utf-8", errors="replace")

    return [line.strip() for line in text.replace("\r", "").split("\n") if line.strip()]


def get_report_summary(transactions, days):

    """Calculates total deposits and sales for a given number of days."""

    end_timestamp = time.time()

    start_timestamp = end_timestamp - (days * 24 * 60 * 60)

    

    daily_deposits = 0

    daily_sales = 0

    

    for trans in transactions:
        # New sale format: (sale, user_id, main_cat, sub_cat, qty, total_price, timestamp)
        # Old format: (type, user_id, amount, timestamp)
        if len(trans) == 7:
            trans_type, _, _, _, _, amount, timestamp = trans
        elif len(trans) == 4:
            trans_type, _, amount, timestamp = trans
        else:
            continue

        if start_timestamp <= timestamp <= end_timestamp:
            if trans_type == 'deposit':
                daily_deposits += amount
            elif trans_type == 'sale':
                daily_sales += amount

    return daily_deposits, daily_sales

    

def get_user_transactions(user_id, transactions):

    """Filters transactions for a specific user."""

    return [t for t in transactions if t[1] == user_id]

# ================ HANDLERS ================

async def check_subscription(update: Update, context: ContextTypes.DEFAULT_TYPE):

    user_id = update.effective_user.id

    try:

        chat_member = await context.bot.get_chat_member(chat_id=f"@{CHANNEL_USERNAME}", user_id=user_id)

        if chat_member.status in [ChatMemberStatus.MEMBER, ChatMemberStatus.ADMINISTRATOR, ChatMemberStatus.OWNER]:

            return True

        else:

            keyboard = InlineKeyboardMarkup([

                [InlineKeyboardButton("Join Channel", url=f"https://t.me/{CHANNEL_USERNAME}")]

            ])

            await update.message.reply_text(

                "❌ আপনি এখনো আমাদের চ্যানেলে জয়েন করেননি।\n"

                "বট ব্যবহার করার জন্য অনুগ্রহ করে নিচের বাটনে ক্লিক করে চ্যানেলে জয়েন করুন।",

                reply_markup=keyboard

            )

            return False

    except Exception as e:

        logger.error(f"Error checking subscription: {e}")

        await update.message.reply_text("চ্যানেল সদস্যতা পরীক্ষা করতে সমস্যা হচ্ছে। অনুগ্রহ করে পরে আবার চেষ্টা করুন।")

        return False

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if not await check_subscription(update, context):

        return ConversationHandler.END

    

    user_id = update.effective_user.id

    username = update.effective_user.username

    first_name = update.effective_user.first_name

    

    # Store user info if it's new

    if user_id not in user_info:

        user_info[user_id] = {

            "username": username,

            "first_name": first_name,

            "last_name": update.effective_user.last_name,

            "id": user_id

        }

        save_user_data()

    current_balance = balances.get(user_id, 0)

    keyboard = [

        [KeyboardButton("🛒 Buy"), KeyboardButton("💰 Balance")],

        [KeyboardButton("💸 Deposit"), KeyboardButton("📞 Help")],
        [KeyboardButton("🤖 AI Support"), KeyboardButton("🔎 AI Stock Search")],
        [KeyboardButton("🛒 AI Order")],

    ]

    if update.effective_user.id == ADMIN_ID:

        keyboard.append([KeyboardButton("⚙️ Admin Panel")])

    

    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)

    await update.message.reply_text(f"👋 Welcome! আপনার বর্তমান ব্যালেন্স: {current_balance} টাকা।", reply_markup=reply_markup)

    return ConversationHandler.END

async def menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if not await check_subscription(update, context):

        return ConversationHandler.END

    text = update.message.text

    

    if text == "🛒 Buy":

        if not categories:

            await update.message.reply_text("⚠️ এখন কোনো category নেই।")

            return ConversationHandler.END

        

        keyboard = []

        for cat in categories.keys():

            keyboard.append([KeyboardButton(cat)])

        

        keyboard.append([KeyboardButton("🔙 Back to Main Menu")])

        await update.message.reply_text("🛒 ক্যাটাগরি বেছে নিন:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))

        return BUY_MENU

    if text == "💰 Balance":

        user_id = update.effective_user.id

        current_balance = balances.get(user_id, 0)

        await update.message.reply_text(f"আপনার বর্তমান ব্যালেন্স: {current_balance} টাকা।")

        return ConversationHandler.END

    if text == "💸 Deposit":

        await update.message.reply_text("আপনি কত টাকা ডিপোজিট করতে চান তা লিখুন:", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True))

        return DEPOSIT

    if text == "📞 Help":

        keyboard = InlineKeyboardMarkup([

            [InlineKeyboardButton("📞 Chat with Admin", url=f"tg://user?id={ADMIN_ID}")]

        ])

        await update.message.reply_text(

            "📞 Admin-এর সাথে যোগাযোগ করতে নিচের বাটনে ক্লিক করুন।",

            reply_markup=keyboard

        )

        return ConversationHandler.END

    if text == "🤖 AI Support":
        return await ai_support_start(update, context)

    if text == "🔎 AI Stock Search":
        return await ai_stock_search_start(update, context)

    if text == "🛒 AI Order":
        return await ai_nl_order_start(update, context)

    if text == "⚙️ Admin Panel":

        if update.effective_user.id == ADMIN_ID:

            # সরাসরি ড্যাশবোর্ড দেখাবে

            return await show_dashboard(update, context)

        else:

            await update.message.reply_text("❌ Unauthorized.")

            return ConversationHandler.END

            

    if text == "🔙 Back to Main Menu":

        return await start(update, context)

    

    return ConversationHandler.END

    

async def show_dashboard(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

    global total_deposits, total_sales, balances, sales_count_per_category, user_info, transaction_log, dashboard_message

    total_balance = sum(balances.values())

    total_users_count = len(user_info)

    

    stock_info = ""

    for main_cat, sub_cats in categories.items():

        # Markdown থেকে HTML-এ পরিবর্তন: **Main_Cat** এর বদলে <b>Main_Cat</b>

        stock_info += f"  - <b>{main_cat}</b>\n" 

        for sub_cat in sub_cats:

            count = count_items(main_cat, sub_cat)

            stock_info += f"  - {sub_cat}: {count} items\n"

    sorted_sales = sorted(sales_count_per_category.items(), key=lambda item: item[1], reverse=True)

    top_selling_info = ""

    for sub_cat, count in sorted_sales:

        top_selling_info += f"  - {sub_cat}: {count} sales\n"

    recent_transactions = ""
    # সর্বশেষ ১০টি লেনদেন দেখাবে (নতুনটি আগে)
    last_10_transactions = transaction_log[-10:]
    if last_10_transactions:
        for trans in reversed(last_10_transactions):
            if len(trans) == 7:
                trans_type, user_id, main_cat, sub_cat, qty, amount, timestamp = trans
            elif len(trans) == 4:
                trans_type, user_id, amount, timestamp = trans
                main_cat, sub_cat, qty = "", "", ""
            else:
                continue

            date_str = time.strftime('%I:%M %p %b %d', time.localtime(timestamp)).lstrip('0')
            user_data = user_info.get(user_id, {})
            username_raw = user_data.get("username")
            if username_raw:
                user_display = f"@{html.escape(str(username_raw))}"
            else:
                user_display = f"ID: {user_id}"

            if trans_type == 'deposit':
                recent_transactions += (
                    f"  - 💸 <b>ডিপোজিট:</b> {amount} টাকা "
                    f"(ব্যবহারকারী: <code>{user_display}</code>) — {date_str}\n"
                )
            elif trans_type == 'sale':
                if main_cat and sub_cat and qty != "":
                    unit_price = amount / qty if qty else 0
                    recent_transactions += (
                        f"  - 🛒 <b>বিক্রয়:</b> {html.escape(str(main_cat))} → "
                        f"{html.escape(str(sub_cat))} | {qty} পিস × {unit_price:.2f} টাকা = "
                        f"<b>{amount:.2f} টাকা</b> (ব্যবহারকারী: <code>{user_display}</code>) — {date_str}\n"
                    )
                else:
                    recent_transactions += (
                        f"  - 🛒 <b>বিক্রয়:</b> {amount} টাকা "
                        f"(ব্যবহারকারী: <code>{user_display}</code>) — {date_str}\n"
                    )
    else:
        recent_transactions = "  - কোনো সাম্প্রতিক লেনদেন নেই।\n"

    daily_deposits, daily_sales = get_report_summary(transaction_log, 1)

    weekly_deposits, weekly_sales = get_report_summary(transaction_log, 7)

    monthly_deposits, monthly_sales = get_report_summary(transaction_log, 30)

    # HTML ফরম্যাটিং ব্যবহার করে ড্যাশবোর্ড টেক্সট

    dashboard_text = (

        f"📝 <b>ড্যাশবোর্ড মেসেজ:</b>\n"

        f"<i>{html.escape(str(dashboard_message))}</i>\n" # *Italic* এর বদলে <i>Italic</i>

        "---------------------------\n"

        "📊 <b>ড্যাশবোর্ড সামারি</b>\n"

        f"👥 <b>মোট ব্যবহারকারী:</b> {total_users_count}\n"

        f"💰 <b>মোট ব্যালেন্স:</b> {sum(balances.values())} টাকা\n"

        f"🛒 <b>মোট বিক্রয়:</b> {total_sales} টাকা\n"

        f"💸 <b>মোট ডিপোজিট:</b> {total_deposits} টাকা\n"

        "---------------------------\n"

        "📈 <b>দৈনিক/সাপ্তাহিক/মাসিক রিপোর্ট</b>\n"

        f"<b>গত ২৪ ঘণ্টা:</b>\n"

        f"  - ডিপোজিট: {daily_deposits} টাকা\n"

        f"  - বিক্রয়: {daily_sales} টাকা\n"

        f"<b>গত ৭ দিন:</b>\n"

        f"  - ডিপোজিট: {weekly_deposits} টাকা\n"

        f"  - বিক্রয়: {weekly_sales} টাকা\n"

        f"<b>গত ৩০ দিন:</b>\n"

        f"  - ডিপোজিট: {monthly_deposits} টাকা\n"

        f"  - বিক্রয়: {monthly_sales} টাকা\n"

        "---------------------------\n"

        "📦 <b>বর্তমান স্টক তথ্য:</b>\n"

        f"{stock_info or '  - কোনো ক্যাটাগরি পাওয়া যায়নি।'}\n"

        "---------------------------\n"

        "📈 <b>সর্বাধিক বিক্রীত ক্যাটাগরি:</b>\n"

        f"{top_selling_info or '  - এখনো কোনো বিক্রয় হয়নি।'}\n"

        "---------------------------\n"

        "📜 <b>সর্বশেষ লেনদেন:</b>\n"

        f"{recent_transactions}\n"

    )

    

    keyboard = [

        [KeyboardButton("🔄 ড্যাশবোর্ড রিফ্রেশ করুন"), KeyboardButton("👥 ব্যবহারকারীর প্রোফাইল")],

        [KeyboardButton("📂 Manage Categories"), KeyboardButton("💰 Edit Price")],

        [KeyboardButton("✏️ Edit User Balance"), KeyboardButton("📢 Send Notice")],
        [KeyboardButton("🤖 Gemini AI"), KeyboardButton("📈 AI Report")],
        [KeyboardButton("📊 AI File Analyzer")],
        [KeyboardButton("📦 Sold Items")],

        [KeyboardButton("💳 Edit Payment Info"), KeyboardButton("💳 Manage Payment Categories")],

        [KeyboardButton("🔙 Back to Main Menu")]

    ]

    

    # *** মূল পরিবর্তন: parse_mode='Markdown' থেকে parse_mode='HTML' ***

    await update.message.reply_text(dashboard_text, reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True), parse_mode='HTML')

    return ADMIN_PANEL

async def handle_dashboard_refresh(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

    if update.message.text == "🔄 ড্যাশবোর্ড রিফ্রেশ করুন":

        return await show_dashboard(update, context)

    return await back_to_admin_panel_handler(update, context)

async def view_user_profile(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

    await update.message.reply_text("✍️ যে ব্যবহারকারীর প্রোফাইল দেখতে চান তার username বা User ID লিখুন:", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True))

    return SEARCH_USER_PROFILE

async def search_and_show_user_profile(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

        

    search_term = update.message.text.strip().lstrip('@')

    

    if search_term == "🔙 অ্যাডমিন প্যানেল":

        return await back_to_admin_panel_handler(update, context)

    found_user_id = None

    

    # Try to search by User ID first if the input is a number

    if search_term.isdigit():

        search_id = int(search_term)

        if search_id in user_info:

            found_user_id = search_id

    

    # If not found by ID, try to search by username

    if not found_user_id:

        for user_id, info in user_info.items():

            # এখানে info.get("username") যদি None হয়, তবে .lower() কল না করেই লজিক কাজ করবে

            if info.get("username") and info["username"].lower() == search_term.lower():

                found_user_id = user_id

                break

    if not found_user_id:

        await update.message.reply_text("❌ এই ইউজারনেম বা ইউজার আইডি এর কোনো ব্যবহারকারী পাওয়া যায়নি।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True))

        return SEARCH_USER_PROFILE

    user_transactions = get_user_transactions(found_user_id, transaction_log)

    

    balance = balances.get(found_user_id, 0)

    deposits = user_deposits.get(found_user_id, 0)

    sales = user_sales.get(found_user_id, 0)

    # Calculate time-based reports

    daily_deposits, daily_sales = get_report_summary(user_transactions, 1)

    weekly_deposits, weekly_sales = get_report_summary(user_transactions, 7)

    monthly_deposits, monthly_sales = get_report_summary(user_transactions, 30)

    yearly_deposits, yearly_sales = get_report_summary(user_transactions, 365)

    

    user_data = user_info.get(found_user_id, {})

    full_name = user_data.get("first_name", "") + (f" {user_data['last_name']}" if user_data.get("last_name") else "")

    # 🔥 মূল সংশোধন: NoneType ত্রুটির সমাধান

    username_raw = user_data.get('username')

    # username_raw যদি None হয়, তবে 'N/A' ব্যবহার হবে। তারপর HTML special characters escape হবে।

    username_safe = (username_raw or 'N/A').replace('<', '&lt;').replace('>', '&gt;')

    

    profile_text = (

        f"👤 <b>ব্যবহারকারী প্রোফাইল:</b>\n"

        f"নাম: {full_name}\n"

        f"Username: <code>@{username_safe}</code>\n" # <code> tags are safe for User ID/username

        f"ID: <code>{found_user_id}</code>\n"

        "---------------------------\n"

        f"💰 <b>বর্তমান ব্যালেন্স:</b> {balance} টাকা\n"

        f"💸 <b>মোট ডিপোজিট:</b> {deposits} টাকা\n"

        f"🛒 <b>মোট খরচ:</b> {sales} টাকা\n"

        "---------------------------\n"

        "📈 <b>লেনদেনের রিপোর্ট</b>\n"

        f"<b>গত ২৪ ঘণ্টা:</b>\n"

        f"  - ডিপোজিট: {daily_deposits} টাকা\n"

        f"  - খরচ: {daily_sales} টাকা\n"

        f"<b>গত ৭ দিন:</b>\n"

        f"  - ডিপোজিট: {weekly_deposits} টাকা\n"

        f"  - খরচ: {weekly_sales} টাকা\n"

        f"<b>গত ৩০ দিন:</b>\n"

        f"  - ডিপোজিট: {monthly_deposits} টাকা\n"

        f"  - খরচ: {monthly_sales} টাকা\n"

        f"<b>গত ১ বছর:</b>\n"

        f"  - ডিপোজিট: {yearly_deposits} টাকা\n"

        f"  - খরচ: {yearly_sales} টাকা\n"

    )

    

    # *** মূল পরিবর্তন: parse_mode='HTML' ব্যবহার করা হয়েছে ***

    await update.message.reply_text(profile_text, parse_mode='HTML', reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True))

    

    return SEARCH_USER_PROFILE

# =========================================================

# ================ NEW: BALANCE EDIT HANDLERS ===============

# =========================================================

async def edit_user_balance_start(update: Update, context: ContextTypes.DEFAULT_TYPE):

    """Starts the process of editing a user's balance."""

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

    

    await update.message.reply_text(

        "✏️ যে ব্যবহারকারীর ব্যালেন্স পরিবর্তন করতে চান তার username বা User ID লিখুন:",

        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True)

    )

    return SEARCH_USER_FOR_BALANCE

async def search_user_for_balance(update: Update, context: ContextTypes.DEFAULT_TYPE):

    """Searches for the user and asks for the balance action."""

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

        

    search_term = update.message.text.strip().lstrip('@')

    

    if search_term == "🔙 অ্যাডমিন প্যানেল":

        return await back_to_admin_panel_handler(update, context)

    found_user_id = None

    

    if search_term.isdigit():

        search_id = int(search_term)

        if search_id in user_info:

            found_user_id = search_id

    

    if not found_user_id:

        for user_id, info in user_info.items():

            if info.get("username") and info["username"].lower() == search_term.lower():

                found_user_id = user_id

                break

    if not found_user_id:

        await update.message.reply_text(

            "❌ এই ইউজারনেম বা ইউজার আইডি এর কোনো ব্যবহারকারী পাওয়া যায়নি। আবার চেষ্টা করুন।",

            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True)

        )

        return SEARCH_USER_FOR_BALANCE

    # Store found user ID for the next steps

    context.user_data['edit_balance_user_id'] = found_user_id

    

    user_data = user_info.get(found_user_id, {})

    username = user_data.get('username', 'N/A')

    current_balance = balances.get(found_user_id, 0)

    keyboard = [

        [KeyboardButton("➕ Add Balance"), KeyboardButton("➖ Remove Balance")],

        [KeyboardButton("✍️ Set New Balance")],

        [KeyboardButton("🔙 অ্যাডমিন প্যানেল")]

    ]

    

    await update.message.reply_text(

        f"👤 ব্যবহারকারী: @{username} (ID: {found_user_id})\n"

        f"💰 বর্তমান ব্যালেন্স: {current_balance} টাকা\n\n"

        f"আপনি কী করতে চান?",

        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

    )

    

    return BALANCE_EDIT_ACTION

async def balance_edit_action_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):

    """Handles the admin's choice of action (Add, Remove, Set)."""

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

    action_text = update.message.text

    

    if action_text == "🔙 অ্যাডমিন প্যানেল":

        return await back_to_admin_panel_handler(update, context)

    if action_text not in ["➕ Add Balance", "➖ Remove Balance", "✍️ Set New Balance"]:

        await update.message.reply_text("❌ অনুগ্রহ করে নিচের বাটন থেকে একটি অপশন বেছে নিন।")

        return BALANCE_EDIT_ACTION

    context.user_data['balance_edit_action'] = action_text

    

    if action_text == "➕ Add Balance":

        prompt = "✍️ কত টাকা যোগ করতে চান তা লিখুন:"

    elif action_text == "➖ Remove Balance":

        prompt = "✍️ কত টাকা সরাতে চান তা লিখুন:"

    else: # "✍️ Set New Balance"

        prompt = "✍️ নতুন ব্যালেন্স কত হবে তা লিখুন:"

        

    await update.message.reply_text(

        prompt,

        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True)

    )

    

    return RECEIVE_BALANCE_EDIT_AMOUNT

async def receive_balance_edit_amount(update: Update, context: ContextTypes.DEFAULT_TYPE):

    """Receives the amount and performs the balance update."""

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

    amount_str = update.message.text.strip()

    

    if amount_str == "🔙 অ্যাডমিন প্যানেল":

        context.user_data.pop('edit_balance_user_id', None)

        context.user_data.pop('balance_edit_action', None)

        return await back_to_admin_panel_handler(update, context)

    if not amount_str.isdigit() or float(amount_str) < 0:

        await update.message.reply_text(

            "❌ অনুগ্রহ করে একটি ধনাত্মক সংখ্যা লিখুন।",

            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True)

        )

        return RECEIVE_BALANCE_EDIT_AMOUNT

        

    amount = float(amount_str)

    user_id = context.user_data.get('edit_balance_user_id')

    action = context.user_data.get('balance_edit_action')

    if not user_id or not action:

        await update.message.reply_text("❌ একটি ত্রুটি ঘটেছে। অনুগ্রহ করে আবার শুরু করুন।")

        return await back_to_admin_panel_handler(update, context)

    old_balance = balances.get(user_id, 0)

    new_balance = 0

    

    if action == "➕ Add Balance":

        new_balance = old_balance + amount

        balances[user_id] = new_balance

        user_message = f"✅ অ্যাডমিন আপনার ব্যালেন্সে {amount} টাকা যোগ করেছে।\nআপনার নতুন ব্যালেন্স: {new_balance} টাকা।"

        admin_message = f"✅ ব্যবহারকারীর ব্যালেন্সে {amount} টাকা যোগ করা হয়েছে।\nনতুন ব্যালেন্স: {new_balance} টাকা।"

    elif action == "➖ Remove Balance":

        new_balance = old_balance - amount

        if new_balance < 0:

            new_balance = 0 # Balance can't be negative

        balances[user_id] = new_balance

        user_message = f"✅ অ্যাডমিন আপনার ব্যালেন্স থেকে {amount} টাকা সরিয়ে নিয়েছে।\nআপনার নতুন ব্যালেন্স: {new_balance} টাকা।"

        admin_message = f"✅ ব্যবহারকারীর ব্যালেন্স থেকে {amount} টাকা সরানো হয়েছে।\nনতুন ব্যালেন্স: {new_balance} টাকা।"

    elif action == "✍️ Set New Balance":

        new_balance = amount

        balances[user_id] = new_balance

        user_message = f"✅ অ্যাডমিন আপনার নতুন ব্যালেন্স {new_balance} টাকা সেট করেছে।"

        admin_message = f"✅ ব্যবহারকারীর নতুন ব্যালেন্স {new_balance} টাকা সেট করা হয়েছে।"

    save_user_data()

    # Notify the user

    try:

        await context.bot.send_message(chat_id=user_id, text=user_message)

    except Exception as e:

        logger.error(f"Failed to notify user {user_id} about balance change: {e}")

        admin_message += "\n⚠️ ব্যবহারকারীকে নোটিশ পাঠানো যায়নি।"

    # Confirm to admin and go back to panel

    await update.message.reply_text(admin_message)

    

    context.user_data.pop('edit_balance_user_id', None)

    context.user_data.pop('balance_edit_action', None)

    

    return await show_dashboard(update, context)

async def deposit_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):

    amount_str = update.message.text.strip()

    if amount_str == "🔙 Back to Main Menu":

        return await start(update, context)

    if not amount_str.isdigit():

        await update.message.reply_text("❌ অনুগ্রহ করে শুধু সংখ্যা লিখুন।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True))

        return DEPOSIT

    amount = int(amount_str)

    

    context.user_data["deposit_amount"] = amount

    

    keyboard = [

        [KeyboardButton("🔙 Back to Main Menu")]

    ]

    

    await update.message.reply_text(

        f"💳 এখন অনুগ্রহ করে {amount} টাকা পেমেন্ট করুন:\n{payment_info}\n\n"

        f"📸 পেমেন্টের পর স্ক্রিনশট পাঠান।",

        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

    )

    return GET_DEPOSIT_AMOUNT

async def receive_deposit_screenshot(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if update.message.text == "🔙 Back to Main Menu":

        return await start(update, context)

    if not update.message.photo:

        await update.message.reply_text("❌ অনুগ্রহ করে শুধু একটি ছবি পাঠান।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True))

        return GET_DEPOSIT_AMOUNT

    

    user = update.effective_user

    deposit_amount = context.user_data.get("deposit_amount", 0)

    

    username = user.username if user.username else 'N/A'

    caption = (

        f"🔔 **নতুন ডিপোজিট রিকোয়েস্ট!** 🔔\n"

        f"User: @{username}\n"

        f"Amount: {deposit_amount}\n"

        f"UserID: {user.id}"

    )

    keyboard = InlineKeyboardMarkup([

        [InlineKeyboardButton("✅ Confirm Deposit", callback_data=f"deposit_confirm:{user.id}:{deposit_amount}"),

         InlineKeyboardButton("❌ Cancel Deposit", callback_data=f"deposit_cancel:{user.id}")]

    ])

    await context.bot.send_photo(chat_id=ADMIN_ID, photo=update.message.photo[-1].file_id, caption=caption, reply_markup=keyboard)

    

    await update.message.reply_text("✅ আপনার ডিপোজিট রিকোয়েস্ট Admin-কে পাঠানো হয়েছে। Admin নিশ্চিত করার পর আপনার ব্যালেন্স যোগ হবে।")

    

    context.user_data.clear()

    return ConversationHandler.END

async def back_to_admin_panel_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

    # এখন Admin Panel-এ ফেরার মানেই ড্যাশবোর্ড দেখানো

    return await show_dashboard(update, context)


def build_ai_context():
    """Create a compact, non-sensitive summary of current bot statistics for Gemini."""
    total_stock = sum(count_items(main_cat, sub_cat) for main_cat, sub_cats in categories.items() for sub_cat in sub_cats)
    category_lines = []
    for main_cat, sub_cats in categories.items():
        for sub_cat in sub_cats:
            category_lines.append(f"- {main_cat} / {sub_cat}: stock={count_items(main_cat, sub_cat)}, price={prices.get(main_cat, {}).get(sub_cat, 'N/A')}")
    recent_lines = []
    for trans in transaction_log[-10:]:
        if len(trans) == 7:
            trans_type, user_id, main_cat, sub_cat, qty, amount, timestamp = trans
            recent_lines.append(f"- {trans_type}: user={user_id}, item={main_cat}/{sub_cat}, qty={qty}, amount={amount}, time={time.strftime('%Y-%m-%d %H:%M', time.localtime(timestamp))}")
        elif len(trans) == 4:
            trans_type, user_id, amount, timestamp = trans
            recent_lines.append(f"- {trans_type}: user={user_id}, amount={amount}, time={time.strftime('%Y-%m-%d %H:%M', time.localtime(timestamp))}")
    return (
        f"Total users: {len(user_info)}\n"
        f"Total balance: {sum(balances.values())}\n"
        f"Total sales: {total_sales}\n"
        f"Total deposits: {total_deposits}\n"
        f"Total stock: {total_stock}\n"
        f"Dashboard message: {dashboard_message}\n"
        f"Categories:\n{chr(10).join(category_lines) if category_lines else '- none'}\n"
        f"Recent transactions:\n{chr(10).join(recent_lines) if recent_lines else '- none'}"
    )


def ask_gemini_sync(prompt: str) -> str:
    return gemini_text_sync(prompt)


async def gemini_ai_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END
    await update.message.reply_text(
        "🤖 Gemini AI Admin Assistant\n\n"
        "আপনি stock, sales, deposit, user ও business report নিয়ে প্রশ্ন করতে পারেন।\n\n"
        "উদাহরণ:\n"
        "• কোন category-তে stock কম?\n"
        "• গত ৭ দিনের sales কেমন?\n"
        "• আমার stock দেখে কীভাবে sales বাড়ানো যায়?\n\n"
        "প্রশ্ন লিখুন।",
        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True),
    )
    return AI_ASSISTANT


async def gemini_ai_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END
    text = (update.message.text or "").strip()
    if text == "🔙 অ্যাডমিন প্যানেল":
        return await back_to_admin_panel_handler(update, context)
    await update.message.reply_text("⏳ Gemini বিশ্লেষণ করছে...")
    prompt = (
        "You are the private AI assistant for a Telegram shop bot administrator. "
        "Answer in simple Bangla. Use only supplied bot facts for statistics. Do not invent data. "
        "Give practical recommendations.\n\n"
        f"BOT CONTEXT:\n{build_ai_context()}\n\n"
        f"DETAILED STOCK:\n{stock_summary()}\n\n"
        f"ADMIN QUESTION:\n{text}"
    )
    answer = await asyncio.to_thread(gemini_text_sync, prompt)
    for i in range(0, len(answer), 4000):
        await update.message.reply_text(answer[i:i + 4000])
    return AI_ASSISTANT


async def ai_support_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "🤖 AI Customer Support\n\nআপনার প্রশ্ন লিখুন। আমি product, price, stock ও bot-এর সাধারণ নিয়ম সম্পর্কে সাহায্য করব।\n\n"
        "⚠️ অর্ডার করতে চাইলে লিখতে পারেন: ‘2টা [product] চাই’।",
        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True),
    )
    return AI_SUPPORT


async def ai_support_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await check_subscription(update, context):
        return ConversationHandler.END
    text = (update.message.text or "").strip()
    if text == "🔙 Back to Main Menu":
        return await start(update, context)
    results = search_stock_items(text, limit=15)
    stock_context = "\n".join(
        f"- {r['main_cat']} / {r['sub_cat']} | price={r['price']} | item={r['item']}" for r in results
    ) or "- matching stock পাওয়া যায়নি"
    prompt = (
        "You are a helpful customer-support agent for a Telegram digital shop. Answer in simple Bangla. "
        "Use only the supplied categories, prices and matching stock. If no match exists, say so honestly. "
        "Do not claim an order was placed. For orders, tell the customer they can use the natural-language order option.\n\n"
        f"CATEGORIES/PRICES:\n{json.dumps(prices, ensure_ascii=False)}\n\n"
        f"MATCHING STOCK:\n{stock_context}\n\n"
        f"CUSTOMER:\n{text}"
    )
    await update.message.reply_text("⏳ উত্তর তৈরি হচ্ছে...")
    answer = await asyncio.to_thread(gemini_text_sync, prompt)
    for i in range(0, len(answer), 4000):
        await update.message.reply_text(answer[i:i + 4000])
    return AI_SUPPORT


async def ai_stock_search_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "🔎 AI Stock/Product Search\n\nProduct-এর নাম, category বা item-এর কোনো অংশ লিখুন।",
        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True),
    )
    return AI_STOCK_SEARCH


async def ai_stock_search_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await check_subscription(update, context):
        return ConversationHandler.END
    text = (update.message.text or "").strip()
    if text == "🔙 Back to Main Menu":
        return await start(update, context)
    results = search_stock_items(text, limit=30)
    if not results:
        await update.message.reply_text("❌ আপনার search-এর সাথে মিলে কোনো stock পাওয়া যায়নি।")
        return AI_STOCK_SEARCH
    lines = ["🔎 Matching products/items:\n"]
    seen = set()
    for r in results:
        key = (r['main_cat'], r['sub_cat'])
        if key not in seen:
            lines.append(f"📦 {r['main_cat']} → {r['sub_cat']} | দাম: {r['price']} টাকা | Stock: {count_items(r['main_cat'], r['sub_cat'])}")
            seen.add(key)
        if len(lines) >= 20:
            break
    lines.append("\n💡 অর্ডার করতে Natural Language Order ব্যবহার করুন।")
    await update.message.reply_text("\n".join(lines))
    return AI_STOCK_SEARCH


async def ai_nl_order_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "🛒 Natural Language Order\n\nউদাহরণ:\n"
        "• 2টা Gmail চাই\n• 5টি [subcategory] দিন\n• [category] থেকে 3টা লাগবে\n\n"
        "আপনার চাহিদা লিখুন।",
        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True),
    )
    return AI_NL_ORDER


async def ai_nl_order_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await check_subscription(update, context):
        return ConversationHandler.END
    text = (update.message.text or "").strip()
    if text == "🔙 Back to Main Menu":
        return await start(update, context)
    await update.message.reply_text("⏳ আপনার order বুঝে নিচ্ছি...")
    parsed = await asyncio.to_thread(parse_order_with_gemini_sync, text)
    if not parsed.get("ok"):
        await update.message.reply_text(f"❌ {parsed.get('message', 'Order বুঝতে পারিনি।')}\n\nCategory ও quantity স্পষ্ট করে লিখুন।")
        return AI_NL_ORDER
    main_cat = match_category(parsed.get("main_category", ""), categories.keys())
    if not main_cat:
        await update.message.reply_text("❌ এই main category পাওয়া যায়নি।")
        return AI_NL_ORDER
    sub_cat = match_category(parsed.get("sub_category", ""), categories.get(main_cat, []))
    if not sub_cat:
        await update.message.reply_text(f"❌ {main_cat}-এর এই sub-category পাওয়া যায়নি।")
        return AI_NL_ORDER
    qty = int(parsed.get("quantity", 0))
    if qty <= 0 or qty > 10000:
        await update.message.reply_text("❌ Quantity 1 থেকে 10000-এর মধ্যে হতে হবে।")
        return AI_NL_ORDER
    stock = count_items(main_cat, sub_cat)
    price = prices.get(main_cat, {}).get(sub_cat, 0)
    total = price * qty
    if stock < qty:
        await update.message.reply_text(f"❌ পর্যাপ্ত stock নেই।\nচাওয়া: {qty}\nআছে: {stock}")
        return AI_NL_ORDER
    context.user_data["order"] = {"main_cat": main_cat, "sub_cat": sub_cat, "qty": qty, "price": total}
    if main_cat in MANUAL_DELIVERY_CATEGORIES:
        await update.message.reply_text(
            f"🧾 Order প্রস্তুত\n\n📦 {main_cat} → {sub_cat}\n🔢 Quantity: {qty}\n💰 মোট: {total} টাকা\n\n"
            f"💳 Payment করুন:\n{payment_info}\n\nতারপর payment screenshot পাঠান।",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True),
        )
        return WAIT_SCREENSHOT
    user_id = update.effective_user.id
    balance = balances.get(user_id, 0)
    if balance < total:
        await update.message.reply_text(
            f"❌ ব্যালেন্স কম।\nপ্রয়োজন: {total} টাকা\nআপনার balance: {balance} টাকা\n\nআগে Deposit করুন।",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True),
        )
        context.user_data.clear()
        return ConversationHandler.END
    await update.message.reply_text(
        f"🧾 Order প্রস্তুত\n\n📦 {main_cat} → {sub_cat}\n🔢 Quantity: {qty}\n💰 মোট: {total} টাকা\n💳 আপনার balance: {balance} টাকা\n\n"
        "নিচের Confirm Purchase চাপলে order সম্পন্ন হবে।",
        reply_markup=ReplyKeyboardMarkup([
            [KeyboardButton("✅ Confirm Purchase")],
            [KeyboardButton("🔙 Back to Main Menu")]
        ], resize_keyboard=True),
    )
    return WAIT_SCREENSHOT



async def ai_report_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END
    await update.message.reply_text(
        "📈 AI Business Report\n\nকত দিনের report চান? যেমন: 7 বা 30\n\n🔙 অ্যাডমিন প্যানেল চাপলে ফিরে যাবেন।",
        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True),
    )
    return AI_REPORT


async def ai_report_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END
    text = (update.message.text or '').strip()
    if text == "🔙 অ্যাডমিন প্যানেল":
        return await back_to_admin_panel_handler(update, context)
    if not text.isdigit():
        await update.message.reply_text("❌ শুধু দিন সংখ্যা লিখুন। যেমন 7")
        return AI_REPORT
    days = max(1, min(int(text), 365))
    report = build_sales_report(days)
    await update.message.reply_text("⏳ AI report তৈরি করছে...")
    prompt = (
        "You are a business analyst for a Telegram shop. Create a concise Bangla report from the exact supplied numbers. "
        "Calculate useful ratios only when mathematically valid. Do not invent data. Give: summary, sales/deposit comparison, "
        "stock risks, best actions, and 3 practical recommendations.\n\n"
        f"REPORT DATA ({days} days):\n{json.dumps(report, ensure_ascii=False, indent=2)}\n\n"
        f"CURRENT STOCK:\n{stock_summary()}"
    )
    answer = await asyncio.to_thread(gemini_text_sync, prompt)
    for i in range(0, len(answer), 4000):
        await update.message.reply_text(answer[i:i + 4000])
    return AI_REPORT


async def ai_file_analyzer_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END
    await update.message.reply_text(
        "📊 TXT/XLSX AI Analyzer\n\nTXT বা XLSX ফাইল পাঠান। আমি data-এর summary, সমস্যা, duplicate, count, pattern ও practical recommendation দিতে পারব।\n\n"
        "⚠️ খুব বড় file হলে analysis-এর জন্য প্রথম অংশ ব্যবহার করা হবে।",
        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True),
    )
    return AI_FILE_ANALYZER


async def ai_file_analyzer_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END
    if update.message.text:
        if update.message.text == "🔙 অ্যাডমিন প্যানেল":
            return await back_to_admin_panel_handler(update, context)
        await update.message.reply_text("📎 এখন একটি .txt বা .xlsx file পাঠান।")
        return AI_FILE_ANALYZER
    if not update.message.document:
        await update.message.reply_text("❌ TXT বা XLSX file পাঠান।")
        return AI_FILE_ANALYZER
    doc = update.message.document
    filename = doc.file_name or "file"
    if not filename.lower().endswith(('.txt', '.xlsx')):
        await update.message.reply_text("❌ শুধু TXT এবং XLSX file সাপোর্ট করা হয়েছে।")
        return AI_FILE_ANALYZER
    await update.message.reply_text("⏳ File পড়ছি এবং Gemini দিয়ে analysis করছি...")
    try:
        tg_file = await context.bot.get_file(doc.file_id)
        data = await tg_file.download_as_bytearray()
        text = extract_file_text(filename, bytes(data))
        prompt = (
            "You are a data analyst for a Telegram shop admin. Analyze the supplied TXT/XLSX extracted data. "
            "Answer in simple Bangla. Do not invent values. Mention data quality issues, duplicate-looking rows, "
            "counts, useful patterns, and practical next steps.\n\n"
            f"FILE: {filename}\nDATA:\n{text}"
        )
        answer = await asyncio.to_thread(gemini_text_sync, prompt)
        for i in range(0, len(answer), 4000):
            await update.message.reply_text(answer[i:i + 4000])
    except Exception as e:
        logger.exception("AI file analysis failed")
        await update.message.reply_text(f"❌ File analysis failed: {e}")
    return AI_FILE_ANALYZER



async def admin_panel_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

    text = update.message.text

    

    if text == "🔙 Back to Main Menu":

        return await start(update, context)

    

    # NEW: ড্যাশবোর্ড রিফ্রেশ হ্যান্ডেল করা হচ্ছে

    if text == "🔄 ড্যাশবোর্ড রিফ্রেশ করুন":

        # ড্যাশবোর্ড রিফ্রেশ করতে show_dashboard কল হবে

        return await show_dashboard(update, context)

        

    # NOTE: OLD 'if text == "📊 Dashboard":' ব্লকটি এখন আর এখানে নেই।

    if text == "🤖 Gemini AI":
        return await gemini_ai_start(update, context)

    if text == "📈 AI Report":
        return await ai_report_start(update, context)

    if text == "📊 AI File Analyzer":
        return await ai_file_analyzer_start(update, context)

    if text == "📦 Sold Items":
        return await admin_download_sold_items(update, context)

    

    if text == "👥 ব্যবহারকারীর প্রোফাইল":

        return await view_user_profile(update, context)

    # --- NEW: Handle Edit User Balance ---

    if text == "✏️ Edit User Balance":

        return await edit_user_balance_start(update, context)

        

    if text == "📢 Send Notice":

        await update.message.reply_text("✍️ যে নোটিশটি পাঠাতে চান তা লিখুন:", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True))

        return SEND_NOTICE

        

    if text == "📂 Manage Categories":

        keyboard = []

        for cat in categories.keys():

            stock_count = get_total_stock(cat)

            keyboard.append([KeyboardButton(f"{cat} ({stock_count})")])

        

        keyboard.append([KeyboardButton("➕ Add Main Category")])

        keyboard.append([KeyboardButton("➖ Remove Main Category")])

        keyboard.append([KeyboardButton("🔙 অ্যাডমিন প্যানেল")])

        await update.message.reply_text("⚙️ Manage Main Categories:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))

        return MANAGE_CATEGORY

    if text == "💰 Edit Price":

        keyboard = [[KeyboardButton(cat)] for cat in categories.keys()]

        keyboard.append([KeyboardButton("🔙 অ্যাডমিন প্যানেল")])

        await update.message.reply_text("✍️ কোন Category-র মূল্য পরিবর্তন করবেন?", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))

        return EDIT_PRICE_MAIN

    if text == "💳 Edit Payment Info":

        await update.message.reply_text("✍️ নতুন পেমেন্ট তথ্য পাঠান:", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True))

        return EDIT_PAYMENT

        

    if text == "💳 Manage Payment Categories":

        return await manage_payment_categories_handler(update, context)

         

    return ADMIN_PANEL

async def manage_payment_categories_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

        

    keyboard_inline = []

    if not categories:

        await update.message.reply_text("⚠️ কোনো ক্যাটাগরি নেই।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True))

        return ADMIN_PANEL

    for main_cat in categories:

        current_status = "Manual Payment 💳" if main_cat in MANUAL_DELIVERY_CATEGORIES else "Balance Payment 💰"

        button_text = "Switch to Balance" if main_cat in MANUAL_DELIVERY_CATEGORIES else "Switch to Manual"

        callback_data = f"toggle_payment:{main_cat}"

        

        keyboard_inline.append([

            InlineKeyboardButton(f"{main_cat} ({current_status})", callback_data="ignore"),

            InlineKeyboardButton(button_text, callback_data=callback_data)

        ])

    

    reply_markup_inline = InlineKeyboardMarkup(keyboard_inline)

    

    # Send the main message with inline buttons

    await update.message.reply_text(

        "⚡️ **পেমেন্ট ক্যাটাগরি নিয়ন্ত্রণ**\n\n"

        "নিচের তালিকা থেকে প্রতিটি ক্যাটাগরির জন্য পেমেন্ট পদ্ধতি পরিবর্তন করতে পারেন।\n"

        "**Balance Payment** মানে ব্যবহারকারী তার ব্যালেন্স থেকে কিনতে পারবে।\n"

        "**Manual Payment** মানে ব্যবহারকারীকে সরাসরি পেমেন্ট করে স্ক্রিনশট পাঠাতে হবে।",

        reply_markup=reply_markup_inline,

        parse_mode='Markdown'

    )

    # Send the second message with the back button

    reply_markup_text = ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True)

    await update.message.reply_text(

        "🔙 অ্যাডমিন প্যানেল-এ ফিরে যেতে নিচের বাটনটি চাপুন।",

        reply_markup=reply_markup_text

    )

    return MANAGE_PAYMENT_CATEGORIES

async def toggle_payment_method(update: Update, context: ContextTypes.DEFAULT_TYPE):

    query = update.callback_query

    await query.answer()

    data = query.data

    

    if update.effective_user.id != ADMIN_ID:

        await query.edit_message_caption("❌ Unauthorized.", reply_markup=None)

        return

        

    if data.startswith("toggle_payment:"):

        _, cat_name = data.split(":", 1)

        

        if cat_name in MANUAL_DELIVERY_CATEGORIES:

            MANUAL_DELIVERY_CATEGORIES.remove(cat_name)

        else:

            MANUAL_DELIVERY_CATEGORIES.append(cat_name)

        

        save_user_data()

        

        # Rebuild the keyboard with the updated status

        keyboard = []

        for main_cat in categories:

            current_status = "Manual Payment 💳" if main_cat in MANUAL_DELIVERY_CATEGORIES else "Balance Payment 💰"

            button_text = "Switch to Balance" if main_cat in MANUAL_DELIVERY_CATEGORIES else "Switch to Manual"

            callback_data = f"toggle_payment:{main_cat}"

            keyboard.append([

                InlineKeyboardButton(f"{main_cat} ({current_status})", callback_data="ignore"),

                InlineKeyboardButton(button_text, callback_data=callback_data)

            ])

            

        reply_markup = InlineKeyboardMarkup(keyboard)

        await query.edit_message_text(

            "⚡️ **পেমেন্ট ক্যাটাগরি নিয়ন্ত্রণ**\n\n"

            "নিচের তালিকা থেকে প্রতিটি ক্যাটাগরির জন্য পেমেন্ট পদ্ধতি পরিবর্তন করতে পারেন।\n"

            "**Balance Payment** মানে ব্যবহারকারী তার ব্যালেন্স থেকে কিনতে পারবে।\n"

            "**Manual Payment** মানে ব্যবহারকারীকে সরাসরি পেমেন্ট করে স্ক্রিনশট পাঠাতে হবে।",

            reply_markup=reply_markup,

            parse_mode='Markdown'

        )

    # In case of "ignore" data, do nothing

    return

async def send_notice_text(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

    

    notice_text = update.message.text

    

    notice_count = 0

    failed_users = []

    

    # Get all users who have started the bot, excluding the admin

    users_to_notify = [uid for uid in user_info.keys() if uid != ADMIN_ID]

    

    if not users_to_notify:

        await update.message.reply_text("⚠️ কোনো ব্যবহারকারীকে নোটিশ পাঠানো হয়নি। সম্ভবত অন্য কোনো ব্যবহারকারী এখনও বট শুরু করেনি।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True))

        return await back_to_admin_panel_handler(update, context)

    for user_id in users_to_notify:

        try:

            await context.bot.send_message(chat_id=user_id, text=f"📢 **নোটিশ:**\n\n{notice_text}", parse_mode='Markdown')

            notice_count += 1

        except Exception:

            # Handle cases where the user has blocked the bot

            failed_users.append(user_id)

            pass

    await update.message.reply_text(f"✅ নোটিশ পাঠানো হয়েছে।\n\nসফল: {notice_count} জন ব্যবহারকারীকে\nব্যর্থ: {len(failed_users)} জন ব্যবহারকারীকে", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True))

    return await back_to_admin_panel_handler(update, context)

async def back_to_manage_main_categories_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):

    """A helper function to handle the return to the main category management view."""

    if "active_main_cat" in context.user_data:

        del context.user_data["active_main_cat"]

    if "active_sub_cat" in context.user_data:

        del context.user_data["active_sub_cat"]

    

    return await manage_category_handler(update, context)

async def manage_category_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

    text = update.message.text

    

    if text == "🔙 অ্যাডমিন প্যানেল":

        return await back_to_admin_panel_handler(update, context)

    if text == "➕ Add Main Category":

        await update.message.reply_text("✍️ নতুন প্রধান Category-র নাম পাঠান:", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True))

        return ADD_MAIN_CAT

    if text == "➖ Remove Main Category":

        keyboard = []

        for cat in categories.keys():

            stock_count = get_total_stock(cat)

            keyboard.append([KeyboardButton(f"{cat} ({stock_count})")])

        

        keyboard.append([KeyboardButton("🔙 অ্যাডমিন প্যানেল")])

        await update.message.reply_text("➖ কোন প্রধান Category remove করবেন?", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))

        return REMOVE_MAIN_CAT

    

    original_cat = text.split(" (")[0]

    

    if original_cat in categories:

        context.user_data["active_main_cat"] = original_cat

        keyboard = []

        for sub_cat in categories[original_cat]:

            stock_count = count_items(original_cat, sub_cat)

            keyboard.append([KeyboardButton(f"{sub_cat} ({stock_count})")])

        

        keyboard.append([KeyboardButton("➕ Add Sub Category")])

        keyboard.append([KeyboardButton("➖ Remove Sub Category")])

        keyboard.append([KeyboardButton("🔙 Manage Main Categories"), KeyboardButton("🔙 অ্যাডমিন প্যানেল")])

        

        await update.message.reply_text(f"⚙️ Manage Sub Categories for **{original_cat}**:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))

        return MANAGE_SUB_CATEGORY

    

    return MANAGE_CATEGORY

async def add_main_category(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

    new_cat = update.message.text.strip()

    

    if new_cat == "🔙 অ্যাডমিন প্যানেল":

        return await back_to_admin_panel_handler(update, context)

        

    if new_cat in categories:

        await update.message.reply_text("⚠️ এই প্রধান ক্যাটাগরি আগে থেকেই আছে।")

    else:

        categories[new_cat] = []

        save_user_data()

        await update.message.reply_text(f"✅ প্রধান ক্যাটাগরি '{new_cat}' যোগ হয়েছে।")

    

    return await manage_category_handler(update, context)

async def remove_main_category(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

    cat_to_remove = update.message.text.split(" (")[0].strip()

    

    if cat_to_remove == "🔙 অ্যাডমিন প্যানেল":

        return await back_to_admin_panel_handler(update, context)

        

    if cat_to_remove in categories:

        # Remove associated TXT inventory files (legacy XLSX also removed)

        if cat_to_remove in categories:

            for sub_cat in categories[cat_to_remove]:

                path = get_excel_path(cat_to_remove, sub_cat)

                if os.path.exists(path):

                    os.remove(path)

                if cat_to_remove in prices and sub_cat in prices[cat_to_remove]:

                    del prices[cat_to_remove][sub_cat]

        if cat_to_remove in prices:

            del prices[cat_to_remove]

        

        del categories[cat_to_remove]

        # Also remove from MANUAL_DELIVERY_CATEGORIES if it exists there

        if cat_to_remove in MANUAL_DELIVERY_CATEGORIES:

            MANUAL_DELIVERY_CATEGORIES.remove(cat_to_remove)

            

        save_user_data()

        await update.message.reply_text(f"✅ প্রধান ক্যাটাগরি '{cat_to_remove}' remove হয়েছে।")

    else:

        await update.message.reply_text("⚠️ এই ক্যাটাগরি পাওয়া যায়নি।")

    

    return await manage_category_handler(update, context)

    

async def manage_sub_category_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

        

    text = update.message.text

    main_cat = context.user_data.get("active_main_cat")

    

    if text == "🔙 Manage Main Categories":

        if "active_main_cat" in context.user_data:

            del context.user_data["active_main_cat"]

        return await manage_category_handler(update, context)

    

    if text == "🔙 অ্যাডমিন প্যানেল":

        return await back_to_admin_panel_handler(update, context)

    if text == "➕ Add Sub Category":

        await update.message.reply_text("✍️ নতুন সাব-ক্যাটাগরির নাম পাঠান:", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Manage Main Categories"), KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True))

        return ADD_SUB_CAT

        

    if text == "➖ Remove Sub Category":

        keyboard = [[KeyboardButton(sub_cat)] for sub_cat in categories.get(main_cat, [])]

        keyboard.append([KeyboardButton("🔙 Manage Main Categories"), KeyboardButton("🔙 অ্যাডমিন প্যানেল")])

        await update.message.reply_text("➖ কোন সাব-ক্যাটাগরি remove করবেন?", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))

        return REMOVE_SUB_CAT

    

    original_sub_cat = text.split(" (")[0]

    

    if original_sub_cat in categories.get(main_cat, []):

        context.user_data["active_sub_cat"] = original_sub_cat

        count = count_items(main_cat, original_sub_cat)

        keyboard = [

            [KeyboardButton("➕ Add Items"), KeyboardButton("📄 Add Items via TXT")],

            [KeyboardButton("🔙 Manage Main Categories"), KeyboardButton("🔙 অ্যাডমিন প্যানেল")]

        ]

        await update.message.reply_text(

            f"⚙️ Sub-Category: {original_sub_cat}\n📦 Items in stock: {count}\n\nআইটেম যোগ করতে `➕ Add Items` চাপুন।",

            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

        )

        return ADD_ITEMS

        

    return MANAGE_SUB_CATEGORY

async def add_sub_category(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

    new_sub_cat = update.message.text.strip()

    main_cat = context.user_data.get("active_main_cat")

    

    if new_sub_cat == "🔙 Manage Main Categories":

        await back_to_manage_main_categories_handler(update, context)

        return MANAGE_CATEGORY

    if new_sub_cat == "🔙 অ্যাডমিন প্যানেল":

        return await back_to_admin_panel_handler(update, context)

    if main_cat and new_sub_cat not in categories[main_cat]:

        categories[main_cat].append(new_sub_cat)

        ensure_excel(main_cat, new_sub_cat)

        save_user_data()

        await update.message.reply_text(f"✅ সাব-ক্যাটাগরি '{new_sub_cat}' যোগ হয়েছে।")

    else:

        await update.message.reply_text("⚠️ এই সাব-ক্যাটাগরি আগে থেকেই আছে অথবা কোনো প্রধান ক্যাটাগরি নির্বাচন করা হয়নি।")

    

    # After adding or handling the error, display the correct MANAGE_SUB_CATEGORY keyboard

    keyboard = []

    for sub_cat in categories.get(main_cat, []):

        stock_count = count_items(main_cat, sub_cat)

        keyboard.append([KeyboardButton(f"{sub_cat} ({stock_count})")])

    

    keyboard.append([KeyboardButton("➕ Add Sub Category")])

    keyboard.append([KeyboardButton("➖ Remove Sub Category")])

    keyboard.append([KeyboardButton("🔙 Manage Main Categories"), KeyboardButton("🔙 অ্যাডমিন প্যানেল")])

    

    await update.message.reply_text(

        f"⚙️ Manage Sub Categories for **{main_cat}**:", 

        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

    )

    

    return MANAGE_SUB_CATEGORY

    

async def remove_sub_category(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

    sub_cat_to_remove = update.message.text.split(" (")[0].strip()

    main_cat = context.user_data.get("active_main_cat")

    

    if sub_cat_to_remove == "🔙 Manage Main Categories":

        await back_to_manage_main_categories_handler(update, context)

        return MANAGE_CATEGORY

    if sub_cat_to_remove == "🔙 অ্যাডমিন প্যানেল":

        return await back_to_admin_panel_handler(update, context)

        

    if main_cat and sub_cat_to_remove in categories.get(main_cat, []):

        categories[main_cat].remove(sub_cat_to_remove)

        path = get_excel_path(main_cat, sub_cat_to_remove)

        if os.path.exists(path):

            os.remove(path)

        if main_cat in prices and sub_cat_to_remove in prices[main_cat]:

            del prices[main_cat][sub_cat_to_remove]

        save_user_data()

        await update.message.reply_text(f"✅ সাব-ক্যাটাগরি '{sub_cat_to_remove}' remove হয়েছে।")

    else:

        await update.message.reply_text("⚠️ এই সাব-ক্যাটাগরি পাওয়া যায়নি।")

    

    # After removing or handling the error, display the correct MANAGE_SUB_CATEGORY keyboard

    keyboard = []

    for sub_cat in categories.get(main_cat, []):

        stock_count = count_items(main_cat, sub_cat)

        keyboard.append([KeyboardButton(f"{sub_cat} ({stock_count})")])

    

    keyboard.append([KeyboardButton("➕ Add Sub Category")])

    keyboard.append([KeyboardButton("➖ Remove Sub Category")])

    keyboard.append([KeyboardButton("🔙 Manage Main Categories"), KeyboardButton("🔙 অ্যাডমিন প্যানেল")])

    

    await update.message.reply_text(

        f"⚙️ Manage Sub Categories for **{main_cat}**:", 

        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

    )

    

    return MANAGE_SUB_CATEGORY

async def add_item_line(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

    text = update.message.text or ""

    

    if text == "🔙 Manage Main Categories":

        await back_to_manage_main_categories_handler(update, context)

        return MANAGE_CATEGORY

    if text == "🔙 অ্যাডমিন প্যানেল":

        return await back_to_admin_panel_handler(update, context)

    

    main_cat = context.user_data.get("active_main_cat")

    sub_cat = context.user_data.get("active_sub_cat")

    

    if not main_cat or not sub_cat:

        await update.message.reply_text("⚠️ কোনো ক্যাটাগরি নির্বাচন করা হয়নি।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True))

        return await manage_category_handler(update, context)

        

    if text in ("➕ Add Items", "📄 Add Items via TXT"):

        if text == "📄 Add Items via TXT":
            await update.message.reply_text(
                f"📄 '{sub_cat}' এর জন্য TXT ফাইল পাঠান।\n\n"
                "নিয়ম: প্রতি লাইনে ১টি item থাকবে।\n"
                "উদাহরণ:\nitem-001\nitem-002\nitem-003",
                reply_markup=ReplyKeyboardMarkup(
                    [[KeyboardButton("🔙 Manage Main Categories"), KeyboardButton("🔙 অ্যাডমিন প্যানেল")]],
                    resize_keyboard=True
                )
            )
        else:
            await update.message.reply_text(
                f"✍️ '{sub_cat}' এর জন্য আইটেম পাঠান (প্রতি লাইনে একটি করে)।\n\n"
                "সব আইটেম পাঠানো হলে `✅ Done` চাপুন।",
                reply_markup=ReplyKeyboardMarkup(
                    [[KeyboardButton("📄 Add Items via TXT")],
                     [KeyboardButton("✅ Done")],
                     [KeyboardButton("🔙 Manage Main Categories"), KeyboardButton("🔙 অ্যাডমিন প্যানেল")]],
                    resize_keyboard=True
                )
            )

        return ADD_ITEMS

    # Admin can upload a TXT file directly in this state.
    if update.message.document:
        document = update.message.document
        file_name = document.file_name or ""
        if not file_name.lower().endswith(".txt"):
            await update.message.reply_text(
                "❌ শুধু .txt ফাইল পাঠান।",
                reply_markup=ReplyKeyboardMarkup(
                    [[KeyboardButton("🔙 Manage Main Categories"), KeyboardButton("🔙 অ্যাডমিন প্যানেল")]],
                    resize_keyboard=True
                )
            )
            return ADD_ITEMS

        try:
            tg_file = await document.get_file()
            file_buffer = io.BytesIO()
            await tg_file.download_to_memory(file_buffer)
            file_buffer.seek(0)
            items = add_items_from_txt_bytes(file_buffer.read())

            if not items:
                await update.message.reply_text("⚠️ TXT ফাইলে কোনো item পাওয়া যায়নি।")
                return ADD_ITEMS

            for item in items:
                add_item_to_excel(main_cat, sub_cat, item)

            await update.message.reply_text(
                f"✅ TXT থেকে {len(items)}টি item '{sub_cat}' এ যোগ হয়েছে।\n"
                f"📦 বর্তমান স্টক: {count_items(main_cat, sub_cat)}",
                reply_markup=ReplyKeyboardMarkup(
                    [[KeyboardButton("📄 Add Items via TXT")],
                     [KeyboardButton("➕ Add Items")],
                     [KeyboardButton("🔙 Manage Main Categories"), KeyboardButton("🔙 অ্যাডমিন প্যানেল")]],
                    resize_keyboard=True
                )
            )
        except Exception as e:
            logger.exception("TXT import failed")
            await update.message.reply_text(f"❌ TXT import করতে সমস্যা হয়েছে: {e}")

        return ADD_ITEMS

        

    if text == "✅ Done":

        count = count_items(main_cat, sub_cat)

        await update.message.reply_text(f"✅ '{sub_cat}' তে মোট {count} আইটেম আছে।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True))

        if "active_main_cat" in context.user_data:

            del context.user_data["active_main_cat"]

        if "active_sub_cat" in context.user_data:

            del context.user_data["active_sub_cat"]

        return await back_to_admin_panel_handler(update, context)

    

    if text:

        items_to_add = text.split('\n')

        added_count = 0

        for item in items_to_add:

            item = item.strip()

            if item:

                add_item_to_excel(main_cat, sub_cat, item)

                added_count += 1

        await update.message.reply_text(f"✅ Added {added_count} item(s).", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("✅ Done")], [KeyboardButton("🔙 Manage Main Categories"), KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True))

    

    return ADD_ITEMS

async def edit_payment_info(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

        

    global payment_info

    new_info = update.message.text.strip()

    

    if new_info == "🔙 অ্যাডমিন প্যানেল":

        return await back_to_admin_panel_handler(update, context)

        

    payment_info = new_info

    await update.message.reply_text("✅ পেমেন্ট তথ্য সফলভাবে আপডেট হয়েছে।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True))

    return await back_to_admin_panel_handler(update, context)

    

async def edit_price_main_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

    text = update.message.text.strip()

    

    if text == "🔙 অ্যাডমিন প্যানেল":

        return await back_to_admin_panel_handler(update, context)

    

    original_cat = text.split(" (")[0]

        

    if original_cat in categories.keys():

        context.user_data['temp_main_cat_for_price'] = original_cat

        keyboard = [[KeyboardButton(sub_cat)] for sub_cat in categories[original_cat]]

        keyboard.append([KeyboardButton("🔙 অ্যাডমিন প্যানেল")])

        await update.message.reply_text(f"✍️ কোন সাব-ক্যাটাগরির মূল্য পরিবর্তন করবেন?", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))

        return EDIT_PRICE_SUB

        

    await update.message.reply_text("❌ এই ক্যাটাগরি পাওয়া যায়নি।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True))

    return EDIT_PRICE_MAIN

async def edit_price_sub_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

        

    text = update.message.text.strip()

    main_cat = context.user_data.get('temp_main_cat_for_price')

    

    if text == "🔙 অ্যাডমিন প্যানেল":

        if 'temp_main_cat_for_price' in context.user_data:

            del context.user_data['temp_main_cat_for_price']

        return await back_to_admin_panel_handler(update, context)

        

    if main_cat and text in categories.get(main_cat, []):

        context.user_data['temp_sub_cat_for_price'] = text

        current_price = prices.get(main_cat, {}).get(text, "সেট করা হয়নি")

        await update.message.reply_text(f"✍️ '{text}' এর বর্তমান মূল্য: {current_price}\nনতুন মূল্য লিখুন:", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True))

        return RECEIVE_NEW_PRICE

        

    await update.message.reply_text("❌ এই সাব-ক্যাটাগরি পাওয়া যায়নি।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True))

    return EDIT_PRICE_SUB

async def receive_price(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if update.effective_user.id != ADMIN_ID:

        return ConversationHandler.END

    price_text = update.message.text.strip()

    

    if price_text == "🔙 অ্যাডমিন প্যানেল":

        if 'temp_main_cat_for_price' in context.user_data:

            del context.user_data['temp_main_cat_for_price']

        if 'temp_sub_cat_for_price' in context.user_data:

            del context.user_data['temp_sub_cat_for_price']

        return await back_to_admin_panel_handler(update, context)

    try:

        new_price = float(price_text)

        main_cat = context.user_data.get('temp_main_cat_for_price')

        sub_cat = context.user_data.get('temp_sub_cat_for_price')

        if main_cat and sub_cat:

            if main_cat not in prices:

                prices[main_cat] = {}

            prices[main_cat][sub_cat] = new_price

            save_user_data()

            await update.message.reply_text(f"✅ '{sub_cat}' এর মূল্য এখন {new_price}।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True))

            if 'temp_main_cat_for_price' in context.user_data:

                del context.user_data['temp_main_cat_for_price']

            if 'temp_sub_cat_for_price' in context.user_data:

                del context.user_data['temp_sub_cat_for_price']

            return await back_to_admin_panel_handler(update, context)

    except ValueError:

        await update.message.reply_text("❌ মূল্য শুধুমাত্র সংখ্যায় লিখুন।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 অ্যাডমিন প্যানেল")]], resize_keyboard=True))

        return RECEIVE_NEW_PRICE

        

    return await back_to_admin_panel_handler(update, context)

async def back_to_categories_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if not await check_subscription(update, context):

        return ConversationHandler.END

    keyboard = []

    for cat in categories.keys():

        keyboard.append([KeyboardButton(cat)])

    

    keyboard.append([KeyboardButton("🔙 Back to Main Menu")])

    await update.message.reply_text("🛒 ক্যাটাগরি বেছে নিন:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))

    return BUY_MENU

    

async def back_to_subcategories_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if not await check_subscription(update, context):

        return ConversationHandler.END

    

    main_cat = context.user_data.get('temp_main_cat_for_buy')

    if not main_cat:

        return await back_to_categories_handler(update, context)

        

    keyboard = []

    for sub_cat in categories[main_cat]:

        stock_count = count_items(main_cat, sub_cat)

        keyboard.append([KeyboardButton(f"{sub_cat} ({stock_count})")])

        

    keyboard.append([KeyboardButton("🔙 Back to Categories")])

    keyboard.append([KeyboardButton("🔙 Back to Main Menu")])

    await update.message.reply_text(f"🛒 **{main_cat}** এর সাব-ক্যাটাগরি বেছে নিন:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))

    return BUY_SUB_MENU

async def user_choose_category(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if not await check_subscription(update, context):

        return ConversationHandler.END

        

    text = update.message.text.strip()

    original_cat = text.split(" (")[0]

    

    if original_cat == "🔙 Back to Main Menu":

        return await start(update, context)

    context.user_data.clear()

    if original_cat in categories.keys():

        context.user_data["temp_main_cat_for_buy"] = original_cat

        

        keyboard = []

        for sub_cat in categories[original_cat]:

            stock_count = count_items(original_cat, sub_cat)

            keyboard.append([KeyboardButton(f"{sub_cat} ({stock_count})")])

        

        keyboard.append([KeyboardButton("🔙 Back to Categories")])

        keyboard.append([KeyboardButton("🔙 Back to Main Menu")])

        

        await update.message.reply_text(f"🛒 **{original_cat}** এর সাব-ক্যাটাগরি বেছে নিন:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))

        return BUY_SUB_MENU

    await update.message.reply_text("❌ এই ক্যাটাগরি পাওয়া যায়নি।")

    return BUY_MENU

async def user_choose_subcategory(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if not await check_subscription(update, context):

        return ConversationHandler.END

        

    text = update.message.text.strip()

    original_sub_cat = text.split(" (")[0]

    main_cat = context.user_data.get('temp_main_cat_for_buy')

    

    if text == "🔙 Back to Categories":

        return await back_to_categories_handler(update, context)

    if text == "🔙 Back to Main Menu":

        return await start(update, context)

        

    if main_cat and original_sub_cat in categories.get(main_cat, []):

        context.user_data["order"] = {"main_cat": main_cat, "sub_cat": original_sub_cat}

        

        price = prices.get(main_cat, {}).get(original_sub_cat, "মূল্য এখনো সেট করা হয়নি।")

        

        keyboard = [

            [KeyboardButton("🔙 Back to Sub-Categories")],

            [KeyboardButton("🔙 Back to Main Menu")]

        ]

        

        await update.message.reply_text(f"✅ আপনি **{original_sub_cat}** সিলেক্ট করেছেন।\n💰 **প্রতিটির দাম:** {price} টাকা\n\n✍️ অনুগ্রহ করে আপনি কতগুলি চান তা সংখ্যায় লিখুন:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))

        return GET_QUANTITY

    await update.message.reply_text("❌ এই সাব-ক্যাটাগরি পাওয়া যায়নি।")

    return BUY_SUB_MENU

async def receive_quantity(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if not await check_subscription(update, context):

        return ConversationHandler.END

    qty = update.message.text.strip()

    if qty == "🔙 Back to Sub-Categories":

        return await back_to_subcategories_handler(update, context)

    if qty == "🔙 Back to Main Menu":

        return await start(update, context)

    

    if not qty.isdigit():

        await update.message.reply_text("❌ অনুগ্রহ করে শুধু সংখ্যা লিখুন।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Sub-Categories"), KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True))

        return GET_QUANTITY

        

    qty = int(qty)

    order = context.user_data.get("order", {})

    order["qty"] = qty

    

    main_cat = order['main_cat']

    sub_cat = order['sub_cat']

    

    price_per_item = prices.get(main_cat, {}).get(sub_cat, 0)

    total_price = price_per_item * qty

    

    order["price"] = total_price

    context.user_data["order"] = order

    is_manual = main_cat in MANUAL_DELIVERY_CATEGORIES

    if not is_manual:

        user_id = update.effective_user.id

        current_balance = balances.get(user_id, 0)

    

        if current_balance >= total_price:

            keyboard = [

                [KeyboardButton("✅ Confirm Purchase")],

                [KeyboardButton("🔙 Back to Sub-Categories")],

                [KeyboardButton("🔙 Back to Main Menu")]

            ]

            await update.message.reply_text(

                f"✅ অর্ডার তৈরি হয়েছে।\n"

                f"ক্যাটাগরি: {order['sub_cat']}\n"

                f"পরিমাণ: {order['qty']}\n"

                f"মোট দাম: {total_price} টাকা\n"

                f"আপনার বর্তমান ব্যালেন্স: {current_balance} টাকা\n\n"

                f"আপনার ব্যালেন্স থেকে পেমেন্ট করতে `✅ Confirm Purchase` চাপুন।",

                reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

            )

        else:

            await update.message.reply_text(f"❌ আপনার যথেষ্ট ব্যালেন্স নেই। আপনার প্রয়োজন {total_price} টাকা কিন্তু ব্যালেন্স আছে {current_balance} টাকা।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Sub-Categories"), KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True))

            return BUY_SUB_MENU

            

        return WAIT_SCREENSHOT

    

    keyboard = [

        [KeyboardButton("🔙 Back to Sub-Categories")],

        [KeyboardButton("🔙 Back to Main Menu")]

    ]

    await update.message.reply_text(

        f"✅ অর্ডার তৈরি হয়েছে।\n"

        f"ক্যাটাগরি: {order['sub_cat']}\n"

        f"পরিমাণ: {order['qty']}\n"

        f"মোট দাম: {total_price} টাকা\n"

        f"⚠️ অনুগ্রহ করে পেমেন্ট করুন:\n{payment_info}\n\n"

        f"📸 পেমেন্টের পর স্ক্রিনশট পাঠান।",

        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

    )

    return WAIT_SCREENSHOT

async def user_send_screenshot(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if not await check_subscription(update, context):

        return ConversationHandler.END

    if update.message.text == "🔙 Back to Sub-Categories":

        return await back_to_subcategories_handler(update, context)

    if update.message.text == "🔙 Back to Main Menu":

        return await start(update, context)

    if update.message.text == "🔙 অ্যাডমিন প্যানেল":

        return await back_to_admin_panel_handler(update, context)

    order = context.user_data.get("order", {})

    if not order:

        return ConversationHandler.END

        

    main_cat = order['main_cat']

    sub_cat = order['sub_cat']

    if update.message.text == "✅ Confirm Purchase" and main_cat not in MANUAL_DELIVERY_CATEGORIES:

        user_id = update.effective_user.id

        total_price = order['price']

        qty = order['qty']

        current_balance = balances.get(user_id, 0)

        

        if current_balance >= total_price:

            balances[user_id] = current_balance - total_price

            

            items = pop_items_from_excel(main_cat, sub_cat, qty)

            if not items:

                balances[user_id] = balances.get(user_id, 0) + total_price 

                await update.message.reply_text("❌ যথেষ্ট আইটেম স্টক এ নেই। আপনার ব্যালেন্স রিফান্ড করা হয়েছে।", reply_markup=ReplyKeyboardRemove())

                return await start(update, context)

            

            record_sold_items(items, main_cat, sub_cat, user_id, prices.get(main_cat, {}).get(sub_cat, 0))

            global total_sales, sales_count_per_category, transaction_log, user_sales

            total_sales += total_price

            sales_count_per_category[sub_cat] = sales_count_per_category.get(sub_cat, 0) + qty

            transaction_log.append(('sale', user_id, main_cat, sub_cat, qty, total_price, time.time()))

            user_sales[user_id] = user_sales.get(user_id, 0) + total_price

            await send_items_xlsx(context, user_id, items, main_cat, sub_cat, qty)

            await update.message.reply_text(
                "✅ ব্যালেন্স ব্যবহার করে আপনার অর্ডার সফল হয়েছে। XLSX ফাইল আপনার কাছে পাঠানো হয়েছে!",
                reply_markup=ReplyKeyboardRemove()
            )

            context.user_data.clear()

            save_user_data()

            return await start(update, context)

        else:

            await update.message.reply_text("❌ আপনার যথেষ্ট ব্যালেন্স নেই।", reply_markup=ReplyKeyboardRemove())

            return await start(update, context)

    # For manual payment via screenshot

    if not update.message.photo:

        await update.message.reply_text("❌ অনুগ্রহ করে শুধু একটি ছবি পাঠান।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Sub-Categories"), KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True))

        return WAIT_SCREENSHOT

        

    user = update.effective_user

    username = user.username if user.username else 'N/A'

    caption = (

        f"🔔 **নতুন অর্ডার!** 🔔\n"

        f"User: @{username}\n"

        f"Category: {order['sub_cat']}\n"

        f"Quantity: {order['qty']}\n"

        f"Price: {order['price']}\n"

        f"UserID: {user.id}"

    )

    keyboard = InlineKeyboardMarkup([

        [InlineKeyboardButton("✅ Confirm Order", callback_data=f"confirm_manual:{user.id}:{order['main_cat']}:{order['sub_cat']}:{order['qty']}:{order['price']}"),

         InlineKeyboardButton("❌ Cancel Order", callback_data=f"cancel_manual:{user.id}")]

    ])

    

    await context.bot.send_photo(chat_id=ADMIN_ID, photo=update.message.photo[-1].file_id, caption=caption, reply_markup=keyboard)

    

    await update.message.reply_text("✅ Screenshot পাঠানো হয়েছে Admin-কে।")

    await start(update, context)

    context.user_data.clear()

    return ConversationHandler.END

async def admin_order_action(update: Update, context: ContextTypes.DEFAULT_TYPE):

    global total_sales, sales_count_per_category, transaction_log, user_sales, prices

    query = update.callback_query

    await query.answer()

    data = query.data

    

    if update.effective_user.id != ADMIN_ID:

        await query.edit_message_caption("❌ Unauthorized.", reply_markup=None)

        return

    

    if data.startswith("confirm:"): # Auto-delivery via balance

        parts = data.split(":")

        uid = int(parts[1])

        main_cat = parts[2]

        sub_cat = parts[3]

        qty = int(parts[4])

        items = pop_items_from_excel(main_cat, sub_cat, qty)

        

        if not items:

            await query.edit_message_caption(query.message.caption + "\n\n❌ Not enough stock.", reply_markup=None)

            return

        price_per_item = prices.get(main_cat, {}).get(sub_cat, 0)
        record_sold_items(items, main_cat, sub_cat, uid, price_per_item)

        total_price = price_per_item * qty

        total_sales += total_price

        sales_count_per_category[sub_cat] = sales_count_per_category.get(sub_cat, 0) + qty

        transaction_log.append(('sale', uid, main_cat, sub_cat, qty, total_price, time.time()))

        user_sales[uid] = user_sales.get(uid, 0) + total_price

        save_user_data()

        await send_items_xlsx(context, uid, items, main_cat, sub_cat, qty)

        await check_low_stock_after_sale(context.application, main_cat, sub_cat)

        await query.edit_message_caption(
            query.message.caption + "\n\n✅ Delivered (XLSX File Sent)",
            reply_markup=None
        )

    

    elif data.startswith("confirm_manual:"):

        parts = data.split(":")

        uid = int(parts[1])

        main_cat = parts[2]

        sub_cat = parts[3]

        qty = int(parts[4])

        total_price = float(parts[5])

        

        # Check stock first

        stock_count = count_items(main_cat, sub_cat)

        if stock_count < qty:

            # Not enough stock, show override option

            keyboard = InlineKeyboardMarkup([

                [InlineKeyboardButton("✅ Force Confirm", callback_data=f"force_confirm:{uid}:{main_cat}:{sub_cat}:{qty}:{total_price}"),

                 InlineKeyboardButton("❌ Cancel", callback_data=f"cancel_manual:{uid}")]

            ])

            await query.edit_message_caption(query.message.caption + "\n\n⚠️ স্টকে পর্যাপ্ত আইটেম নেই। আপনি কি নিশ্চিত করতে চান?\n\nযদি নিশ্চিত করেন, আপনাকে ম্যানুয়ালি আইটেমটি পাঠাতে হবে।", reply_markup=keyboard)

            return

            

        # Manually pop items from excel

        items = pop_items_from_excel(main_cat, sub_cat, qty)
        if not items:
            await query.edit_message_caption(query.message.caption + "\n\n❌ Stock changed before delivery. Order was not completed.", reply_markup=None)
            return

        price_per_item = prices.get(main_cat, {}).get(sub_cat, 0)
        record_sold_items(items, main_cat, sub_cat, uid, price_per_item)

        total_sales += total_price

        sales_count_per_category[sub_cat] = sales_count_per_category.get(sub_cat, 0) + qty

        transaction_log.append(('sale', uid, main_cat, sub_cat, qty, total_price, time.time()))

        user_sales[uid] = user_sales.get(uid, 0) + total_price

        save_user_data()

        await send_items_xlsx(context, uid, items, main_cat, sub_cat, qty)

        await check_low_stock_after_sale(context.application, main_cat, sub_cat)

        await query.edit_message_caption(
            query.message.caption + f"\n\n✅ Confirmed by Admin. XLSX file sent to user.",
            reply_markup=None
        )

    elif data.startswith("force_confirm:"):

        parts = data.split(":")

        uid = int(parts[1])

        main_cat = parts[2]

        sub_cat = parts[3]

        qty = int(parts[4])

        total_price = float(parts[5])

        # Force confirm even without stock

        total_sales += total_price

        sales_count_per_category[sub_cat] = sales_count_per_category.get(sub_cat, 0) + qty

        transaction_log.append(('sale', uid, main_cat, sub_cat, qty, total_price, time.time()))

        user_sales[uid] = user_sales.get(uid, 0) + total_price

        save_user_data()

        

        # Admin must manually deliver, so no items are popped.

        await context.bot.send_message(

            chat_id=uid, 

            text="✅ আপনার অর্ডার নিশ্চিত করা হয়েছে। অ্যাডমিন শীঘ্রই আপনাকে বিস্তারিত তথ্য পাঠাবেন।"

        )

        await query.edit_message_caption(query.message.caption + f"\n\n✅ Force Confirmed by Admin. User will be contacted manually.", reply_markup=None)

    elif data.startswith("cancel_manual:"):

        _, uid = data.split(":")

        

        await context.bot.send_message(

            chat_id=uid,

            text="❌ দুঃখিত, আপনার অর্ডারটি বাতিল করা হয়েছে।"

        )

        await query.edit_message_caption(query.message.caption + "\n\n❌ Cancelled by Admin", reply_markup=None)

        

    else: # Default cancel

        _, uid = data.split(":")

        

        await context.bot.send_message(

            chat_id=uid,

            text="❌ দুঃখিত, আপনার অর্ডারটি বাতিল করা হয়েছে।"

        )

        

        await query.edit_message_caption(query.message.caption + "\n\n❌ Cancelled by Admin", reply_markup=None)

async def admin_deposit_action(update: Update, context: ContextTypes.DEFAULT_TYPE):

    global total_deposits, transaction_log, user_deposits, balances

    query = update.callback_query

    await query.answer() # বাটন ক্লিকের লোডিং স্টেট সরিয়ে দেয়

    data = query.data

    # শুধুমাত্র অ্যাডমিনই এই অ্যাকশনটি করতে পারবে

    if update.effective_user.id != ADMIN_ID:

        try:

            # Unauthorized হলে বাটন রিমুভ করে দেওয়া উচিত

            await query.edit_message_caption("❌ Unauthorized.", reply_markup=None)

        except:

            pass

        return

    

    # === ডিপোজিট নিশ্চিতকরণ (Confirm) ব্লক ===

    if data.startswith("deposit_confirm:"):

        parts = data.split(":")

        uid = int(parts[1])

        amount = float(parts[2])

        

        # ব্যালেন্স আপডেট করা হলো

        balances[uid] = balances.get(uid, 0) + amount

        

        # অন্যান্য ডেটা আপডেট ও সেভ করা হলো

        total_deposits += amount

        transaction_log.append(('deposit', uid, amount, time.time()))

        user_deposits[uid] = user_deposits.get(uid, 0) + amount

        save_user_data()

        # ব্যবহারকারীকে সফলতার মেসেজ

        await context.bot.send_message(

            chat_id=uid,

            text=f"✅ আপনার ডিপোজিট সফল হয়েছে। আপনার ব্যালেন্সে {amount} টাকা যোগ হয়েছে।\n"

                 f"নতুন ব্যালেন্স: {balances[uid]} টাকা।"

        )

        

        # অ্যাডমিনের মেসেজ আপডেট করা হলো (নতুন মোট ব্যালেন্স সহ)

        try:

            await query.edit_message_caption(

                # ক্যাপশনের শেষে নতুন টেক্সট যোগ করা হচ্ছে

                query.message.caption + 

                f"\n\n✅ নিশ্চিত করা হয়েছে। {amount} টাকা ইউজার <code>{uid}</code> এর ব্যালেন্সে যোগ করা হয়েছে। "

                f"<b>বর্তমান মোট ব্যালেন্স: {balances[uid]} টাকা।</b>", 

                reply_markup=None,

                parse_mode='HTML' # ত্রুটি এড়াতে এবং বোল্ড করার জন্য HTML ব্যবহার করা হয়েছে

            )

        except Exception as e:

            logger.error(f"Failed to edit message caption on deposit confirmation: {e}")

            # ক্যাপশন আপডেট না হলে অ্যাডমিনকে বিকল্প মেসেজ পাঠানো হবে

            await context.bot.send_message(

                chat_id=ADMIN_ID, 

                text=f"⚠️ Failed to update deposit confirmation message. User ID: {uid}, Amount: {amount}"

            )

            

    # === ডিপোজিট বাতিলকরণ (Cancel) ব্লক ===

    else: # deposit_cancel:

        _, uid = data.split(":")

        uid = int(uid)

        

        # ব্যবহারকারীকে বাতিলের মেসেজ

        await context.bot.send_message(

            chat_id=uid,

            text="❌ দুঃখিত, আপনার ডিপোজিট রিকোয়েস্ট বাতিল করা হয়েছে।"

        )

        

        # অ্যাডমিনের মেসেজ আপডেট করা হলো

        try:

            await query.edit_message_caption(

                query.message.caption + "\n\n❌ Deposit cancelled by Admin", 

                reply_markup=None

            )

        except Exception as e:

            logger.error(f"Failed to edit message caption on deposit cancellation: {e}")

            await context.bot.send_message(

                chat_id=ADMIN_ID, 

                text=f"⚠️ Failed to update deposit cancellation message. User ID: {uid}"

            )

# ================ MAIN =================

def main():
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN environment variable is not set.")

    app = (
        ApplicationBuilder()
        .token(BOT_TOKEN)
        .post_init(post_init)
        .post_shutdown(post_shutdown)
        .build()
    )

    conv = ConversationHandler(

        entry_points=[

            CommandHandler("start", start),

            MessageHandler(filters.TEXT & ~filters.COMMAND, menu_handler)

        ],

        states={

            BUY_MENU: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), start),

                MessageHandler(filters.TEXT & ~filters.COMMAND, user_choose_category)

            ],

            BUY_SUB_MENU: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Categories"), back_to_categories_handler),

                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), start),

                MessageHandler(filters.TEXT & ~filters.COMMAND, user_choose_subcategory)

            ],

            GET_QUANTITY: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Sub-Categories"), back_to_subcategories_handler),

                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), start),

                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_quantity)

            ],

            WAIT_SCREENSHOT: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Sub-Categories"), back_to_subcategories_handler),

                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), start),

                MessageHandler(filters.TEXT & filters.Regex("✅ Confirm Purchase"), user_send_screenshot),

                MessageHandler(filters.PHOTO | filters.TEXT, user_send_screenshot),

            ],

            ADMIN_PANEL: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), start),

                MessageHandler(filters.TEXT & ~filters.COMMAND, admin_panel_handler)

            ],

            MANAGE_CATEGORY: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 অ্যাডমিন প্যানেল"), back_to_admin_panel_handler),

                MessageHandler(filters.TEXT & ~filters.COMMAND, manage_category_handler)

            ],

            MANAGE_SUB_CATEGORY: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 অ্যাডমিন প্যানেল"), back_to_admin_panel_handler),

                MessageHandler(filters.TEXT & filters.Regex("🔙 Manage Main Categories"), back_to_manage_main_categories_handler),

                MessageHandler(filters.TEXT & ~filters.COMMAND, manage_sub_category_handler)

            ],

            ADD_MAIN_CAT: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 অ্যাডমিন প্যানেল"), back_to_admin_panel_handler),

                MessageHandler(filters.TEXT & ~filters.COMMAND, add_main_category)

            ],

            REMOVE_MAIN_CAT: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 অ্যাডমিন প্যানেল"), back_to_admin_panel_handler),

                MessageHandler(filters.TEXT & ~filters.COMMAND, remove_main_category)

            ],

            ADD_SUB_CAT: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 অ্যাডমিন প্যানেল"), back_to_admin_panel_handler),

                MessageHandler(filters.TEXT & filters.Regex("🔙 Manage Main Categories"), back_to_manage_main_categories_handler),

                MessageHandler(filters.TEXT & ~filters.COMMAND, add_sub_category)

            ],

            REMOVE_SUB_CAT: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 অ্যাডমিন প্যানেল"), back_to_admin_panel_handler),

                MessageHandler(filters.TEXT & filters.Regex("🔙 Manage Main Categories"), back_to_manage_main_categories_handler),

                MessageHandler(filters.TEXT & ~filters.COMMAND, remove_sub_category)

            ],

            ADD_ITEMS: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 অ্যাডমিন প্যানেল"), back_to_admin_panel_handler),

                MessageHandler(filters.TEXT & filters.Regex("🔙 Manage Main Categories"), back_to_manage_main_categories_handler),

                MessageHandler(filters.Document.ALL, add_item_line),

                MessageHandler(filters.TEXT & ~filters.COMMAND, add_item_line)

            ],

            EDIT_PAYMENT: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 অ্যাডমিন প্যানেল"), back_to_admin_panel_handler),

                MessageHandler(filters.TEXT & ~filters.COMMAND, edit_payment_info)

            ],

            

            EDIT_PRICE_MAIN: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 অ্যাডমিন প্যানেল"), back_to_admin_panel_handler),

                MessageHandler(filters.TEXT & ~filters.COMMAND, edit_price_main_handler)

            ],

            EDIT_PRICE_SUB: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 অ্যাডমিন প্যানেল"), back_to_admin_panel_handler),

                MessageHandler(filters.TEXT & ~filters.COMMAND, edit_price_sub_handler)

            ],

            RECEIVE_NEW_PRICE: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 অ্যাডমিন প্যানেল"), back_to_admin_panel_handler),

                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_price)

            ],

            DEPOSIT: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), start),

                MessageHandler(filters.TEXT & ~filters.COMMAND, deposit_handler),

            ],

            GET_DEPOSIT_AMOUNT: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), start),

                MessageHandler(filters.PHOTO, receive_deposit_screenshot)

            ],

            DASHBOARD: [

                MessageHandler(filters.TEXT & filters.Regex("🔄 ড্যাশবোর্ড রিফ্রেশ করুন"), handle_dashboard_refresh),

                MessageHandler(filters.TEXT & filters.Regex("🔙 অ্যাডমিন প্যানেল"), back_to_admin_panel_handler)

            ],

            SEND_NOTICE: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 অ্যাডমিন প্যানেল"), back_to_admin_panel_handler),

                MessageHandler(filters.TEXT & ~filters.COMMAND, send_notice_text)

            ],

            SEARCH_USER_PROFILE: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 অ্যাডমিন প্যানেল"), back_to_admin_panel_handler),

                MessageHandler(filters.TEXT & ~filters.COMMAND, search_and_show_user_profile)

            ],

            MANAGE_PAYMENT_CATEGORIES: [

                CallbackQueryHandler(toggle_payment_method),

                MessageHandler(filters.TEXT & filters.Regex("🔙 অ্যাডমিন প্যানেল"), back_to_admin_panel_handler)

            ],

            # --- NEW STATES HANDLERS ---

            SEARCH_USER_FOR_BALANCE: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 অ্যাডমিন প্যানেল"), back_to_admin_panel_handler),

                MessageHandler(filters.TEXT & ~filters.COMMAND, search_user_for_balance)

            ],

            BALANCE_EDIT_ACTION: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 অ্যাডমিন প্যানেল"), back_to_admin_panel_handler),

                MessageHandler(filters.TEXT & ~filters.COMMAND, balance_edit_action_handler)

            ],

            RECEIVE_BALANCE_EDIT_AMOUNT: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 অ্যাডমিন প্যানেল"), back_to_admin_panel_handler),

                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_balance_edit_amount)

            ],

            AI_SUPPORT: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), start),
                MessageHandler(filters.TEXT & ~filters.COMMAND, ai_support_handler)
            ],

            AI_STOCK_SEARCH: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), start),
                MessageHandler(filters.TEXT & ~filters.COMMAND, ai_stock_search_handler)
            ],

            AI_NL_ORDER: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), start),
                MessageHandler(filters.TEXT & ~filters.COMMAND, ai_nl_order_handler)
            ],

            AI_REPORT: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 অ্যাডমিন প্যানেল"), back_to_admin_panel_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, ai_report_handler)
            ],

            AI_FILE_ANALYZER: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 অ্যাডমিন প্যানেল"), back_to_admin_panel_handler),
                MessageHandler(filters.Document.ALL, ai_file_analyzer_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, ai_file_analyzer_handler)
            ],

            AI_ASSISTANT: [

                MessageHandler(filters.TEXT & filters.Regex("🔙 অ্যাডমিন প্যানেল"), back_to_admin_panel_handler),

                MessageHandler(filters.TEXT & ~filters.COMMAND, gemini_ai_handler)

            ]

        },

        fallbacks=[CommandHandler("start", start)]

    )

    

    app.add_handler(conv)

    app.add_handler(CallbackQueryHandler(admin_sold_category_callback, pattern="^sold(?:main|sub|back|close)"))
    app.add_handler(CallbackQueryHandler(admin_order_action, pattern="^(confirm:|cancel:|confirm_manual:|cancel_manual:|force_confirm:)"))

    app.add_handler(CallbackQueryHandler(admin_deposit_action, pattern="^(deposit_confirm:|deposit_cancel:)"))

    app.add_handler(CallbackQueryHandler(toggle_payment_method, pattern="^toggle_payment:"))

    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND & filters.Regex("🔙 Back to Main Menu"), start))

    

    logger.info("🤖 Bot running...")

    app.run_polling()

if __name__ == "__main__":

    main()
