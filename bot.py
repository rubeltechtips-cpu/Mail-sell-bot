import sys
import types

if sys.version_info >= (3, 13):
    if 'imghdr' not in sys.modules:
        imghdr = types.ModuleType('imghdr')
        imghdr.what = lambda f, h=None: None
        sys.modules['imghdr'] = imghdr

import logging
import os
import io
import time
import json
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler

from telegram import (
    Update, ReplyKeyboardMarkup, KeyboardButton, InputFile,
    ReplyKeyboardRemove, InlineKeyboardMarkup, InlineKeyboardButton,
    ChatMember
)
from telegram.ext import (
    Updater, CommandHandler, MessageHandler,
    Filters, CallbackContext, ConversationHandler, CallbackQueryHandler
)
from openpyxl import Workbook

# ================ HTTP SERVER ================
class HealthCheckHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path == '/':
            self.send_response(200)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            self.wfile.write(b'<h1>Bot is running!</h1>')
        elif self.path == '/health':
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            self.wfile.write(b'{"status": "ok"}')
        else:
            self.send_response(404)
            self.end_headers()
    
    def log_message(self, format, *args):
        pass

def run_http_server():
    try:
        port = int(os.environ.get('PORT', 8000))
        server = HTTPServer(('0.0.0.0', port), HealthCheckHandler)
        logging.info(f"HTTP Server running on port {port}")
        server.serve_forever()
    except Exception as e:
        logging.error(f"HTTP Server error: {e}")

threading.Thread(target=run_http_server, daemon=True).start()

# ================ CONFIG ================
BOT_TOKEN = "8349208659:AAEyJikjx1tUri_PztFGRca_lPT0WilJ0N0"
ADMIN_ID = 8061006207
ADMIN_USERNAME = "Rubel_QSB"
CHANNEL_USERNAME = "quick_sell_bd"
DATA_DIR = "categories"
TXT_DIR = "txt_files"
EXCEL_DIR = "excel_files"

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(TXT_DIR, exist_ok=True)
os.makedirs(EXCEL_DIR, exist_ok=True)

# ================ STATES ================
(
    MAIN_MENU, BUY_MENU, BUY_SUB_MENU, ADMIN_PANEL,
    ADD_MAIN_CAT, REMOVE_MAIN_CAT, MANAGE_CATEGORY, MANAGE_SUB_CATEGORY,
    ADD_SUB_CAT, REMOVE_SUB_CAT, ADD_ITEMS_TXT, EDIT_PAYMENT,
    EDIT_PRICE_MAIN, EDIT_PRICE_SUB, RECEIVE_NEW_PRICE, GET_QUANTITY,
    WAIT_SCREENSHOT, DEPOSIT, GET_DEPOSIT_AMOUNT, DASHBOARD,
    SEND_NOTICE, VIEW_USER_PROFILE, SEARCH_USER_PROFILE,
    MANAGE_PAYMENT_CATEGORIES, SEARCH_USER_FOR_BALANCE,
    BALANCE_EDIT_ACTION, RECEIVE_BALANCE_EDIT_AMOUNT, CONFIRM_ORDER
) = range(28)

logging.basicConfig(format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ================ DATA ================
categories = {}
prices = {}
payment_info = "বিকাশ: 017XXXXXXXX\nনগদ: 018XXXXXXXX\nবিন্যান্স: yourmail@gmail.com"
balances = {}
user_sales = {}
user_deposits = {}
user_info = {}
total_deposits = 0
total_sales = 0
sales_count_per_category = {}
transaction_log = []
dashboard_message = "স্বাগতম! এটি আপনার বটের ড্যাশবোর্ড।"
MANUAL_DELIVERY_CATEGORIES = []

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
        logger.info("No user_data.json found. Starting fresh.")

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
def get_txt_path(main_cat, sub_cat):
    file_name = f"{main_cat}_{sub_cat}.txt".replace(" ", "_").replace("-", "_")
    return os.path.join(TXT_DIR, file_name)

def get_excel_path(main_cat, sub_cat):
    file_name = f"{main_cat}_{sub_cat}.xlsx".replace(" ", "_").replace("-", "_")
    return os.path.join(EXCEL_DIR, file_name)

def ensure_txt_file(main_cat, sub_cat):
    path = get_txt_path(main_cat, sub_cat)
    if not os.path.exists(path):
        with open(path, 'w', encoding='utf-8') as f:
            f.write("")

def add_items_from_txt(main_cat, sub_cat, txt_content):
    ensure_txt_file(main_cat, sub_cat)
    path = get_txt_path(main_cat, sub_cat)
    with open(path, 'a', encoding='utf-8') as f:
        f.write(txt_content + '\n')

def pop_items_from_txt(main_cat, sub_cat, qty):
    path = get_txt_path(main_cat, sub_cat)
    if not os.path.exists(path):
        return []
    with open(path, 'r', encoding='utf-8') as f:
        items = [line.strip() for line in f.readlines() if line.strip()]
    if len(items) < qty:
        return []
    result = items[:qty]
    remaining_items = items[qty:]
    with open(path, 'w', encoding='utf-8') as f:
        for item in remaining_items:
            f.write(item + '\n')
    return result

def count_items(main_cat, sub_cat):
    path = get_txt_path(main_cat, sub_cat)
    if not os.path.exists(path):
        return 0
    with open(path, 'r', encoding='utf-8') as f:
        items = [line.strip() for line in f.readlines() if line.strip()]
    return len(items)

def create_xlsx_file(items, file_name):
    wb = Workbook()
    ws = wb.active
    ws.title = "Items"
    for item in items:
        ws.append([item])
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def get_total_stock(main_cat):
    total = 0
    if main_cat in categories:
        for sub_cat in categories[main_cat]:
            total += count_items(main_cat, sub_cat)
    return total

def get_report_summary(transactions, days):
    end_timestamp = time.time()
    start_timestamp = end_timestamp - (days * 24 * 60 * 60)
    daily_deposits = 0
    daily_sales = 0
    for trans_type, _, amount, timestamp in transactions:
        if start_timestamp <= timestamp <= end_timestamp:
            if trans_type == 'deposit':
                daily_deposits += amount
            elif trans_type == 'sale':
                daily_sales += amount
    return daily_deposits, daily_sales

def get_user_transactions(user_id, transactions):
    return [t for t in transactions if t[1] == user_id]

# ================ CHECK SUBSCRIPTION ================
def check_subscription(update, context):
    user_id = update.effective_user.id
    try:
        chat_member = context.bot.get_chat_member(chat_id=f"@{CHANNEL_USERNAME}", user_id=user_id)
        if chat_member.status in [ChatMember.MEMBER, ChatMember.ADMINISTRATOR, ChatMember.CREATOR]:
            return True
        else:
            keyboard = InlineKeyboardMarkup([
                [InlineKeyboardButton("Join Channel", url=f"https://t.me/{CHANNEL_USERNAME}")]
            ])
            update.message.reply_text(
                "❌ আপনি এখনো আমাদের চ্যানেলে জয়েন করেননি।\n"
                "বট ব্যবহার করার জন্য অনুগ্রহ করে নিচের বাটনে ক্লিক করে চ্যানেলে জয়েন করুন।",
                reply_markup=keyboard
            )
            return False
    except Exception as e:
        logger.error(f"Error checking subscription: {e}")
        update.message.reply_text("চ্যানেল সদস্যতা পরীক্ষা করতে সমস্যা হচ্ছে।")
        return False

# ================ START ================
def start(update, context):
    if not check_subscription(update, context):
        return ConversationHandler.END
    
    user_id = update.effective_user.id
    username = update.effective_user.username
    first_name = update.effective_user.first_name
    
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
    ]
    if update.effective_user.id == ADMIN_ID:
        keyboard.append([KeyboardButton("⚙️ Admin Panel")])
    
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)
    update.message.reply_text(f"👋 স্বাগতম! আপনার বর্তমান ব্যালেন্স: {current_balance} টাকা।", reply_markup=reply_markup)
    return MAIN_MENU

def menu_handler(update, context):
    if not check_subscription(update, context):
        return ConversationHandler.END

    text = update.message.text
    
    if text == "🛒 Buy":
        if not categories:
            update.message.reply_text("⚠️ এখন কোনো ক্যাটাগরি নেই।")
            return MAIN_MENU
        keyboard = []
        for cat in categories.keys():
            keyboard.append([KeyboardButton(cat)])
        keyboard.append([KeyboardButton("🔙 Back to Main Menu")])
        update.message.reply_text("🛒 ক্যাটাগরি বেছে নিন:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
        return BUY_MENU

    if text == "💰 Balance":
        user_id = update.effective_user.id
        current_balance = balances.get(user_id, 0)
        update.message.reply_text(f"আপনার বর্তমান ব্যালেন্স: {current_balance} টাকা।")
        return MAIN_MENU

    if text == "💸 Deposit":
        update.message.reply_text("আপনি কত টাকা ডিপোজিট করতে চান তা লিখুন:", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True))
        return DEPOSIT

    if text == "📞 Help":
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("📞 Contact Admin", url=f"tg://user?id={ADMIN_ID}")]
        ])
        update.message.reply_text("📞 অ্যাডমিনের সাথে যোগাযোগ করতে নিচের বাটনে ক্লিক করুন।", reply_markup=keyboard)
        return MAIN_MENU

    if text == "⚙️ Admin Panel":
        if update.effective_user.id == ADMIN_ID:
            return show_dashboard(update, context)
        else:
            update.message.reply_text("❌ অননুমোদিত।")
            return MAIN_MENU
            
    if text == "🔙 Back to Main Menu":
        return start(update, context)
    
    return MAIN_MENU

# ================ DASHBOARD ================
def show_dashboard(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU

    global total_deposits, total_sales, balances, sales_count_per_category, user_info, transaction_log, dashboard_message

    total_users_count = len(user_info)
    
    stock_info = ""
    for main_cat, sub_cats in categories.items():
        stock_info += f"  - {main_cat}\n" 
        for sub_cat in sub_cats:
            count = count_items(main_cat, sub_cat)
            stock_info += f"  - {sub_cat}: {count} টি আইটেম\n"

    sorted_sales = sorted(sales_count_per_category.items(), key=lambda item: item[1], reverse=True)
    top_selling_info = ""
    for sub_cat, count in sorted_sales[:10]:
        top_selling_info += f"  - {sub_cat}: {count} বিক্রয়\n"
    if len(sorted_sales) > 10:
        top_selling_info += f"  - ... এবং আরও {len(sorted_sales) - 10}টি ক্যাটাগরি"

    recent_transactions = ""
    last_5_transactions = transaction_log[-5:]
    if last_5_transactions:
        for trans in reversed(last_5_transactions):
            trans_type, user_id, amount, timestamp = trans
            date_str = time.strftime('%H:%M %b %d', time.localtime(timestamp))
            user_data = user_info.get(user_id, {})
            username = user_data.get("username", "N/A")
            if trans_type == 'deposit':
                recent_transactions += f"  - 💸 ডিপোজিট: {amount} টাকা (@{username}) {date_str}\n" 
            elif trans_type == 'sale':
                recent_transactions += f"  - 🛒 বিক্রয়: {amount} টাকা (@{username}) {date_str}\n"
    else:
        recent_transactions = "  - কোনো সাম্প্রতিক লেনদেন নেই।\n"
    
    daily_deposits, daily_sales = get_report_summary(transaction_log, 1)
    weekly_deposits, weekly_sales = get_report_summary(transaction_log, 7)
    monthly_deposits, monthly_sales = get_report_summary(transaction_log, 30)

    dashboard_text = (
        f"📝 ড্যাশবোর্ড মেসেজ:\n"
        f"{dashboard_message}\n"
        "---------------------------\n"
        "📊 ড্যাশবোর্ড সামারি\n"
        f"👥 মোট ব্যবহারকারী: {total_users_count}\n"
        f"💰 মোট ব্যালেন্স: {sum(balances.values())} টাকা\n"
        f"🛒 মোট বিক্রয়: {total_sales} টাকা\n"
        f"💸 মোট ডিপোজিট: {total_deposits} টাকা\n"
        "---------------------------\n"
        "📈 দৈনিক/সাপ্তাহিক/মাসিক রিপোর্ট\n"
        f"গত ২৪ ঘণ্টা:\n"
        f"  - ডিপোজিট: {daily_deposits} টাকা\n"
        f"  - বিক্রয়: {daily_sales} টাকা\n"
        f"গত ৭ দিন:\n"
        f"  - ডিপোজিট: {weekly_deposits} টাকা\n"
        f"  - বিক্রয়: {weekly_sales} টাকা\n"
        f"গত ৩০ দিন:\n"
        f"  - ডিপোজিট: {monthly_deposits} টাকা\n"
        f"  - বিক্রয়: {monthly_sales} টাকা\n"
        "---------------------------\n"
        "📦 বর্তমান স্টক তথ্য:\n"
        f"{stock_info or '  - কোনো ক্যাটাগরি পাওয়া যায়নি।'}\n"
        "---------------------------\n"
        "📈 সর্বাধিক বিক্রীত ক্যাটাগরি:\n"
        f"{top_selling_info or '  - এখনো কোনো বিক্রয় হয়নি।'}\n"
        "---------------------------\n"
        "📜 সর্বশেষ লেনদেন:\n"
        f"{recent_transactions}\n"
    )
    
    if len(dashboard_text) > 4000:
        parts = [dashboard_text[i:i+4000] for i in range(0, len(dashboard_text), 4000)]
        for part in parts:
            update.message.reply_text(part)
    else:
        update.message.reply_text(dashboard_text)
    
    keyboard = [
        [KeyboardButton("🔄 Refresh Dashboard"), KeyboardButton("👥 User Profile")],
        [KeyboardButton("📂 Manage Categories"), KeyboardButton("💰 Edit Price")],
        [KeyboardButton("✏️ Edit Balance"), KeyboardButton("📢 Send Notice")],
        [KeyboardButton("💳 Payment Info"), KeyboardButton("💳 Payment Categories")],
        [KeyboardButton("🔙 Back to Main Menu")]
    ]
    
    update.message.reply_text("⚙️ অ্যাডমিন প্যানেল:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
    return ADMIN_PANEL

def handle_dashboard_refresh(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    if update.message.text == "🔄 Refresh Dashboard":
        return show_dashboard(update, context)
    return back_to_admin_panel_handler(update, context)

# ================ USER PROFILE ================
def view_user_profile(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    update.message.reply_text("✍️ যে ব্যবহারকারীর প্রোফাইল দেখতে চান তার ইউজারনাম বা ইউজার আইডি লিখুন:", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
    return SEARCH_USER_PROFILE

def search_and_show_user_profile(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    search_term = update.message.text.strip().lstrip('@')
    if search_term == "🔙 Admin Panel":
        return back_to_admin_panel_handler(update, context)
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
        update.message.reply_text("❌ এই ইউজারনেম বা ইউজার আইডি এর কোনো ব্যবহারকারী পাওয়া যায়নি।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
        return SEARCH_USER_PROFILE
    user_transactions = get_user_transactions(found_user_id, transaction_log)
    balance = balances.get(found_user_id, 0)
    deposits = user_deposits.get(found_user_id, 0)
    sales = user_sales.get(found_user_id, 0)
    daily_deposits, daily_sales = get_report_summary(user_transactions, 1)
    weekly_deposits, weekly_sales = get_report_summary(user_transactions, 7)
    monthly_deposits, monthly_sales = get_report_summary(user_transactions, 30)
    yearly_deposits, yearly_sales = get_report_summary(user_transactions, 365)
    user_data = user_info.get(found_user_id, {})
    full_name = user_data.get("first_name", "") + (f" {user_data['last_name']}" if user_data.get("last_name") else "")
    username = user_data.get('username', 'N/A')
    profile_text = (
        f"👤 ব্যবহারকারী প্রোফাইল:\n"
        f"নাম: {full_name}\n"
        f"ইউজারনেম: @{username}\n"
        f"আইডি: {found_user_id}\n"
        "---------------------------\n"
        f"💰 বর্তমান ব্যালেন্স: {balance} টাকা\n"
        f"💸 মোট ডিপোজিট: {deposits} টাকা\n"
        f"🛒 মোট খরচ: {sales} টাকা\n"
        "---------------------------\n"
        "📈 লেনদেনের রিপোর্ট\n"
        f"গত ২৪ ঘণ্টা:\n"
        f"  - ডিপোজিট: {daily_deposits} টাকা\n"
        f"  - খরচ: {daily_sales} টাকা\n"
        f"গত ৭ দিন:\n"
        f"  - ডিপোজিট: {weekly_deposits} টাকা\n"
        f"  - খরচ: {weekly_sales} টাকা\n"
        f"গত ৩০ দিন:\n"
        f"  - ডিপোজিট: {monthly_deposits} টাকা\n"
        f"  - খরচ: {monthly_sales} টাকা\n"
        f"গত ১ বছর:\n"
        f"  - ডিপোজিট: {yearly_deposits} টাকা\n"
        f"  - খরচ: {yearly_sales} টাকা\n"
    )
    update.message.reply_text(profile_text, reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
    return SEARCH_USER_PROFILE

# ================ BALANCE EDIT ================
def edit_user_balance_start(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    update.message.reply_text(
        "✍️ যে ব্যবহারকারীর ব্যালেন্স পরিবর্তন করতে চান তার ইউজারনাম বা ইউজার আইডি লিখুন:",
        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True)
    )
    return SEARCH_USER_FOR_BALANCE

def search_user_for_balance(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    search_term = update.message.text.strip().lstrip('@')
    if search_term == "🔙 Admin Panel":
        return back_to_admin_panel_handler(update, context)
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
        update.message.reply_text("❌ এই ইউজারনেম বা ইউজার আইডি এর কোনো ব্যবহারকারী পাওয়া যায়নি।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
        return SEARCH_USER_FOR_BALANCE
    context.user_data['edit_balance_user_id'] = found_user_id
    user_data = user_info.get(found_user_id, {})
    username = user_data.get('username', 'N/A')
    current_balance = balances.get(found_user_id, 0)
    keyboard = [
        [KeyboardButton("➕ Add Balance"), KeyboardButton("➖ Remove Balance")],
        [KeyboardButton("✍️ Set New Balance")],
        [KeyboardButton("🔙 Admin Panel")]
    ]
    update.message.reply_text(
        f"👤 ব্যবহারকারী: @{username} (আইডি: {found_user_id})\n"
        f"💰 বর্তমান ব্যালেন্স: {current_balance} টাকা\n\n"
        f"আপনি কী করতে চান?",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return BALANCE_EDIT_ACTION

def balance_edit_action_handler(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    action_text = update.message.text
    if action_text == "🔙 Admin Panel":
        return back_to_admin_panel_handler(update, context)
    if action_text not in ["➕ Add Balance", "➖ Remove Balance", "✍️ Set New Balance"]:
        update.message.reply_text("❌ অনুগ্রহ করে নিচের বাটন থেকে একটি অপশন বেছে নিন।")
        return BALANCE_EDIT_ACTION
    context.user_data['balance_edit_action'] = action_text
    if action_text == "➕ Add Balance":
        prompt = "✍️ কত টাকা যোগ করতে চান তা লিখুন:"
    elif action_text == "➖ Remove Balance":
        prompt = "✍️ কত টাকা সরাতে চান তা লিখুন:"
    else:
        prompt = "✍️ নতুন ব্যালেন্স কত হবে তা লিখুন:"
    update.message.reply_text(prompt, reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
    return RECEIVE_BALANCE_EDIT_AMOUNT

def receive_balance_edit_amount(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    amount_str = update.message.text.strip()
    if amount_str == "🔙 Admin Panel":
        context.user_data.pop('edit_balance_user_id', None)
        context.user_data.pop('balance_edit_action', None)
        return back_to_admin_panel_handler(update, context)
    if not amount_str.isdigit() or float(amount_str) < 0:
        update.message.reply_text("❌ অনুগ্রহ করে একটি ধনাত্মক সংখ্যা লিখুন।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
        return RECEIVE_BALANCE_EDIT_AMOUNT
    amount = float(amount_str)
    user_id = context.user_data.get('edit_balance_user_id')
    action = context.user_data.get('balance_edit_action')
    if not user_id or not action:
        update.message.reply_text("❌ একটি ত্রুটি ঘটেছে। অনুগ্রহ করে আবার শুরু করুন।")
        return back_to_admin_panel_handler(update, context)
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
            new_balance = 0
        balances[user_id] = new_balance
        user_message = f"✅ অ্যাডমিন আপনার ব্যালেন্স থেকে {amount} টাকা সরিয়ে নিয়েছে।\nআপনার নতুন ব্যালেন্স: {new_balance} টাকা।"
        admin_message = f"✅ ব্যবহারকারীর ব্যালেন্স থেকে {amount} টাকা সরানো হয়েছে।\nনতুন ব্যালেন্স: {new_balance} টাকা।"
    elif action == "✍️ Set New Balance":
        new_balance = amount
        balances[user_id] = new_balance
        user_message = f"✅ অ্যাডমিন আপনার নতুন ব্যালেন্স {new_balance} টাকা সেট করেছে।"
        admin_message = f"✅ ব্যবহারকারীর নতুন ব্যালেন্স {new_balance} টাকা সেট করা হয়েছে।"
    save_user_data()
    try:
        context.bot.send_message(chat_id=user_id, text=user_message)
    except Exception as e:
        logger.error(f"Failed to notify user {user_id} about balance change: {e}")
        admin_message += "\n⚠️ ব্যবহারকারীকে নোটিশ পাঠানো যায়নি।"
    update.message.reply_text(admin_message)
    context.user_data.pop('edit_balance_user_id', None)
    context.user_data.pop('balance_edit_action', None)
    return show_dashboard(update, context)

# ================ DEPOSIT ================
def deposit_handler(update, context):
    amount_str = update.message.text.strip()
    if amount_str == "🔙 Back to Main Menu":
        return start(update, context)
    if not amount_str.isdigit():
        update.message.reply_text("❌ অনুগ্রহ করে শুধু সংখ্যা লিখুন।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True))
        return DEPOSIT
    amount = int(amount_str)
    context.user_data["deposit_amount"] = amount
    keyboard = [[KeyboardButton("🔙 Back to Main Menu")]]
    update.message.reply_text(
        f"💳 এখন অনুগ্রহ করে {amount} টাকা পেমেন্ট করুন:\n{payment_info}\n\n"
        f"📸 পেমেন্টের পর স্ক্রিনশট পাঠান।",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return GET_DEPOSIT_AMOUNT

def receive_deposit_screenshot(update, context):
    if update.message.text == "🔙 Back to Main Menu":
        return start(update, context)
    if not update.message.photo:
        update.message.reply_text("❌ অনুগ্রহ করে শুধু একটি ছবি পাঠান।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True))
        return GET_DEPOSIT_AMOUNT
    user = update.effective_user
    deposit_amount = context.user_data.get("deposit_amount", 0)
    username = user.username if user.username else 'N/A'
    caption = (
        f"🔔 নতুন ডিপোজিট রিকোয়েস্ট! 🔔\n"
        f"ব্যবহারকারী: @{username}\n"
        f"পরিমাণ: {deposit_amount}\n"
        f"ইউজার আইডি: {user.id}"
    )
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Confirm Deposit", callback_data=f"deposit_confirm:{user.id}:{deposit_amount}"),
         InlineKeyboardButton("❌ Cancel Deposit", callback_data=f"deposit_cancel:{user.id}")]
    ])
    context.bot.send_photo(chat_id=ADMIN_ID, photo=update.message.photo[-1].file_id, caption=caption, reply_markup=keyboard)
    update.message.reply_text("✅ আপনার ডিপোজিট রিকোয়েস্ট অ্যাডমিনকে পাঠানো হয়েছে। অ্যাডমিন নিশ্চিত করার পর আপনার ব্যালেন্স যোগ হবে।")
    context.user_data.clear()
    return ConversationHandler.END

# ================ ADMIN PANEL ================
def back_to_admin_panel_handler(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    return show_dashboard(update, context)

def admin_panel_handler(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    text = update.message.text
    if text == "🔙 Back to Main Menu":
        return start(update, context)
    if text == "🔄 Refresh Dashboard":
        return show_dashboard(update, context)
    if text == "👥 User Profile":
        return view_user_profile(update, context)
    if text == "✏️ Edit Balance":
        return edit_user_balance_start(update, context)
    if text == "📢 Send Notice":
        update.message.reply_text("✍️ যে নোটিশটি পাঠাতে চান তা লিখুন:", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
        return SEND_NOTICE
    if text == "📂 Manage Categories":
        keyboard = []
        for cat in categories.keys():
            stock_count = get_total_stock(cat)
            keyboard.append([KeyboardButton(f"{cat} ({stock_count})")])
        keyboard.append([KeyboardButton("➕ Add Main Category")])
        keyboard.append([KeyboardButton("➖ Remove Main Category")])
        keyboard.append([KeyboardButton("🔙 Admin Panel")])
        update.message.reply_text("⚙️ প্রধান ক্যাটাগরি ব্যবস্থাপনা:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
        return MANAGE_CATEGORY
    if text == "💰 Edit Price":
        keyboard = [[KeyboardButton(cat)] for cat in categories.keys()]
        keyboard.append([KeyboardButton("🔙 Admin Panel")])
        update.message.reply_text("✍️ কোন ক্যাটাগরির মূল্য পরিবর্তন করবেন?", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
        return EDIT_PRICE_MAIN
    if text == "💳 Payment Info":
        update.message.reply_text("✍️ নতুন পেমেন্ট তথ্য পাঠান:", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
        return EDIT_PAYMENT
    if text == "💳 Payment Categories":
        return manage_payment_categories_handler(update, context)
    return ADMIN_PANEL

def manage_payment_categories_handler(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    keyboard_inline = []
    if not categories:
        update.message.reply_text("⚠️ কোনো ক্যাটাগরি নেই।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
        return ADMIN_PANEL
    for main_cat in categories:
        current_status = "Manual 💳" if main_cat in MANUAL_DELIVERY_CATEGORIES else "Balance 💰"
        button_text = "Switch to Balance" if main_cat in MANUAL_DELIVERY_CATEGORIES else "Switch to Manual"
        callback_data = f"toggle_payment:{main_cat}"
        keyboard_inline.append([
            InlineKeyboardButton(f"{main_cat} ({current_status})", callback_data="ignore"),
            InlineKeyboardButton(button_text, callback_data=callback_data)
        ])
    reply_markup_inline = InlineKeyboardMarkup(keyboard_inline)
    update.message.reply_text(
        "⚡️ পেমেন্ট ক্যাটাগরি নিয়ন্ত্রণ\n\n"
        "নিচের তালিকা থেকে প্রতিটি ক্যাটাগরির জন্য পেমেন্ট পদ্ধতি পরিবর্তন করতে পারেন।\n"
        "Balance Payment মানে ব্যবহারকারী তার ব্যালেন্স থেকে কিনতে পারবে।\n"
        "Manual Payment মানে ব্যবহারকারীকে সরাসরি পেমেন্ট করে স্ক্রিনশট পাঠাতে হবে।",
        reply_markup=reply_markup_inline
    )
    reply_markup_text = ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True)
    update.message.reply_text("🔙 অ্যাডমিন প্যানেলে ফিরে যেতে নিচের বাটনটি চাপুন।", reply_markup=reply_markup_text)
    return MANAGE_PAYMENT_CATEGORIES

def toggle_payment_method(update, context):
    query = update.callback_query
    query.answer()
    data = query.data
    if update.effective_user.id != ADMIN_ID:
        query.edit_message_caption("❌ অননুমোদিত।", reply_markup=None)
        return
    if data.startswith("toggle_payment:"):
        _, cat_name = data.split(":", 1)
        if cat_name in MANUAL_DELIVERY_CATEGORIES:
            MANUAL_DELIVERY_CATEGORIES.remove(cat_name)
        else:
            MANUAL_DELIVERY_CATEGORIES.append(cat_name)
        save_user_data()
        keyboard = []
        for main_cat in categories:
            current_status = "Manual 💳" if main_cat in MANUAL_DELIVERY_CATEGORIES else "Balance 💰"
            button_text = "Switch to Balance" if main_cat in MANUAL_DELIVERY_CATEGORIES else "Switch to Manual"
            callback_data = f"toggle_payment:{main_cat}"
            keyboard.append([
                InlineKeyboardButton(f"{main_cat} ({current_status})", callback_data="ignore"),
                InlineKeyboardButton(button_text, callback_data=callback_data)
            ])
        reply_markup = InlineKeyboardMarkup(keyboard)
        query.edit_message_text(
            "⚡️ পেমেন্ট ক্যাটাগরি নিয়ন্ত্রণ\n\n"
            "নিচের তালিকা থেকে প্রতিটি ক্যাটাগরির জন্য পেমেন্ট পদ্ধতি পরিবর্তন করতে পারেন।\n"
            "Balance Payment মানে ব্যবহারকারী তার ব্যালেন্স থেকে কিনতে পারবে।\n"
            "Manual Payment মানে ব্যবহারকারীকে সরাসরি পেমেন্ট করে স্ক্রিনশট পাঠাতে হবে।",
            reply_markup=reply_markup
        )
    return

def send_notice_text(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    notice_text = update.message.text
    notice_count = 0
    failed_users = []
    users_to_notify = [uid for uid in user_info.keys() if uid != ADMIN_ID]
    if not users_to_notify:
        update.message.reply_text("⚠️ কোনো ব্যবহারকারীকে নোটিশ পাঠানো হয়নি।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
        return back_to_admin_panel_handler(update, context)
    for user_id in users_to_notify:
        try:
            context.bot.send_message(chat_id=user_id, text=f"📢 নোটিশ:\n\n{notice_text}")
            notice_count += 1
        except Exception:
            failed_users.append(user_id)
    update.message.reply_text(f"✅ নোটিশ পাঠানো হয়েছে।\n\nসফল: {notice_count} জন\nব্যর্থ: {len(failed_users)} জন", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
    return back_to_admin_panel_handler(update, context)

# ================ MANAGE CATEGORIES ================
def back_to_manage_main_categories_handler(update, context):
    if "active_main_cat" in context.user_data:
        del context.user_data["active_main_cat"]
    if "active_sub_cat" in context.user_data:
        del context.user_data["active_sub_cat"]
    return manage_category_handler(update, context)

def manage_category_handler(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    text = update.message.text
    if text == "🔙 Admin Panel":
        return back_to_admin_panel_handler(update, context)
    if text == "➕ Add Main Category":
        update.message.reply_text("✍️ নতুন প্রধান ক্যাটাগরির নাম পাঠান:", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
        return ADD_MAIN_CAT
    if text == "➖ Remove Main Category":
        keyboard = []
        for cat in categories.keys():
            stock_count = get_total_stock(cat)
            keyboard.append([KeyboardButton(f"{cat} ({stock_count})")])
        keyboard.append([KeyboardButton("🔙 Admin Panel")])
        update.message.reply_text("➖ কোন প্রধান ক্যাটাগরি সরাতে চান?", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
        return REMOVE_MAIN_CAT
    original_cat = text.split(" (")[0]
    if original_cat in categories:
        context.user_data["active_main_cat"] = original_cat
        keyboard = []
        if categories.get(original_cat):
            for sub_cat in categories[original_cat]:
                stock_count = count_items(original_cat, sub_cat)
                keyboard.append([KeyboardButton(f"{sub_cat} ({stock_count})")])
        else:
            keyboard.append([KeyboardButton("⚠️ No Sub Categories")])
        keyboard.append([KeyboardButton("➕ Add Sub Category")])
        keyboard.append([KeyboardButton("➖ Remove Sub Category")])
        keyboard.append([KeyboardButton("➕ Add Items (TXT)")])
        keyboard.append([KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")])
        update.message.reply_text(f"⚙️ {original_cat} এর সাব-ক্যাটাগরি:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
        return MANAGE_SUB_CATEGORY
    return MANAGE_CATEGORY

def add_main_category(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    new_cat = update.message.text.strip()
    if new_cat == "🔙 Admin Panel":
        return back_to_admin_panel_handler(update, context)
    if new_cat in categories:
        update.message.reply_text("⚠️ এই প্রধান ক্যাটাগরি আগে থেকেই আছে।")
    else:
        categories[new_cat] = []
        save_user_data()
        update.message.reply_text(f"✅ প্রধান ক্যাটাগরি '{new_cat}' যোগ হয়েছে।")
    return manage_category_handler(update, context)

def remove_main_category(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    cat_to_remove = update.message.text.split(" (")[0].strip()
    if cat_to_remove == "🔙 Admin Panel":
        return back_to_admin_panel_handler(update, context)
    if cat_to_remove in categories:
        for sub_cat in categories[cat_to_remove]:
            txt_path = get_txt_path(cat_to_remove, sub_cat)
            if os.path.exists(txt_path):
                os.remove(txt_path)
            excel_path = get_excel_path(cat_to_remove, sub_cat)
            if os.path.exists(excel_path):
                os.remove(excel_path)
            if cat_to_remove in prices and sub_cat in prices[cat_to_remove]:
                del prices[cat_to_remove][sub_cat]
        if cat_to_remove in prices:
            del prices[cat_to_remove]
        del categories[cat_to_remove]
        if cat_to_remove in MANUAL_DELIVERY_CATEGORIES:
            MANUAL_DELIVERY_CATEGORIES.remove(cat_to_remove)
        save_user_data()
        update.message.reply_text(f"✅ প্রধান ক্যাটাগরি '{cat_to_remove}' সরানো হয়েছে।")
    else:
        update.message.reply_text("⚠️ এই ক্যাটাগরি পাওয়া যায়নি।")
    return manage_category_handler(update, context)

def manage_sub_category_handler(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    text = update.message.text
    main_cat = context.user_data.get("active_main_cat")
    if text == "🔙 Manage Categories":
        if "active_main_cat" in context.user_data:
            del context.user_data["active_main_cat"]
        if "active_sub_cat" in context.user_data:
            del context.user_data["active_sub_cat"]
        return manage_category_handler(update, context)
    if text == "🔙 Admin Panel":
        return back_to_admin_panel_handler(update, context)
    if text == "➕ Add Sub Category":
        update.message.reply_text("✍️ নতুন সাব-ক্যাটাগরির নাম পাঠান:", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
        return ADD_SUB_CAT
    if text == "➖ Remove Sub Category":
        if not categories.get(main_cat, []):
            update.message.reply_text("⚠️ কোনো সাব-ক্যাটাগরি নেই।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Manage Categories")]], resize_keyboard=True))
            return MANAGE_SUB_CATEGORY
        keyboard = [[KeyboardButton(sub_cat)] for sub_cat in categories.get(main_cat, [])]
        keyboard.append([KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")])
        update.message.reply_text("➖ কোন সাব-ক্যাটাগরি সরাতে চান?", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
        return REMOVE_SUB_CAT
    if text == "➕ Add Items (TXT)":
        if not categories.get(main_cat, []):
            update.message.reply_text("⚠️ কোনো সাব-ক্যাটাগরি নেই। আগে সাব-ক্যাটাগরি তৈরি করুন।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Manage Categories")]], resize_keyboard=True))
            return MANAGE_SUB_CATEGORY
        context.user_data["active_sub_cat"] = None
        keyboard = [[KeyboardButton(sub_cat)] for sub_cat in categories.get(main_cat, [])]
        keyboard.append([KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")])
        update.message.reply_text(
            f"📁 {main_cat} ক্যাটাগরির জন্য সাব-ক্যাটাগরি নির্বাচন করুন:\n\n"
            f"নিচের সাব-ক্যাটাগরি থেকে একটি নির্বাচন করুন যেখানে আইটেম যোগ করতে চান:",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
        return ADD_ITEMS_TXT
    original_sub_cat = text.split(" (")[0]
    if original_sub_cat in categories.get(main_cat, []):
        context.user_data["active_sub_cat"] = original_sub_cat
        count = count_items(main_cat, original_sub_cat)
        txt_path = get_txt_path(main_cat, original_sub_cat)
        preview = "কোনো আইটেম নেই"
        if os.path.exists(txt_path):
            with open(txt_path, 'r', encoding='utf-8') as f:
                items = [line.strip() for line in f.readlines() if line.strip()]
            if items:
                preview = "\n".join(items[:5])
                if len(items) > 5:
                    preview += f"\n... এবং আরও {len(items) - 5}টি আইটেম"
        keyboard = [
            [KeyboardButton("➕ Add Items (TXT)")],
            [KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")]
        ]
        update.message.reply_text(
            f"⚙️ সাব-ক্যাটাগরি: {original_sub_cat}\n"
            f"📦 মোট আইটেম: {count} টি\n"
            f"📄 আইটেম প্রিভিউ:\n{preview}\n\n"
            f"আইটেম যোগ করতে '➕ Add Items (TXT)' চাপুন।",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
        return ADD_ITEMS_TXT
    return MANAGE_SUB_CATEGORY

def add_sub_category(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    new_sub_cat = update.message.text.strip()
    main_cat = context.user_data.get("active_main_cat")
    if new_sub_cat == "🔙 Manage Categories":
        back_to_manage_main_categories_handler(update, context)
        return MANAGE_CATEGORY
    if new_sub_cat == "🔙 Admin Panel":
        return back_to_admin_panel_handler(update, context)
    if main_cat and new_sub_cat not in categories[main_cat]:
        categories[main_cat].append(new_sub_cat)
        ensure_txt_file(main_cat, new_sub_cat)
        save_user_data()
        update.message.reply_text(f"✅ সাব-ক্যাটাগরি '{new_sub_cat}' যোগ হয়েছে।")
    else:
        update.message.reply_text("⚠️ এই সাব-ক্যাটাগরি আগে থেকেই আছে।")
    keyboard = []
    if categories.get(main_cat):
        for sub_cat in categories[main_cat]:
            stock_count = count_items(main_cat, sub_cat)
            keyboard.append([KeyboardButton(f"{sub_cat} ({stock_count})")])
    keyboard.append([KeyboardButton("➕ Add Sub Category")])
    keyboard.append([KeyboardButton("➖ Remove Sub Category")])
    keyboard.append([KeyboardButton("➕ Add Items (TXT)")])
    keyboard.append([KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")])
    update.message.reply_text(f"⚙️ {main_cat} এর সাব-ক্যাটাগরি:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
    return MANAGE_SUB_CATEGORY

def remove_sub_category(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    sub_cat_to_remove = update.message.text.split(" (")[0].strip()
    main_cat = context.user_data.get("active_main_cat")
    if sub_cat_to_remove == "🔙 Manage Categories":
        back_to_manage_main_categories_handler(update, context)
        return MANAGE_CATEGORY
    if sub_cat_to_remove == "🔙 Admin Panel":
        return back_to_admin_panel_handler(update, context)
    if main_cat and sub_cat_to_remove in categories.get(main_cat, []):
        categories[main_cat].remove(sub_cat_to_remove)
        txt_path = get_txt_path(main_cat, sub_cat_to_remove)
        if os.path.exists(txt_path):
            os.remove(txt_path)
        excel_path = get_excel_path(main_cat, sub_cat_to_remove)
        if os.path.exists(excel_path):
            os.remove(excel_path)
        if main_cat in prices and sub_cat_to_remove in prices[main_cat]:
            del prices[main_cat][sub_cat_to_remove]
        save_user_data()
        update.message.reply_text(f"✅ সাব-ক্যাটাগরি '{sub_cat_to_remove}' সরানো হয়েছে।")
    else:
        update.message.reply_text("⚠️ এই সাব-ক্যাটাগরি পাওয়া যায়নি।")
    keyboard = []
    if categories.get(main_cat):
        for sub_cat in categories[main_cat]:
            stock_count = count_items(main_cat, sub_cat)
            keyboard.append([KeyboardButton(f"{sub_cat} ({stock_count})")])
    else:
        keyboard.append([KeyboardButton("⚠️ No Sub Categories")])
    keyboard.append([KeyboardButton("➕ Add Sub Category")])
    keyboard.append([KeyboardButton("➖ Remove Sub Category")])
    keyboard.append([KeyboardButton("➕ Add Items (TXT)")])
    keyboard.append([KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")])
    update.message.reply_text(f"⚙️ {main_cat} এর সাব-ক্যাটাগরি:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
    return MANAGE_SUB_CATEGORY

def add_items_txt_handler(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    text = update.message.text
    if text == "🔙 Manage Categories":
        back_to_manage_main_categories_handler(update, context)
        return MANAGE_CATEGORY
    if text == "🔙 Admin Panel":
        return back_to_admin_panel_handler(update, context)
    main_cat = context.user_data.get("active_main_cat")
    if not context.user_data.get("active_sub_cat"):
        if text in categories.get(main_cat, []):
            context.user_data["active_sub_cat"] = text
            update.message.reply_text(
                f"📁 {main_cat} → {text}\n\n"
                f"এখন টেক্সট ফাইল আপলোড করুন যা আইটেম ধারণ করে।\n"
                f"প্রতি লাইনে একটি আইটেম থাকতে হবে।\n\n"
                f"ফাইল আপলোড করুন অথবা '✅ Done' চাপুন।",
                reply_markup=ReplyKeyboardMarkup(
                    [[KeyboardButton("✅ Done"), KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")]],
                    resize_keyboard=True
                )
            )
            return ADD_ITEMS_TXT
        else:
            update.message.reply_text("❌ অনুগ্রহ করে একটি বৈধ সাব-ক্যাটাগরি নির্বাচন করুন।")
            return ADD_ITEMS_TXT
    sub_cat = context.user_data.get("active_sub_cat")
    if text == "✅ Done":
        count = count_items(main_cat, sub_cat)
        update.message.reply_text(f"✅ '{sub_cat}' তে মোট {count} টি আইটেম আছে।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
        context.user_data.pop("active_sub_cat", None)
        return back_to_admin_panel_handler(update, context)
    if update.message.document:
        file = update.message.document
        if file.file_name.endswith('.txt'):
            try:
                file_obj = file.get_file()
                file_content = file_obj.download_as_bytearray()
                txt_content = file_content.decode('utf-8')
                add_items_from_txt(main_cat, sub_cat, txt_content)
                count = count_items(main_cat, sub_cat)
                update.message.reply_text(
                    f"✅ '{sub_cat}' তে আইটেম যোগ হয়েছে।\n"
                    f"বর্তমান মোট আইটেম: {count} টি\n\n"
                    f"আরও ফাইল আপলোড করতে পারেন অথবা '✅ Done' চাপুন।",
                    reply_markup=ReplyKeyboardMarkup(
                        [[KeyboardButton("✅ Done"), KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")]],
                        resize_keyboard=True
                    )
                )
                return ADD_ITEMS_TXT
            except Exception as e:
                logger.error(f"Error processing TXT file: {e}")
                update.message.reply_text("❌ ফাইল প্রসেস করতে সমস্যা হয়েছে।")
                return ADD_ITEMS_TXT
        else:
            update.message.reply_text("❌ অনুগ্রহ করে শুধু .txt ফাইল আপলোড করুন।")
            return ADD_ITEMS_TXT
    if text and text not in ["✅ Done", "🔙 Manage Categories", "🔙 Admin Panel"]:
        add_items_from_txt(main_cat, sub_cat, text)
        count = count_items(main_cat, sub_cat)
        update.message.reply_text(
            f"✅ '{sub_cat}' তে আইটেম যোগ হয়েছে।\n"
            f"বর্তমান মোট আইটেম: {count} টি\n\n"
            f"আরও টেক্সট পাঠাতে পারেন অথবা '✅ Done' চাপুন।",
            reply_markup=ReplyKeyboardMarkup(
                [[KeyboardButton("✅ Done"), KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")]],
                resize_keyboard=True
            )
        )
        return ADD_ITEMS_TXT
    return ADD_ITEMS_TXT

# ================ EDIT PRICE ================
def edit_payment_info(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    global payment_info
    new_info = update.message.text.strip()
    if new_info == "🔙 Admin Panel":
        return back_to_admin_panel_handler(update, context)
    payment_info = new_info
    update.message.reply_text("✅ পেমেন্ট তথ্য সফলভাবে আপডেট হয়েছে।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
    return back_to_admin_panel_handler(update, context)

def edit_price_main_handler(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    text = update.message.text.strip()
    if text == "🔙 Admin Panel":
        return back_to_admin_panel_handler(update, context)
    original_cat = text.split(" (")[0]
    if original_cat in categories.keys():
        context.user_data['temp_main_cat_for_price'] = original_cat
        keyboard = [[KeyboardButton(sub_cat)] for sub_cat in categories[original_cat]]
        keyboard.append([KeyboardButton("🔙 Admin Panel")])
        update.message.reply_text(f"✍️ কোন সাব-ক্যাটাগরির মূল্য পরিবর্তন করবেন?", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
        return EDIT_PRICE_SUB
    update.message.reply_text("❌ এই ক্যাটাগরি পাওয়া যায়নি।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
    return EDIT_PRICE_MAIN

def edit_price_sub_handler(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    text = update.message.text.strip()
    main_cat = context.user_data.get('temp_main_cat_for_price')
    if text == "🔙 Admin Panel":
        if 'temp_main_cat_for_price' in context.user_data:
            del context.user_data['temp_main_cat_for_price']
        return back_to_admin_panel_handler(update, context)
    if main_cat and text in categories.get(main_cat, []):
        context.user_data['temp_sub_cat_for_price'] = text
        current_price = prices.get(main_cat, {}).get(text, "সেট করা হয়নি")
        update.message.reply_text(f"✍️ '{text}' এর বর্তমান মূল্য: {current_price}\nনতুন মূল্য লিখুন:", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
        return RECEIVE_NEW_PRICE
    update.message.reply_text("❌ এই সাব-ক্যাটাগরি পাওয়া যায়নি।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
    return EDIT_PRICE_SUB

def receive_price(update, context):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    price_text = update.message.text.strip()
    if price_text == "🔙 Admin Panel":
        if 'temp_main_cat_for_price' in context.user_data:
            del context.user_data['temp_main_cat_for_price']
        if 'temp_sub_cat_for_price' in context.user_data:
            del context.user_data['temp_sub_cat_for_price']
        return back_to_admin_panel_handler(update, context)
    try:
        new_price = float(price_text)
        main_cat = context.user_data.get('temp_main_cat_for_price')
        sub_cat = context.user_data.get('temp_sub_cat_for_price')
        if main_cat and sub_cat:
            if main_cat not in prices:
                prices[main_cat] = {}
            prices[main_cat][sub_cat] = new_price
            save_user_data()
            update.message.reply_text(f"✅ '{sub_cat}' এর মূল্য এখন {new_price} টাকা।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
            if 'temp_main_cat_for_price' in context.user_data:
                del context.user_data['temp_main_cat_for_price']
            if 'temp_sub_cat_for_price' in context.user_data:
                del context.user_data['temp_sub_cat_for_price']
            return back_to_admin_panel_handler(update, context)
    except ValueError:
        update.message.reply_text("❌ মূল্য শুধুমাত্র সংখ্যায় লিখুন।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
        return RECEIVE_NEW_PRICE
    return back_to_admin_panel_handler(update, context)

# ================ BUY FLOW ================
def back_to_categories_handler(update, context):
    if not check_subscription(update, context):
        return ConversationHandler.END
    keyboard = []
    for cat in categories.keys():
        keyboard.append([KeyboardButton(cat)])
    keyboard.append([KeyboardButton("🔙 Back to Main Menu")])
    update.message.reply_text("🛒 ক্যাটাগরি বেছে নিন:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
    return BUY_MENU

def back_to_subcategories_handler(update, context):
    if not check_subscription(update, context):
        return ConversationHandler.END
    main_cat = context.user_data.get('temp_main_cat_for_buy')
    if not main_cat:
        return back_to_categories_handler(update, context)
    keyboard = []
    for sub_cat in categories[main_cat]:
        stock_count = count_items(main_cat, sub_cat)
        keyboard.append([KeyboardButton(f"{sub_cat} ({stock_count})")])
    keyboard.append([KeyboardButton("🔙 Back to Categories")])
    keyboard.append([KeyboardButton("🔙 Back to Main Menu")])
    update.message.reply_text(f"🛒 {main_cat} এর সাব-ক্যাটাগরি বেছে নিন:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
    return BUY_SUB_MENU

def user_choose_category(update, context):
    if not check_subscription(update, context):
        return ConversationHandler.END
    text = update.message.text.strip()
    original_cat = text.split(" (")[0]
    if original_cat == "🔙 Back to Main Menu":
        return start(update, context)
    context.user_data.clear()
    if original_cat in categories.keys():
        context.user_data["temp_main_cat_for_buy"] = original_cat
        keyboard = []
        for sub_cat in categories[original_cat]:
            stock_count = count_items(original_cat, sub_cat)
            keyboard.append([KeyboardButton(f"{sub_cat} ({stock_count})")])
        keyboard.append([KeyboardButton("🔙 Back to Categories")])
        keyboard.append([KeyboardButton("🔙 Back to Main Menu")])
        update.message.reply_text(f"🛒 {original_cat} এর সাব-ক্যাটাগরি বেছে নিন:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
        return BUY_SUB_MENU
    update.message.reply_text("❌ এই ক্যাটাগরি পাওয়া যায়নি।")
    return BUY_MENU

def user_choose_subcategory(update, context):
    if not check_subscription(update, context):
        return ConversationHandler.END
    text = update.message.text.strip()
    original_sub_cat = text.split(" (")[0]
    main_cat = context.user_data.get('temp_main_cat_for_buy')
    if text == "🔙 Back to Categories":
        return back_to_categories_handler(update, context)
    if text == "🔙 Back to Main Menu":
        return start(update, context)
    if main_cat and original_sub_cat in categories.get(main_cat, []):
        context.user_data["order"] = {"main_cat": main_cat, "sub_cat": original_sub_cat}
        price = prices.get(main_cat, {}).get(original_sub_cat, "মূল্য এখনো সেট করা হয়নি।")
        keyboard = [
            [KeyboardButton("🔙 Back to Sub Categories")],
            [KeyboardButton("🔙 Back to Main Menu")]
        ]
        update.message.reply_text(f"✅ আপনি {original_sub_cat} সিলেক্ট করেছেন।\n💰 প্রতিটির দাম: {price} টাকা\n\n✍️ অনুগ্রহ করে আপনি কতগুলি চান তা সংখ্যায় লিখুন:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
        return GET_QUANTITY
    update.message.reply_text("❌ এই সাব-ক্যাটাগরি পাওয়া যায়নি।")
    return BUY_SUB_MENU

def receive_quantity(update, context):
    if not check_subscription(update, context):
        return ConversationHandler.END
    qty = update.message.text.strip()
    if qty == "🔙 Back to Sub Categories":
        return back_to_subcategories_handler(update, context)
    if qty == "🔙 Back to Main Menu":
        return start(update, context)
    if not qty.isdigit():
        update.message.reply_text("❌ অনুগ্রহ করে শুধু সংখ্যা লিখুন।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Sub Categories"), KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True))
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
                [KeyboardButton("🔙 Back to Sub Categories")],
                [KeyboardButton("🔙 Back to Main Menu")]
            ]
            update.message.reply_text(
                f"✅ অর্ডার তৈরি হয়েছে।\n"
                f"ক্যাটাগরি: {order['sub_cat']}\n"
                f"পরিমাণ: {order['qty']} টি\n"
                f"মোট দাম: {total_price} টাকা\n"
                f"আপনার বর্তমান ব্যালেন্স: {current_balance} টাকা\n\n"
                f"আপনার ব্যালেন্স থেকে পেমেন্ট করতে '✅ Confirm Purchase' চাপুন।",
                reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
            )
            return CONFIRM_ORDER
        else:
            update.message.reply_text(f"❌ আপনার যথেষ্ট ব্যালেন্স নেই। আপনার প্রয়োজন {total_price} টাকা কিন্তু ব্যালেন্স আছে {current_balance} টাকা।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Sub Categories"), KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True))
            return BUY_SUB_MENU
    keyboard = [
        [KeyboardButton("🔙 Back to Sub Categories")],
        [KeyboardButton("🔙 Back to Main Menu")]
    ]
    update.message.reply_text(
        f"✅ অর্ডার তৈরি হয়েছে।\n"
        f"ক্যাটাগরি: {order['sub_cat']}\n"
        f"পরিমাণ: {order['qty']} টি\n"
        f"মোট দাম: {total_price} টাকা\n"
        f"⚠️ অনুগ্রহ করে পেমেন্ট করুন:\n{payment_info}\n\n"
        f"📸 পেমেন্টের পর স্ক্রিনশট পাঠান।",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return WAIT_SCREENSHOT

def confirm_order(update, context):
    if not check_subscription(update, context):
        return ConversationHandler.END
    text = update.message.text
    if text == "🔙 Back to Sub Categories":
        return back_to_subcategories_handler(update, context)
    if text == "🔙 Back to Main Menu":
        return start(update, context)
    if text == "✅ Confirm Purchase":
        order = context.user_data.get("order", {})
        if not order:
            return ConversationHandler.END
        user_id = update.effective_user.id
        main_cat = order['main_cat']
        sub_cat = order['sub_cat']
        total_price = order['price']
        qty = order['qty']
        current_balance = balances.get(user_id, 0)
        if current_balance >= total_price:
            balances[user_id] = current_balance - total_price
            items = pop_items_from_txt(main_cat, sub_cat, qty)
            if not items:
                balances[user_id] = balances.get(user_id, 0) + total_price
                update.message.reply_text("❌ যথেষ্ট আইটেম স্টকে নেই। আপনার ব্যালেন্স ফেরত দেওয়া হয়েছে।", reply_markup=ReplyKeyboardRemove())
                return start(update, context)
            excel_buffer = create_xlsx_file(items, f"{sub_cat}_order.xlsx")
            context.bot.send_document(
                chat_id=user_id,
                document=InputFile(excel_buffer, filename=f"{sub_cat}_order_{int(time.time())}.xlsx"),
                caption=f"✅ আপনার অর্ডার সম্পূর্ণ হয়েছে!\n"
                        f"📦 {sub_cat} - {qty} টি আইটেম\n"
                        f"💰 মোট: {total_price} টাকা\n"
                        f"📄 এক্সেল ফাইলে আপনার অর্ডার সংযুক্ত আছে।"
            )
            global total_sales, sales_count_per_category, transaction_log, user_sales
            total_sales += total_price
            sales_count_per_category[sub_cat] = sales_count_per_category.get(sub_cat, 0) + qty
            transaction_log.append(('sale', user_id, total_price, time.time()))
            user_sales[user_id] = user_sales.get(user_id, 0) + total_price
            update.message.reply_text("✅ ব্যালেন্স ব্যবহার করে আপনার অর্ডার সফল হয়েছে!", reply_markup=ReplyKeyboardRemove())
            context.user_data.clear()
            save_user_data()
            return start(update, context)
        else:
            update.message.reply_text("❌ আপনার যথেষ্ট ব্যালেন্স নেই।", reply_markup=ReplyKeyboardRemove())
            return start(update, context)
    return CONFIRM_ORDER

def user_send_screenshot(update, context):
    if not check_subscription(update, context):
        return ConversationHandler.END
    if update.message.text == "🔙 Back to Sub Categories":
        return back_to_subcategories_handler(update, context)
    if update.message.text == "🔙 Back to Main Menu":
        return start(update, context)
    order = context.user_data.get("order", {})
    if not order:
        return ConversationHandler.END
    if not update.message.photo:
        update.message.reply_text("❌ অনুগ্রহ করে শুধু একটি ছবি পাঠান।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Sub Categories"), KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True))
        return WAIT_SCREENSHOT
    user = update.effective_user
    username = user.username if user.username else 'N/A'
    caption = (
        f"🔔 নতুন অর্ডার! 🔔\n"
        f"ব্যবহারকারী: @{username}\n"
        f"ক্যাটাগরি: {order['sub_cat']}\n"
        f"পরিমাণ: {order['qty']} টি\n"
        f"মূল্য: {order['price']} টাকা\n"
        f"ইউজার আইডি: {user.id}"
    )
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Confirm Order", callback_data=f"confirm_manual:{user.id}:{order['main_cat']}:{order['sub_cat']}:{order['qty']}:{order['price']}"),
         InlineKeyboardButton("❌ Cancel Order", callback_data=f"cancel_manual:{user.id}")]
    ])
    context.bot.send_photo(chat_id=ADMIN_ID, photo=update.message.photo[-1].file_id, caption=caption, reply_markup=keyboard)
    update.message.reply_text("✅ স্ক্রিনশট অ্যাডমিনকে পাঠানো হয়েছে।")
    start(update, context)
    context.user_data.clear()
    return ConversationHandler.END

# ================ ADMIN ORDER ACTIONS ================
def admin_order_action(update, context):
    global total_sales, sales_count_per_category, transaction_log, user_sales
    query = update.callback_query
    query.answer()
    data = query.data
    if update.effective_user.id != ADMIN_ID:
        query.edit_message_caption("❌ অননুমোদিত।", reply_markup=None)
        return
    if data.startswith("confirm_manual:"):
        parts = data.split(":")
        uid = int(parts[1])
        main_cat = parts[2]
        sub_cat = parts[3]
        qty = int(parts[4])
        total_price = float(parts[5])
        stock_count = count_items(main_cat, sub_cat)
        if stock_count < qty:
            keyboard = InlineKeyboardMarkup([
                [InlineKeyboardButton("✅ Force Confirm", callback_data=f"force_confirm:{uid}:{main_cat}:{sub_cat}:{qty}:{total_price}"),
                 InlineKeyboardButton("❌ Cancel", callback_data=f"cancel_manual:{uid}")]
            ])
            query.edit_message_caption(query.message.caption + "\n\n⚠️ স্টকে পর্যাপ্ত আইটেম নেই। আপনি কি নিশ্চিত করতে চান?\n\nযদি নিশ্চিত করেন, আপনাকে ম্যানুয়ালি আইটেমটি পাঠাতে হবে।", reply_markup=keyboard)
            return
        items = pop_items_from_txt(main_cat, sub_cat, qty)
        if not items:
            query.edit_message_caption(query.message.caption + "\n\n❌ স্টকে আইটেম নেই।", reply_markup=None)
            return
        excel_buffer = create_xlsx_file(items, f"{sub_cat}_order.xlsx")
        context.bot.send_document(
            chat_id=uid,
            document=InputFile(excel_buffer, filename=f"{sub_cat}_order_{int(time.time())}.xlsx"),
            caption=f"✅ আপনার অর্ডার নিশ্চিত করা হয়েছে!\n"
                    f"📦 {sub_cat} - {qty} টি আইটেম\n"
                    f"💰 মোট: {total_price} টাকা\n"
                    f"📄 এক্সেল ফাইলে আপনার অর্ডার সংযুক্ত আছে।"
        )
        total_sales += total_price
        sales_count_per_category[sub_cat] = sales_count_per_category.get(sub_cat, 0) + qty
        transaction_log.append(('sale', uid, total_price, time.time()))
        user_sales[uid] = user_sales.get(uid, 0) + total_price
        save_user_data()
        query.edit_message_caption(query.message.caption + f"\n\n✅ অ্যাডমিন দ্বারা নিশ্চিত। এক্সেল ফাইল ইউজারকে পাঠানো হয়েছে।", reply_markup=None)
    elif data.startswith("force_confirm:"):
        parts = data.split(":")
        uid = int(parts[1])
        main_cat = parts[2]
        sub_cat = parts[3]
        qty = int(parts[4])
        total_price = float(parts[5])
        context.bot.send_message(
            chat_id=uid,
            text="✅ আপনার অর্ডার নিশ্চিত করা হয়েছে। অ্যাডমিন শীঘ্রই আপনাকে বিস্তারিত তথ্য পাঠাবেন।\n\n"
                 "⚠️ দয়া করে অপেক্ষা করুন, অ্যাডমিন আপনার আইটেম প্রস্তুত করছে।"
        )
        context.bot.send_message(
            chat_id=ADMIN_ID,
            text=f"⚠️ ফোর্স নিশ্চিত অর্ডার\n"
                 f"ব্যবহারকারী: {uid}\n"
                 f"ক্যাটাগরি: {sub_cat}\n"
                 f"পরিমাণ: {qty} টি\n"
                 f"মোট: {total_price} টাকা\n\n"
                 f"স্টকে পর্যাপ্ত আইটেম নেই। অনুগ্রহ করে ম্যানুয়ালি আইটেম পাঠান।"
        )
        total_sales += total_price
        sales_count_per_category[sub_cat] = sales_count_per_category.get(sub_cat, 0) + qty
        transaction_log.append(('sale', uid, total_price, time.time()))
        user_sales[uid] = user_sales.get(uid, 0) + total_price
        save_user_data()
        query.edit_message_caption(query.message.caption + f"\n\n✅ ফোর্স নিশ্চিত করা হয়েছে। ইউজারকে ম্যানুয়ালি যোগাযোগ করা হবে।", reply_markup=None)
    elif data.startswith("cancel_manual:"):
        _, uid = data.split(":")
        context.bot.send_message(chat_id=uid, text="❌ দুঃখিত, আপনার অর্ডারটি বাতিল করা হয়েছে।")
        query.edit_message_caption(query.message.caption + "\n\n❌ অ্যাডমিন দ্বারা বাতিল", reply_markup=None)

# ================ ADMIN DEPOSIT ACTIONS ================
def admin_deposit_action(update, context):
    global total_deposits, transaction_log, user_deposits, balances
    query = update.callback_query
    query.answer()
    data = query.data
    if update.effective_user.id != ADMIN_ID:
        try:
            query.edit_message_caption("❌ অননুমোদিত।", reply_markup=None)
        except:
            pass
        return
    if data.startswith("deposit_confirm:"):
        parts = data.split(":")
        uid = int(parts[1])
        amount = float(parts[2])
        balances[uid] = balances.get(uid, 0) + amount
        total_deposits += amount
        transaction_log.append(('deposit', uid, amount, time.time()))
        user_deposits[uid] = user_deposits.get(uid, 0) + amount
        save_user_data()
        context.bot.send_message(
            chat_id=uid,
            text=f"✅ আপনার ডিপোজিট সফল হয়েছে। আপনার ব্যালেন্সে {amount} টাকা যোগ হয়েছে।\n"
                 f"নতুন ব্যালেন্স: {balances[uid]} টাকা।"
        )
        try:
            query.edit_message_caption(
                query.message.caption +
                f"\n\n✅ নিশ্চিত করা হয়েছে। {amount} টাকা ইউজার {uid} এর ব্যালেন্সে যোগ করা হয়েছে। "
                f"বর্তমান মোট ব্যালেন্স: {balances[uid]} টাকা।",
                reply_markup=None
            )
        except Exception as e:
            logger.error(f"Failed to edit message caption on deposit confirmation: {e}")
    else:
        _, uid = data.split(":")
        uid = int(uid)
        context.bot.send_message(chat_id=uid, text="❌ দুঃখিত, আপনার ডিপোজিট রিকোয়েস্ট বাতিল করা হয়েছে।")
        try:
            query.edit_message_caption(query.message.caption + "\n\n❌ অ্যাডমিন দ্বারা ডিপোজিট বাতিল", reply_markup=None)
        except Exception as e:
            logger.error(f"Failed to edit message caption on deposit cancellation: {e}")

# ================ MAIN ================
def main():
    logger.info("🤖 Starting Telegram Bot...")
    try:
        updater = Updater(token=BOT_TOKEN, use_context=True)
        dp = updater.dispatcher

        conv = ConversationHandler(
            entry_points=[
                CommandHandler("start", start),
                MessageHandler(Filters.text & ~Filters.command, menu_handler)
            ],
            states={
                MAIN_MENU: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Back to Main Menu"), start),
                    MessageHandler(Filters.text & ~Filters.command, menu_handler)
                ],
                BUY_MENU: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Back to Main Menu"), start),
                    MessageHandler(Filters.text & ~Filters.command, user_choose_category)
                ],
                BUY_SUB_MENU: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Back to Categories"), back_to_categories_handler),
                    MessageHandler(Filters.text & Filters.regex("🔙 Back to Main Menu"), start),
                    MessageHandler(Filters.text & ~Filters.command, user_choose_subcategory)
                ],
                GET_QUANTITY: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Back to Sub Categories"), back_to_subcategories_handler),
                    MessageHandler(Filters.text & Filters.regex("🔙 Back to Main Menu"), start),
                    MessageHandler(Filters.text & ~Filters.command, receive_quantity)
                ],
                CONFIRM_ORDER: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Back to Sub Categories"), back_to_subcategories_handler),
                    MessageHandler(Filters.text & Filters.regex("🔙 Back to Main Menu"), start),
                    MessageHandler(Filters.text & Filters.regex("✅ Confirm Purchase"), confirm_order),
                    MessageHandler(Filters.text & ~Filters.command, confirm_order)
                ],
                WAIT_SCREENSHOT: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Back to Sub Categories"), back_to_subcategories_handler),
                    MessageHandler(Filters.text & Filters.regex("🔙 Back to Main Menu"), start),
                    MessageHandler(Filters.photo | Filters.text, user_send_screenshot),
                ],
                ADMIN_PANEL: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Back to Main Menu"), start),
                    MessageHandler(Filters.text & ~Filters.command, admin_panel_handler)
                ],
                MANAGE_CATEGORY: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                    MessageHandler(Filters.text & ~Filters.command, manage_category_handler)
                ],
                MANAGE_SUB_CATEGORY: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                    MessageHandler(Filters.text & Filters.regex("🔙 Manage Categories"), back_to_manage_main_categories_handler),
                    MessageHandler(Filters.text & ~Filters.command, manage_sub_category_handler)
                ],
                ADD_MAIN_CAT: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                    MessageHandler(Filters.text & ~Filters.command, add_main_category)
                ],
                REMOVE_MAIN_CAT: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                    MessageHandler(Filters.text & ~Filters.command, remove_main_category)
                ],
                ADD_SUB_CAT: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                    MessageHandler(Filters.text & Filters.regex("🔙 Manage Categories"), back_to_manage_main_categories_handler),
                    MessageHandler(Filters.text & ~Filters.command, add_sub_category)
                ],
                REMOVE_SUB_CAT: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                    MessageHandler(Filters.text & Filters.regex("🔙 Manage Categories"), back_to_manage_main_categories_handler),
                    MessageHandler(Filters.text & ~Filters.command, remove_sub_category)
                ],
                ADD_ITEMS_TXT: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                    MessageHandler(Filters.text & Filters.regex("🔙 Manage Categories"), back_to_manage_main_categories_handler),
                    MessageHandler(Filters.text & Filters.regex("✅ Done"), add_items_txt_handler),
                    MessageHandler(Filters.document | Filters.text, add_items_txt_handler),
                ],
                EDIT_PAYMENT: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                    MessageHandler(Filters.text & ~Filters.command, edit_payment_info)
                ],
                EDIT_PRICE_MAIN: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                    MessageHandler(Filters.text & ~Filters.command, edit_price_main_handler)
                ],
                EDIT_PRICE_SUB: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                    MessageHandler(Filters.text & ~Filters.command, edit_price_sub_handler)
                ],
                RECEIVE_NEW_PRICE: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                    MessageHandler(Filters.text & ~Filters.command, receive_price)
                ],
                DEPOSIT: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Back to Main Menu"), start),
                    MessageHandler(Filters.text & ~Filters.command, deposit_handler),
                ],
                GET_DEPOSIT_AMOUNT: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Back to Main Menu"), start),
                    MessageHandler(Filters.photo, receive_deposit_screenshot)
                ],
                DASHBOARD: [
                    MessageHandler(Filters.text & Filters.regex("🔄 Refresh Dashboard"), handle_dashboard_refresh),
                    MessageHandler(Filters.text & Filters.regex("🔙 Admin Panel"), back_to_admin_panel_handler)
                ],
                SEND_NOTICE: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                    MessageHandler(Filters.text & ~Filters.command, send_notice_text)
                ],
                SEARCH_USER_PROFILE: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                    MessageHandler(Filters.text & ~Filters.command, search_and_show_user_profile)
                ],
                MANAGE_PAYMENT_CATEGORIES: [
                    CallbackQueryHandler(toggle_payment_method),
                    MessageHandler(Filters.text & Filters.regex("🔙 Admin Panel"), back_to_admin_panel_handler)
                ],
                SEARCH_USER_FOR_BALANCE: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                    MessageHandler(Filters.text & ~Filters.command, search_user_for_balance)
                ],
                BALANCE_EDIT_ACTION: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                    MessageHandler(Filters.text & ~Filters.command, balance_edit_action_handler)
                ],
                RECEIVE_BALANCE_EDIT_AMOUNT: [
                    MessageHandler(Filters.text & Filters.regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                    MessageHandler(Filters.text & ~Filters.command, receive_balance_edit_amount)
                ]
            },
            fallbacks=[CommandHandler("start", start)]
        )

        dp.add_handler(conv)
        dp.add_handler(CallbackQueryHandler(admin_order_action, pattern="^(confirm:|cancel:|confirm_manual:|cancel_manual:|force_confirm:)"))
        dp.add_handler(CallbackQueryHandler(admin_deposit_action, pattern="^(deposit_confirm:|deposit_cancel:)"))
        dp.add_handler(CallbackQueryHandler(toggle_payment_method, pattern="^toggle_payment:"))
        dp.add_handler(MessageHandler(Filters.text & ~Filters.command & Filters.regex("🔙 Back to Main Menu"), start))

        logger.info("🤖 Bot is running and polling...")
        updater.start_polling()
        updater.idle()

    except Exception as e:
        logger.error(f"Error starting bot: {e}")
        raise

if __name__ == "__main__":
    main()
