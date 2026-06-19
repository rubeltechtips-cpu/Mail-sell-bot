import logging
import os
import io
import time
import sys
import asyncio
import json
import warnings
warnings.filterwarnings("ignore")

from telegram import (
    Update, ReplyKeyboardMarkup, KeyboardButton, InputFile,
    ReplyKeyboardRemove, InlineKeyboardMarkup, InlineKeyboardButton
)
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    filters, ContextTypes, ConversationHandler, CallbackQueryHandler
)
from telegram.constants import ChatMemberStatus
from openpyxl import Workbook, load_workbook

# ================ CONFIG ================
BOT_TOKEN = os.environ.get('BOT_TOKEN', '8196949746:AAHPGwCkmoA-tYPe-vXwXro-ERp6a3a4s68')
ADMIN_ID = 8061006207
ADMIN_USERNAME = "Rubel_QSB"
CHANNEL_USERNAME = "quick_sell_bd"
DATA_DIR = "categories"

os.makedirs(DATA_DIR, exist_ok=True)

# ================ LOGGER ================
logging.basicConfig(format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ================ STATES ================
(
    MAIN_MENU,
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
    SEARCH_USER_FOR_BALANCE,
    BALANCE_EDIT_ACTION,
    RECEIVE_BALANCE_EDIT_AMOUNT,
    SEARCH_USER_FOR_EDIT,
    EDIT_USER_FIELD,
    RECEIVE_USER_EDIT_VALUE
) = range(30)

# ================ DATA ================
categories = {}
prices = {}
payment_info = "Bkash: 017XXXXXXXX\nNagad: 018XXXXXXXX\nBinance: yourmail@gmail.com"
balances = {}
user_sales = {}
user_deposits = {}
user_info = {}
total_deposits = 0
total_sales = 0
sales_count_per_category = {}
transaction_log = []
dashboard_message = "Welcome! This is your bot dashboard."
MANUAL_DELIVERY_CATEGORIES = []

# ================ USER DATA PERSISTENCE ================
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
def get_excel_path(main_cat: str, sub_cat: str) -> str:
    file_name = f"{main_cat}_{sub_cat}.xlsx".replace(" ", "_").replace("-", "_")
    return os.path.join(DATA_DIR, file_name)

def ensure_excel(main_cat: str, sub_cat: str):
    path = get_excel_path(main_cat, sub_cat)
    if not os.path.exists(path):
        wb = Workbook()
        ws = wb.active
        ws.title = "Items"
        wb.save(path)

def add_item_to_excel(main_cat: str, sub_cat: str, item: str):
    ensure_excel(main_cat, sub_cat)
    wb = load_workbook(get_excel_path(main_cat, sub_cat))
    ws = wb.active
    ws.append([item])
    wb.save(get_excel_path(main_cat, sub_cat))

def pop_items_from_excel(main_cat: str, sub_cat: str, qty: int):
    path = get_excel_path(main_cat, sub_cat)
    if not os.path.exists(path):
        return []
    wb = load_workbook(path)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    items = [r[0] for r in rows if r[0]]
    if len(items) < qty:
        return []
    
    result = items[:qty]
    new_items = items[qty:]
    wb = Workbook()
    ws = wb.active
    ws.title = "Items"
    for item in new_items:
        ws.append([item])
    wb.save(path)
    
    return result

def count_items(main_cat: str, sub_cat: str) -> int:
    path = get_excel_path(main_cat, sub_cat)
    if not os.path.exists(path):
        return 0
    wb = load_workbook(path)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    return len([r[0] for r in rows if r[0]])

def get_total_stock(main_cat: str) -> int:
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

# ================ MAIN MENU KEYBOARD ================
def get_main_menu_keyboard(user_id=None):
    keyboard = [
        [KeyboardButton("🛒 Buy"), KeyboardButton("💰 Balance")],
        [KeyboardButton("💸 Deposit"), KeyboardButton("📞 Help")],
    ]
    if user_id == ADMIN_ID:
        keyboard.append([KeyboardButton("⚙️ Admin Panel")])
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)

# ================ BACK BUTTONS ================
BACK_TO_MAIN = KeyboardButton("🔙 Back to Main Menu")
BACK_TO_CATEGORIES = KeyboardButton("🔙 Back to Categories")
BACK_TO_SUB_CATEGORIES = KeyboardButton("🔙 Back to Sub-Categories")
BACK_TO_ADMIN = KeyboardButton("🔙 Back to Admin Panel")
BACK_TO_MANAGE_MAIN = KeyboardButton("🔙 Back to Manage Main Categories")

# ================ CHANNEL CHECK (DISABLED FOR TESTING) ================
async def check_subscription(update: Update, context: ContextTypes.DEFAULT_TYPE):
    return True  # টেস্টের জন্য বাইপাস

# ================ HANDLERS ================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await check_subscription(update, context):
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
    await update.message.reply_text(
        f"👋 Welcome! Your current balance: {current_balance} Taka.",
        reply_markup=get_main_menu_keyboard(user_id)
    )
    return MAIN_MENU

async def back_to_main(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    current_balance = balances.get(user_id, 0)
    await update.message.reply_text(
        f"👋 Welcome back! Your current balance: {current_balance} Taka.",
        reply_markup=get_main_menu_keyboard(user_id)
    )
    return MAIN_MENU

async def menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await check_subscription(update, context):
        return ConversationHandler.END

    text = update.message.text
    
    if text == "🛒 Buy":
        if not categories:
            await update.message.reply_text("⚠️ No categories available.")
            return MAIN_MENU
        
        keyboard = []
        for cat in categories.keys():
            keyboard.append([KeyboardButton(cat)])
        keyboard.append([BACK_TO_MAIN])
        await update.message.reply_text(
            "🛒 Choose a category:",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
        return BUY_MENU

    if text == "💰 Balance":
        user_id = update.effective_user.id
        current_balance = balances.get(user_id, 0)
        await update.message.reply_text(
            f"💰 Your current balance: {current_balance} Taka.",
            reply_markup=get_main_menu_keyboard(user_id)
        )
        return MAIN_MENU

    if text == "💸 Deposit":
        await update.message.reply_text(
            "💰 Enter the amount you want to deposit:",
            reply_markup=ReplyKeyboardMarkup([[BACK_TO_MAIN]], resize_keyboard=True)
        )
        return DEPOSIT

    if text == "📞 Help":
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("📞 Chat with Admin", url=f"tg://user?id={ADMIN_ID}")]
        ])
        await update.message.reply_text(
            "📞 Click the button below to contact admin.",
            reply_markup=keyboard
        )
        return MAIN_MENU

    if text == "⚙️ Admin Panel":
        if update.effective_user.id == ADMIN_ID:
            return await show_dashboard(update, context)
        else:
            await update.message.reply_text("❌ Unauthorized.")
            return MAIN_MENU
            
    return MAIN_MENU

# ================ DASHBOARD ================
async def show_dashboard(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END

    global total_deposits, total_sales, balances, sales_count_per_category, user_info, transaction_log, dashboard_message

    total_balance = sum(balances.values())
    total_users_count = len(user_info)
    
    stock_info = ""
    for main_cat, sub_cats in categories.items():
        stock_info += f"  - <b>{main_cat}</b>\n"
        for sub_cat in sub_cats:
            count = count_items(main_cat, sub_cat)
            stock_info += f"    - {sub_cat}: {count} items\n"

    sorted_sales = sorted(sales_count_per_category.items(), key=lambda item: item[1], reverse=True)
    top_selling_info = ""
    for sub_cat, count in sorted_sales:
        top_selling_info += f"  - {sub_cat}: {count} sales\n"

    recent_transactions = ""
    last_5_transactions = transaction_log[-5:]
    if last_5_transactions:
        for trans in reversed(last_5_transactions):
            trans_type, user_id, amount, timestamp = trans
            date_str = time.strftime('%H:%M %b %d', time.localtime(timestamp))
            user_data = user_info.get(user_id, {})
            username = (user_data.get("username") or "N/A").replace('_', '\\_')
            
            if trans_type == 'deposit':
                recent_transactions += f"  - 💸 Deposit: {amount} (User: <code>@{username}</code>) at {date_str}\n"
            elif trans_type == 'sale':
                recent_transactions += f"  - 🛒 Sale: {amount} (User: <code>@{username}</code>) at {date_str}\n"
    else:
        recent_transactions = "  - No recent transactions.\n"
    
    daily_deposits, daily_sales = get_report_summary(transaction_log, 1)
    weekly_deposits, weekly_sales = get_report_summary(transaction_log, 7)
    monthly_deposits, monthly_sales = get_report_summary(transaction_log, 30)

    dashboard_text = (
        f"📝 <b>Dashboard Message:</b>\n"
        f"<i>{dashboard_message}</i>\n"
        "---------------------------\n"
        "📊 <b>Dashboard Summary</b>\n"
        f"👥 <b>Total Users:</b> {total_users_count}\n"
        f"💰 <b>Total Balance:</b> {sum(balances.values())} Taka\n"
        f"🛒 <b>Total Sales:</b> {total_sales} Taka\n"
        f"💸 <b>Total Deposits:</b> {total_deposits} Taka\n"
        "---------------------------\n"
        "📈 <b>Daily/Weekly/Monthly Report</b>\n"
        f"<b>Last 24 Hours:</b>\n"
        f"  - Deposits: {daily_deposits} Taka\n"
        f"  - Sales: {daily_sales} Taka\n"
        f"<b>Last 7 Days:</b>\n"
        f"  - Deposits: {weekly_deposits} Taka\n"
        f"  - Sales: {weekly_sales} Taka\n"
        f"<b>Last 30 Days:</b>\n"
        f"  - Deposits: {monthly_deposits} Taka\n"
        f"  - Sales: {monthly_sales} Taka\n"
        "---------------------------\n"
        "📦 <b>Current Stock Info:</b>\n"
        f"{stock_info or '  - No categories found.'}\n"
        "---------------------------\n"
        "📈 <b>Top Selling Categories:</b>\n"
        f"{top_selling_info or '  - No sales yet.'}\n"
        "---------------------------\n"
        "📜 <b>Recent Transactions:</b>\n"
        f"{recent_transactions}\n"
    )
    
    keyboard = [
        [KeyboardButton("🔄 Refresh Dashboard"), KeyboardButton("👥 User Profile")],
        [KeyboardButton("📂 Manage Categories"), KeyboardButton("💰 Edit Price")],
        [KeyboardButton("✏️ Edit User Balance"), KeyboardButton("✏️ Edit User Info")],
        [KeyboardButton("📢 Send Notice"), KeyboardButton("💳 Edit Payment Info")],
        [KeyboardButton("💳 Manage Payment Categories")],
        [BACK_TO_MAIN]
    ]
    
    await update.message.reply_text(
        dashboard_text,
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True),
        parse_mode='HTML'
    )
    return ADMIN_PANEL

# ================ BUY HANDLERS ================
async def user_choose_category(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await check_subscription(update, context):
        return ConversationHandler.END
        
    text = update.message.text.strip()
    original_cat = text.split(" (")[0]
    
    if original_cat == "🔙 Back to Main Menu":
        return await back_to_main(update, context)

    context.user_data.clear()

    if original_cat in categories.keys():
        context.user_data["temp_main_cat_for_buy"] = original_cat
        
        keyboard = []
        for sub_cat in categories[original_cat]:
            stock_count = count_items(original_cat, sub_cat)
            keyboard.append([KeyboardButton(f"{sub_cat} ({stock_count})")])
        
        keyboard.append([BACK_TO_CATEGORIES])
        keyboard.append([BACK_TO_MAIN])
        
        await update.message.reply_text(
            f"🛒 Choose sub-category for **{original_cat}**:",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
        return BUY_SUB_MENU

    await update.message.reply_text("❌ Category not found.")
    return BUY_MENU

async def user_choose_subcategory(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await check_subscription(update, context):
        return ConversationHandler.END
        
    text = update.message.text.strip()
    original_sub_cat = text.split(" (")[0]
    main_cat = context.user_data.get('temp_main_cat_for_buy')
    
    if text == "🔙 Back to Categories":
        return await back_to_categories(update, context)
    if text == "🔙 Back to Main Menu":
        return await back_to_main(update, context)
        
    if main_cat and original_sub_cat in categories.get(main_cat, []):
        context.user_data["order"] = {"main_cat": main_cat, "sub_cat": original_sub_cat}
        
        price = prices.get(main_cat, {}).get(original_sub_cat, "Price not set yet.")
        
        keyboard = [
            [BACK_TO_SUB_CATEGORIES],
            [BACK_TO_MAIN]
        ]
        
        await update.message.reply_text(
            f"✅ You selected **{original_sub_cat}**.\n💰 **Price per item:** {price} Taka\n\n"
            f"✍️ Enter quantity:",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
        return GET_QUANTITY

    await update.message.reply_text("❌ Sub-category not found.")
    return BUY_SUB_MENU

async def back_to_categories(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await check_subscription(update, context):
        return ConversationHandler.END

    keyboard = []
    for cat in categories.keys():
        keyboard.append([KeyboardButton(cat)])
    keyboard.append([BACK_TO_MAIN])
    await update.message.reply_text(
        "🛒 Choose a category:",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return BUY_MENU

async def back_to_subcategories(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await check_subscription(update, context):
        return ConversationHandler.END
    
    main_cat = context.user_data.get('temp_main_cat_for_buy')
    if not main_cat:
        return await back_to_categories(update, context)
        
    keyboard = []
    for sub_cat in categories[main_cat]:
        stock_count = count_items(main_cat, sub_cat)
        keyboard.append([KeyboardButton(f"{sub_cat} ({stock_count})")])
        
    keyboard.append([BACK_TO_CATEGORIES])
    keyboard.append([BACK_TO_MAIN])

    await update.message.reply_text(
        f"🛒 Choose sub-category for **{main_cat}**:",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return BUY_SUB_MENU

async def receive_quantity(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await check_subscription(update, context):
        return ConversationHandler.END

    qty = update.message.text.strip()

    if qty == "🔙 Back to Sub-Categories":
        return await back_to_subcategories(update, context)
    if qty == "🔙 Back to Main Menu":
        return await back_to_main(update, context)
    
    if not qty.isdigit():
        await update.message.reply_text(
            "❌ Please enter a valid number.",
            reply_markup=ReplyKeyboardMarkup([[BACK_TO_SUB_CATEGORIES], [BACK_TO_MAIN]], resize_keyboard=True)
        )
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
                [BACK_TO_SUB_CATEGORIES],
                [BACK_TO_MAIN]
            ]
            await update.message.reply_text(
                f"✅ Order created.\n"
                f"Category: {order['sub_cat']}\n"
                f"Quantity: {order['qty']}\n"
                f"Total Price: {total_price} Taka\n"
                f"Your balance: {current_balance} Taka\n\n"
                f"Click `✅ Confirm Purchase` to pay from balance.",
                reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
            )
        else:
            await update.message.reply_text(
                f"❌ Insufficient balance. You need {total_price} Taka but have {current_balance} Taka.",
                reply_markup=ReplyKeyboardMarkup([[BACK_TO_SUB_CATEGORIES], [BACK_TO_MAIN]], resize_keyboard=True)
            )
            return BUY_SUB_MENU
            
        return WAIT_SCREENSHOT
    
    keyboard = [
        [BACK_TO_SUB_CATEGORIES],
        [BACK_TO_MAIN]
    ]
    await update.message.reply_text(
        f"✅ Order created.\n"
        f"Category: {order['sub_cat']}\n"
        f"Quantity: {order['qty']}\n"
        f"Total Price: {total_price} Taka\n"
        f"⚠️ Please send payment to:\n{payment_info}\n\n"
        f"📸 Send a screenshot after payment.",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return WAIT_SCREENSHOT

async def user_send_screenshot(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await check_subscription(update, context):
        return ConversationHandler.END

    if update.message.text == "🔙 Back to Sub-Categories":
        return await back_to_subcategories(update, context)
    if update.message.text == "🔙 Back to Main Menu":
        return await back_to_main(update, context)

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
                await update.message.reply_text(
                    "❌ Not enough items in stock. Your balance has been refunded.",
                    reply_markup=ReplyKeyboardRemove()
                )
                return await back_to_main(update, context)
            
            global total_sales, sales_count_per_category, transaction_log, user_sales
            total_sales += total_price
            sales_count_per_category[sub_cat] = sales_count_per_category.get(sub_cat, 0) + qty
            transaction_log.append(('sale', user_id, total_price, time.time()))
            user_sales[user_id] = user_sales.get(user_id, 0) + total_price

            item_text = "\n".join(items)
            await context.bot.send_message(
                chat_id=user_id,
                text=f"✅ Your order:\n{item_text}"
            )

            await update.message.reply_text(
                "✅ Order completed successfully!",
                reply_markup=ReplyKeyboardRemove()
            )
            context.user_data.clear()
            save_user_data()
            return await back_to_main(update, context)
        else:
            await update.message.reply_text(
                "❌ Insufficient balance.",
                reply_markup=ReplyKeyboardRemove()
            )
            return await back_to_main(update, context)

    if not update.message.photo:
        await update.message.reply_text(
            "❌ Please send a photo.",
            reply_markup=ReplyKeyboardMarkup([[BACK_TO_SUB_CATEGORIES], [BACK_TO_MAIN]], resize_keyboard=True)
        )
        return WAIT_SCREENSHOT
        
    user = update.effective_user
    username = user.username if user.username else 'N/A'
    caption = (
        f"🔔 **New Order!** 🔔\n"
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
    
    await context.bot.send_photo(
        chat_id=ADMIN_ID,
        photo=update.message.photo[-1].file_id,
        caption=caption,
        reply_markup=keyboard
    )
    
    await update.message.reply_text("✅ Screenshot sent to admin.")
    await back_to_main(update, context)
    context.user_data.clear()
    return MAIN_MENU

# ================ DEPOSIT HANDLERS ================
async def deposit_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    amount_str = update.message.text.strip()

    if amount_str == "🔙 Back to Main Menu":
        return await back_to_main(update, context)

    if not amount_str.isdigit():
        await update.message.reply_text(
            "❌ Please enter a valid number.",
            reply_markup=ReplyKeyboardMarkup([[BACK_TO_MAIN]], resize_keyboard=True)
        )
        return DEPOSIT

    amount = int(amount_str)
    context.user_data["deposit_amount"] = amount
    
    await update.message.reply_text(
        f"💳 Please send {amount} Taka to:\n{payment_info}\n\n"
        f"📸 Send a screenshot after payment.",
        reply_markup=ReplyKeyboardMarkup([[BACK_TO_MAIN]], resize_keyboard=True)
    )
    return GET_DEPOSIT_AMOUNT

async def receive_deposit_screenshot(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.text == "🔙 Back to Main Menu":
        return await back_to_main(update, context)

    if not update.message.photo:
        await update.message.reply_text(
            "❌ Please send a photo.",
            reply_markup=ReplyKeyboardMarkup([[BACK_TO_MAIN]], resize_keyboard=True)
        )
        return GET_DEPOSIT_AMOUNT
    
    user = update.effective_user
    deposit_amount = context.user_data.get("deposit_amount", 0)
    
    username = user.username if user.username else 'N/A'
    caption = (
        f"🔔 **New Deposit Request!** 🔔\n"
        f"User: @{username}\n"
        f"Amount: {deposit_amount}\n"
        f"UserID: {user.id}"
    )

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Confirm Deposit", callback_data=f"deposit_confirm:{user.id}:{deposit_amount}"),
         InlineKeyboardButton("❌ Cancel Deposit", callback_data=f"deposit_cancel:{user.id}")]
    ])

    await context.bot.send_photo(
        chat_id=ADMIN_ID,
        photo=update.message.photo[-1].file_id,
        caption=caption,
        reply_markup=keyboard
    )
    
    await update.message.reply_text(
        "✅ Your deposit request has been sent to admin.",
        reply_markup=get_main_menu_keyboard(user.id)
    )
    
    context.user_data.clear()
    return MAIN_MENU

# ================ ADMIN PANEL HANDLERS ================
async def admin_panel_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END

    text = update.message.text
    
    if text == "🔙 Back to Main Menu":
        return await back_to_main(update, context)
    
    if text == "🔄 Refresh Dashboard":
        return await show_dashboard(update, context)
    
    if text == "👥 User Profile":
        return await view_user_profile(update, context)

    if text == "✏️ Edit User Balance":
        return await edit_user_balance_start(update, context)
    
    if text == "✏️ Edit User Info":
        return await edit_user_info_start(update, context)
        
    if text == "📢 Send Notice":
        await update.message.reply_text(
            "✍️ Enter the notice message:",
            reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
        )
        return SEND_NOTICE
        
    if text == "📂 Manage Categories":
        keyboard = []
        for cat in categories.keys():
            stock_count = get_total_stock(cat)
            keyboard.append([KeyboardButton(f"{cat} ({stock_count})")])
        
        keyboard.append([KeyboardButton("➕ Add Main Category")])
        keyboard.append([KeyboardButton("➖ Remove Main Category")])
        keyboard.append([BACK_TO_ADMIN])
        await update.message.reply_text(
            "⚙️ Manage Main Categories:",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
        return MANAGE_CATEGORY

    if text == "💰 Edit Price":
        keyboard = [[KeyboardButton(cat)] for cat in categories.keys()]
        keyboard.append([BACK_TO_ADMIN])
        await update.message.reply_text(
            "✍️ Which category price to edit?",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
        return EDIT_PRICE_MAIN

    if text == "💳 Edit Payment Info":
        await update.message.reply_text(
            "✍️ Enter new payment info:",
            reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
        )
        return EDIT_PAYMENT
        
    if text == "💳 Manage Payment Categories":
        return await manage_payment_categories_handler(update, context)
         
    return ADMIN_PANEL

# ================ USER PROFILE VIEW ================
async def view_user_profile(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END
    await update.message.reply_text(
        "✍️ Enter username or User ID to view profile:",
        reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
    )
    return SEARCH_USER_PROFILE

async def search_and_show_user_profile(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END
        
    search_term = update.message.text.strip().lstrip('@')
    
    if search_term == "🔙 Back to Admin Panel":
        return await show_dashboard(update, context)

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
            "❌ User not found. Please try again.",
            reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
        )
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
    username_safe = (user_data.get('username') or 'N/A').replace('<', '&lt;').replace('>', '&gt;')
    
    profile_text = (
        f"👤 <b>User Profile:</b>\n"
        f"Name: {full_name}\n"
        f"Username: <code>@{username_safe}</code>\n"
        f"ID: <code>{found_user_id}</code>\n"
        "---------------------------\n"
        f"💰 <b>Current Balance:</b> {balance} Taka\n"
        f"💸 <b>Total Deposits:</b> {deposits} Taka\n"
        f"🛒 <b>Total Spent:</b> {sales} Taka\n"
        "---------------------------\n"
        "📈 <b>Transaction Report</b>\n"
        f"<b>Last 24 Hours:</b>\n"
        f"  - Deposits: {daily_deposits} Taka\n"
        f"  - Spent: {daily_sales} Taka\n"
        f"<b>Last 7 Days:</b>\n"
        f"  - Deposits: {weekly_deposits} Taka\n"
        f"  - Spent: {weekly_sales} Taka\n"
        f"<b>Last 30 Days:</b>\n"
        f"  - Deposits: {monthly_deposits} Taka\n"
        f"  - Spent: {monthly_sales} Taka\n"
        f"<b>Last 1 Year:</b>\n"
        f"  - Deposits: {yearly_deposits} Taka\n"
        f"  - Spent: {yearly_sales} Taka\n"
    )
    
    await update.message.reply_text(
        profile_text,
        parse_mode='HTML',
        reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
    )
    return SEARCH_USER_PROFILE

# ================ EDIT USER INFO ================
async def edit_user_info_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END
    
    await update.message.reply_text(
        "✍️ Enter username or User ID to edit:",
        reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
    )
    return SEARCH_USER_FOR_EDIT

async def search_user_for_edit(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END
        
    search_term = update.message.text.strip().lstrip('@')
    
    if search_term == "🔙 Back to Admin Panel":
        return await show_dashboard(update, context)

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
            "❌ User not found. Please try again.",
            reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
        )
        return SEARCH_USER_FOR_EDIT

    context.user_data['edit_user_id'] = found_user_id
    
    user_data = user_info.get(found_user_id, {})
    username = user_data.get('username', 'N/A')
    first_name = user_data.get('first_name', 'N/A')
    last_name = user_data.get('last_name', 'N/A')
    
    keyboard = [
        [KeyboardButton("✏️ Edit First Name"), KeyboardButton("✏️ Edit Last Name")],
        [KeyboardButton("✏️ Edit Username")],
        [BACK_TO_ADMIN]
    ]
    
    await update.message.reply_text(
        f"👤 <b>Editing User:</b>\n"
        f"ID: <code>{found_user_id}</code>\n"
        f"Username: @{username}\n"
        f"First Name: {first_name}\n"
        f"Last Name: {last_name}\n\n"
        f"Choose what to edit:",
        parse_mode='HTML',
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return EDIT_USER_FIELD

async def edit_user_field_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END

    text = update.message.text
    
    if text == "🔙 Back to Admin Panel":
        context.user_data.pop('edit_user_id', None)
        return await show_dashboard(update, context)
    
    if text not in ["✏️ Edit First Name", "✏️ Edit Last Name", "✏️ Edit Username"]:
        await update.message.reply_text("❌ Please choose an option from the buttons below.")
        return EDIT_USER_FIELD
    
    field_map = {
        "✏️ Edit First Name": "first_name",
        "✏️ Edit Last Name": "last_name",
        "✏️ Edit Username": "username"
    }
    
    context.user_data['edit_field'] = field_map[text]
    field_name = context.user_data['edit_field'].replace('_', ' ').title()
    
    await update.message.reply_text(
        f"✍️ Enter new {field_name}:",
        reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
    )
    return RECEIVE_USER_EDIT_VALUE

async def receive_user_edit_value(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END
    
    new_value = update.message.text.strip()
    
    if new_value == "🔙 Back to Admin Panel":
        context.user_data.pop('edit_user_id', None)
        context.user_data.pop('edit_field', None)
        return await show_dashboard(update, context)
    
    user_id = context.user_data.get('edit_user_id')
    field = context.user_data.get('edit_field')
    
    if not user_id or not field:
        await update.message.reply_text("❌ Error. Please try again.")
        return await show_dashboard(update, context)
    
    if user_id not in user_info:
        await update.message.reply_text("❌ User not found.")
        return await show_dashboard(update, context)
    
    old_value = user_info[user_id].get(field, 'N/A')
    user_info[user_id][field] = new_value if new_value else None
    
    save_user_data()
    
    await update.message.reply_text(
        f"✅ {field.replace('_', ' ').title()} updated successfully!\n"
        f"Old: {old_value}\n"
        f"New: {new_value}",
        reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
    )
    
    context.user_data.pop('edit_user_id', None)
    context.user_data.pop('edit_field', None)
    return await show_dashboard(update, context)

# ================ EDIT USER BALANCE ================
async def edit_user_balance_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END
    
    await update.message.reply_text(
        "✍️ Enter username or User ID to edit balance:",
        reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
    )
    return SEARCH_USER_FOR_BALANCE

async def search_user_for_balance(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END
        
    search_term = update.message.text.strip().lstrip('@')
    
    if search_term == "🔙 Back to Admin Panel":
        return await show_dashboard(update, context)

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
            "❌ User not found. Please try again.",
            reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
        )
        return SEARCH_USER_FOR_BALANCE

    context.user_data['edit_balance_user_id'] = found_user_id
    
    user_data = user_info.get(found_user_id, {})
    username = user_data.get('username', 'N/A')
    current_balance = balances.get(found_user_id, 0)

    keyboard = [
        [KeyboardButton("➕ Add Balance"), KeyboardButton("➖ Remove Balance")],
        [KeyboardButton("✍️ Set New Balance")],
        [BACK_TO_ADMIN]
    ]
    
    await update.message.reply_text(
        f"👤 User: @{username} (ID: {found_user_id})\n"
        f"💰 Current Balance: {current_balance} Taka\n\n"
        f"What would you like to do?",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return BALANCE_EDIT_ACTION

async def balance_edit_action_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END

    action_text = update.message.text
    
    if action_text == "🔙 Back to Admin Panel":
        context.user_data.pop('edit_balance_user_id', None)
        return await show_dashboard(update, context)

    if action_text not in ["➕ Add Balance", "➖ Remove Balance", "✍️ Set New Balance"]:
        await update.message.reply_text("❌ Please choose an option from the buttons below.")
        return BALANCE_EDIT_ACTION

    context.user_data['balance_edit_action'] = action_text
    
    if action_text == "➕ Add Balance":
        prompt = "✍️ Enter amount to add:"
    elif action_text == "➖ Remove Balance":
        prompt = "✍️ Enter amount to remove:"
    else:
        prompt = "✍️ Enter new balance amount:"
        
    await update.message.reply_text(
        prompt,
        reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
    )
    return RECEIVE_BALANCE_EDIT_AMOUNT

async def receive_balance_edit_amount(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END

    amount_str = update.message.text.strip()
    
    if amount_str == "🔙 Back to Admin Panel":
        context.user_data.pop('edit_balance_user_id', None)
        context.user_data.pop('balance_edit_action', None)
        return await show_dashboard(update, context)

    if not amount_str.isdigit() or float(amount_str) < 0:
        await update.message.reply_text(
            "❌ Please enter a valid positive number.",
            reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
        )
        return RECEIVE_BALANCE_EDIT_AMOUNT
        
    amount = float(amount_str)
    user_id = context.user_data.get('edit_balance_user_id')
    action = context.user_data.get('balance_edit_action')

    if not user_id or not action:
        await update.message.reply_text("❌ An error occurred. Please try again.")
        return await show_dashboard(update, context)

    old_balance = balances.get(user_id, 0)
    new_balance = 0
    
    if action == "➕ Add Balance":
        new_balance = old_balance + amount
        balances[user_id] = new_balance
        user_message = f"✅ Admin has added {amount} Taka to your balance.\nYour new balance: {new_balance} Taka."
        admin_message = f"✅ Added {amount} Taka to user's balance.\nNew balance: {new_balance} Taka."

    elif action == "➖ Remove Balance":
        new_balance = max(0, old_balance - amount)
        balances[user_id] = new_balance
        user_message = f"✅ Admin has removed {amount} Taka from your balance.\nYour new balance: {new_balance} Taka."
        admin_message = f"✅ Removed {amount} Taka from user's balance.\nNew balance: {new_balance} Taka."

    elif action == "✍️ Set New Balance":
        new_balance = amount
        balances[user_id] = new_balance
        user_message = f"✅ Admin has set your new balance to {new_balance} Taka."
        admin_message = f"✅ User's balance set to {new_balance} Taka."

    save_user_data()

    try:
        await context.bot.send_message(chat_id=user_id, text=user_message)
    except Exception as e:
        logger.error(f"Failed to notify user {user_id}: {e}")
        admin_message += "\n⚠️ Could not notify user."

    await update.message.reply_text(admin_message)
    
    context.user_data.pop('edit_balance_user_id', None)
    context.user_data.pop('balance_edit_action', None)
    return await show_dashboard(update, context)

# ================ SEND NOTICE ================
async def send_notice_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END
    
    notice_text = update.message.text
    
    if notice_text == "🔙 Back to Admin Panel":
        return await show_dashboard(update, context)
    
    users_to_notify = [uid for uid in user_info.keys() if uid != ADMIN_ID]
    
    if not users_to_notify:
        await update.message.reply_text(
            "⚠️ No users to notify.",
            reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
        )
        return await show_dashboard(update, context)

    success_count = 0
    for user_id in users_to_notify:
        try:
            await context.bot.send_message(
                chat_id=user_id,
                text=f"📢 **Notice:**\n\n{notice_text}",
                parse_mode='Markdown'
            )
            success_count += 1
        except Exception:
            pass

    await update.message.reply_text(
        f"✅ Notice sent!\n\nSuccessful: {success_count} users\nFailed: {len(users_to_notify) - success_count} users",
        reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
    )
    return await show_dashboard(update, context)

# ================ MANAGE CATEGORIES ================
async def manage_category_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END

    text = update.message.text
    
    if text == "🔙 Back to Admin Panel":
        return await show_dashboard(update, context)

    if text == "➕ Add Main Category":
        await update.message.reply_text(
            "✍️ Enter new main category name:",
            reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
        )
        return ADD_MAIN_CAT

    if text == "➖ Remove Main Category":
        keyboard = []
        for cat in categories.keys():
            stock_count = get_total_stock(cat)
            keyboard.append([KeyboardButton(f"{cat} ({stock_count})")])
        
        keyboard.append([BACK_TO_ADMIN])
        await update.message.reply_text(
            "➖ Which main category to remove?",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
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
        keyboard.append([BACK_TO_MANAGE_MAIN])
        keyboard.append([BACK_TO_ADMIN])
        
        await update.message.reply_text(
            f"⚙️ Manage Sub Categories for **{original_cat}**:",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
        return MANAGE_SUB_CATEGORY
    
    return MANAGE_CATEGORY

async def add_main_category(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END

    new_cat = update.message.text.strip()
    
    if new_cat == "🔙 Back to Admin Panel":
        return await show_dashboard(update, context)
        
    if new_cat in categories:
        await update.message.reply_text("⚠️ Category already exists.")
    else:
        categories[new_cat] = []
        save_user_data()
        await update.message.reply_text(f"✅ Main category '{new_cat}' added.")
    
    return await manage_category_handler(update, context)

async def remove_main_category(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END

    cat_to_remove = update.message.text.split(" (")[0].strip()
    
    if cat_to_remove == "🔙 Back to Admin Panel":
        return await show_dashboard(update, context)
        
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
        if cat_to_remove in MANUAL_DELIVERY_CATEGORIES:
            MANUAL_DELIVERY_CATEGORIES.remove(cat_to_remove)
            
        save_user_data()
        await update.message.reply_text(f"✅ Main category '{cat_to_remove}' removed.")
    else:
        await update.message.reply_text("⚠️ Category not found.")
    
    return await manage_category_handler(update, context)

# ================ MANAGE SUB CATEGORIES ================
async def manage_sub_category_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END
        
    text = update.message.text
    main_cat = context.user_data.get("active_main_cat")
    
    if text == "🔙 Back to Manage Main Categories":
        context.user_data.pop("active_main_cat", None)
        return await manage_category_handler(update, context)
    
    if text == "🔙 Back to Admin Panel":
        context.user_data.pop("active_main_cat", None)
        return await show_dashboard(update, context)

    if text == "➕ Add Sub Category":
        await update.message.reply_text(
            "✍️ Enter new sub-category name:",
            reply_markup=ReplyKeyboardMarkup([[BACK_TO_MANAGE_MAIN], [BACK_TO_ADMIN]], resize_keyboard=True)
        )
        return ADD_SUB_CAT
        
    if text == "➖ Remove Sub Category":
        keyboard = [[KeyboardButton(sub_cat)] for sub_cat in categories.get(main_cat, [])]
        keyboard.append([BACK_TO_MANAGE_MAIN])
        keyboard.append([BACK_TO_ADMIN])
        await update.message.reply_text(
            "➖ Which sub-category to remove?",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
        return REMOVE_SUB_CAT
    
    original_sub_cat = text.split(" (")[0]
    
    if original_sub_cat in categories.get(main_cat, []):
        context.user_data["active_sub_cat"] = original_sub_cat
        count = count_items(main_cat, original_sub_cat)
        keyboard = [
            [KeyboardButton("➕ Add Items")],
            [BACK_TO_MANAGE_MAIN],
            [BACK_TO_ADMIN]
        ]
        await update.message.reply_text(
            f"⚙️ Sub-Category: {original_sub_cat}\n📦 Items in stock: {count}\n\n"
            f"Click `➕ Add Items` to add items.",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
        return ADD_ITEMS
        
    return MANAGE_SUB_CATEGORY

async def add_sub_category(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END

    new_sub_cat = update.message.text.strip()
    main_cat = context.user_data.get("active_main_cat")
    
    if new_sub_cat == "🔙 Back to Manage Main Categories":
        return await manage_sub_category_handler(update, context)

    if new_sub_cat == "🔙 Back to Admin Panel":
        context.user_data.pop("active_main_cat", None)
        return await show_dashboard(update, context)

    if main_cat and new_sub_cat not in categories[main_cat]:
        categories[main_cat].append(new_sub_cat)
        ensure_excel(main_cat, new_sub_cat)
        save_user_data()
        await update.message.reply_text(f"✅ Sub-category '{new_sub_cat}' added.")
    else:
        await update.message.reply_text("⚠️ Sub-category already exists or no main category selected.")
    
    return await manage_sub_category_handler(update, context)

async def remove_sub_category(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END

    sub_cat_to_remove = update.message.text.split(" (")[0].strip()
    main_cat = context.user_data.get("active_main_cat")
    
    if sub_cat_to_remove == "🔙 Back to Manage Main Categories":
        return await manage_sub_category_handler(update, context)

    if sub_cat_to_remove == "🔙 Back to Admin Panel":
        context.user_data.pop("active_main_cat", None)
        return await show_dashboard(update, context)
        
    if main_cat and sub_cat_to_remove in categories.get(main_cat, []):
        categories[main_cat].remove(sub_cat_to_remove)
        path = get_excel_path(main_cat, sub_cat_to_remove)
        if os.path.exists(path):
            os.remove(path)
        if main_cat in prices and sub_cat_to_remove in prices[main_cat]:
            del prices[main_cat][sub_cat_to_remove]
        save_user_data()
        await update.message.reply_text(f"✅ Sub-category '{sub_cat_to_remove}' removed.")
    else:
        await update.message.reply_text("⚠️ Sub-category not found.")
    
    return await manage_sub_category_handler(update, context)

# ================ ADD ITEMS ================
async def add_item_line(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END

    text = update.message.text
    
    if text == "🔙 Back to Manage Main Categories":
        return await manage_sub_category_handler(update, context)

    if text == "🔙 Back to Admin Panel":
        context.user_data.pop("active_main_cat", None)
        context.user_data.pop("active_sub_cat", None)
        return await show_dashboard(update, context)
    
    main_cat = context.user_data.get("active_main_cat")
    sub_cat = context.user_data.get("active_sub_cat")
    
    if not main_cat or not sub_cat:
        await update.message.reply_text(
            "⚠️ No category selected.",
            reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
        )
        return await manage_category_handler(update, context)
        
    if text == "➕ Add Items":
        await update.message.reply_text(
            f"✍️ Enter items for '{sub_cat}' (one per line).\n\n"
            f"Click `✅ Done` when finished.",
            reply_markup=ReplyKeyboardMarkup(
                [[KeyboardButton("✅ Done")],
                 [BACK_TO_MANAGE_MAIN],
                 [BACK_TO_ADMIN]],
                resize_keyboard=True
            )
        )
        return ADD_ITEMS
        
    if text == "✅ Done":
        count = count_items(main_cat, sub_cat)
        await update.message.reply_text(
            f"✅ '{sub_cat}' has {count} items total.",
            reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
        )
        return await manage_sub_category_handler(update, context)
    
    if text:
        items_to_add = text.split('\n')
        added_count = 0
        for item in items_to_add:
            item = item.strip()
            if item:
                add_item_to_excel(main_cat, sub_cat, item)
                added_count += 1
        await update.message.reply_text(
            f"✅ Added {added_count} item(s).",
            reply_markup=ReplyKeyboardMarkup(
                [[KeyboardButton("✅ Done")],
                 [BACK_TO_MANAGE_MAIN],
                 [BACK_TO_ADMIN]],
                resize_keyboard=True
            )
        )
    
    return ADD_ITEMS

# ================ EDIT PAYMENT INFO ================
async def edit_payment_info(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END
        
    global payment_info
    new_info = update.message.text.strip()
    
    if new_info == "🔙 Back to Admin Panel":
        return await show_dashboard(update, context)
        
    payment_info = new_info
    await update.message.reply_text(
        "✅ Payment info updated successfully.",
        reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
    )
    return await show_dashboard(update, context)

# ================ EDIT PRICE ================
async def edit_price_main_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END

    text = update.message.text.strip()
    
    if text == "🔙 Back to Admin Panel":
        return await show_dashboard(update, context)
    
    original_cat = text.split(" (")[0]
        
    if original_cat in categories.keys():
        context.user_data['temp_main_cat_for_price'] = original_cat
        keyboard = [[KeyboardButton(sub_cat)] for sub_cat in categories[original_cat]]
        keyboard.append([BACK_TO_ADMIN])
        await update.message.reply_text(
            f"✍️ Which sub-category price to edit?",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
        return EDIT_PRICE_SUB
        
    await update.message.reply_text(
        "❌ Category not found.",
        reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
    )
    return EDIT_PRICE_MAIN

async def edit_price_sub_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END
        
    text = update.message.text.strip()
    main_cat = context.user_data.get('temp_main_cat_for_price')
    
    if text == "🔙 Back to Admin Panel":
        context.user_data.pop('temp_main_cat_for_price', None)
        return await show_dashboard(update, context)
        
    if main_cat and text in categories.get(main_cat, []):
        context.user_data['temp_sub_cat_for_price'] = text
        current_price = prices.get(main_cat, {}).get(text, "Not set")
        await update.message.reply_text(
            f"✍️ Current price for '{text}': {current_price}\n"
            f"Enter new price:",
            reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
        )
        return RECEIVE_NEW_PRICE
        
    await update.message.reply_text(
        "❌ Sub-category not found.",
        reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
    )
    return EDIT_PRICE_SUB

async def receive_price(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END

    price_text = update.message.text.strip()
    
    if price_text == "🔙 Back to Admin Panel":
        context.user_data.pop('temp_main_cat_for_price', None)
        context.user_data.pop('temp_sub_cat_for_price', None)
        return await show_dashboard(update, context)

    try:
        new_price = float(price_text)
        main_cat = context.user_data.get('temp_main_cat_for_price')
        sub_cat = context.user_data.get('temp_sub_cat_for_price')
        if main_cat and sub_cat:
            if main_cat not in prices:
                prices[main_cat] = {}
            prices[main_cat][sub_cat] = new_price
            save_user_data()
            await update.message.reply_text(
                f"✅ Price for '{sub_cat}' is now {new_price} Taka.",
                reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
            )
            context.user_data.pop('temp_main_cat_for_price', None)
            context.user_data.pop('temp_sub_cat_for_price', None)
            return await show_dashboard(update, context)
    except ValueError:
        await update.message.reply_text(
            "❌ Please enter a valid number.",
            reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
        )
        return RECEIVE_NEW_PRICE
        
    return await show_dashboard(update, context)

# ================ MANAGE PAYMENT CATEGORIES ================
async def manage_payment_categories_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return ConversationHandler.END
        
    if not categories:
        await update.message.reply_text(
            "⚠️ No categories available.",
            reply_markup=ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
        )
        return ADMIN_PANEL

    keyboard_inline = []
    for main_cat in categories:
        current_status = "Manual Payment 💳" if main_cat in MANUAL_DELIVERY_CATEGORIES else "Balance Payment 💰"
        button_text = "Switch to Balance" if main_cat in MANUAL_DELIVERY_CATEGORIES else "Switch to Manual"
        callback_data = f"toggle_payment:{main_cat}"
        
        keyboard_inline.append([
            InlineKeyboardButton(f"{main_cat} ({current_status})", callback_data="ignore"),
            InlineKeyboardButton(button_text, callback_data=callback_data)
        ])
    
    reply_markup_inline = InlineKeyboardMarkup(keyboard_inline)
    
    await update.message.reply_text(
        "⚡️ **Payment Category Control**\n\n"
        "Toggle payment method for each category.\n"
        "**Balance Payment**: User pays from balance.\n"
        "**Manual Payment**: User sends payment screenshot.",
        reply_markup=reply_markup_inline,
        parse_mode='Markdown'
    )

    reply_markup_text = ReplyKeyboardMarkup([[BACK_TO_ADMIN]], resize_keyboard=True)
    await update.message.reply_text(
        "Click below to go back:",
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
            "⚡️ **Payment Category Control**\n\n"
            "Toggle payment method for each category.\n"
            "**Balance Payment**: User pays from balance.\n"
            "**Manual Payment**: User sends payment screenshot.",
            reply_markup=reply_markup,
            parse_mode='Markdown'
        )

# ================ ADMIN ORDER/DEPOSIT ACTIONS ================
async def admin_order_action(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global total_sales, sales_count_per_category, transaction_log, user_sales, prices
    query = update.callback_query
    await query.answer()
    data = query.data
    
    if update.effective_user.id != ADMIN_ID:
        await query.edit_message_caption("❌ Unauthorized.", reply_markup=None)
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
            await query.edit_message_caption(
                query.message.caption + "\n\n⚠️ Not enough stock. Force confirm?",
                reply_markup=keyboard
            )
            return
            
        items = pop_items_from_excel(main_cat, sub_cat, qty)
        
        total_sales += total_price
        sales_count_per_category[sub_cat] = sales_count_per_category.get(sub_cat, 0) + qty
        transaction_log.append(('sale', uid, total_price, time.time()))
        user_sales[uid] = user_sales.get(uid, 0) + total_price
        save_user_data()

        await context.bot.send_message(
            chat_id=uid,
            text="✅ Your order has been confirmed. Admin will contact you with details."
        )
        await query.edit_message_caption(
            query.message.caption + f"\n\n✅ Confirmed by Admin.",
            reply_markup=None
        )

    elif data.startswith("force_confirm:"):
        parts = data.split(":")
        uid = int(parts[1])
        main_cat = parts[2]
        sub_cat = parts[3]
        qty = int(parts[4])
        total_price = float(parts[5])

        total_sales += total_price
        sales_count_per_category[sub_cat] = sales_count_per_category.get(sub_cat, 0) + qty
        transaction_log.append(('sale', uid, total_price, time.time()))
        user_sales[uid] = user_sales.get(uid, 0) + total_price
        save_user_data()
        
        await context.bot.send_message(
            chat_id=uid,
            text="✅ Your order has been confirmed. Admin will contact you with details."
        )
        await query.edit_message_caption(
            query.message.caption + f"\n\n✅ Force Confirmed by Admin.",
            reply_markup=None
        )

    elif data.startswith("cancel_manual:"):
        _, uid = data.split(":")
        
        await context.bot.send_message(
            chat_id=uid,
            text="❌ Your order has been cancelled."
        )
        await query.edit_message_caption(
            query.message.caption + "\n\n❌ Cancelled by Admin",
            reply_markup=None
        )

async def admin_deposit_action(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global total_deposits, transaction_log, user_deposits, balances
    query = update.callback_query
    await query.answer()
    data = query.data

    if update.effective_user.id != ADMIN_ID:
        try:
            await query.edit_message_caption("❌ Unauthorized.", reply_markup=None)
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

        await context.bot.send_message(
            chat_id=uid,
            text=f"✅ Deposit successful! {amount} Taka added to your balance.\n"
                 f"New balance: {balances[uid]} Taka."
        )
        
        try:
            await query.edit_message_caption(
                query.message.caption +
                f"\n\n✅ Confirmed. {amount} Taka added to user <code>{uid}</code>'s balance.",
                reply_markup=None,
                parse_mode='HTML'
            )
        except Exception as e:
            logger.error(f"Failed to edit message: {e}")
            await context.bot.send_message(
                chat_id=ADMIN_ID,
                text=f"⚠️ Failed to update deposit message. User ID: {uid}, Amount: {amount}"
            )
            
    else:
        _, uid = data.split(":")
        uid = int(uid)
        
        await context.bot.send_message(
            chat_id=uid,
            text="❌ Your deposit request has been cancelled."
        )
        
        try:
            await query.edit_message_caption(
                query.message.caption + "\n\n❌ Cancelled by Admin",
                reply_markup=None
            )
        except Exception as e:
            logger.error(f"Failed to edit message: {e}")

# ================ MAIN FUNCTION ================
def main():
    global application
    
    # Application তৈরি করুন
    application = Application.builder().token(BOT_TOKEN).build()
    
    # Conversation Handler তৈরি করুন
    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            MessageHandler(filters.TEXT & ~filters.COMMAND, menu_handler)
        ],
        states={
            MAIN_MENU: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, menu_handler)
            ],
            BUY_MENU: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), back_to_main),
                MessageHandler(filters.TEXT & ~filters.COMMAND, user_choose_category)
            ],
            BUY_SUB_MENU: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Categories"), back_to_categories),
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), back_to_main),
                MessageHandler(filters.TEXT & ~filters.COMMAND, user_choose_subcategory)
            ],
            GET_QUANTITY: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Sub-Categories"), back_to_subcategories),
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), back_to_main),
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_quantity)
            ],
            WAIT_SCREENSHOT: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Sub-Categories"), back_to_subcategories),
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), back_to_main),
                MessageHandler(filters.TEXT & filters.Regex("✅ Confirm Purchase"), user_send_screenshot),
                MessageHandler(filters.PHOTO | filters.TEXT, user_send_screenshot),
            ],
            DEPOSIT: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), back_to_main),
                MessageHandler(filters.TEXT & ~filters.COMMAND, deposit_handler),
            ],
            GET_DEPOSIT_AMOUNT: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), back_to_main),
                MessageHandler(filters.PHOTO, receive_deposit_screenshot)
            ],
            ADMIN_PANEL: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), back_to_main),
                MessageHandler(filters.TEXT & ~filters.COMMAND, admin_panel_handler)
            ],
            SEND_NOTICE: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Admin Panel"), show_dashboard),
                MessageHandler(filters.TEXT & ~filters.COMMAND, send_notice_text)
            ],
            SEARCH_USER_PROFILE: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Admin Panel"), show_dashboard),
                MessageHandler(filters.TEXT & ~filters.COMMAND, search_and_show_user_profile)
            ],
            SEARCH_USER_FOR_BALANCE: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Admin Panel"), show_dashboard),
                MessageHandler(filters.TEXT & ~filters.COMMAND, search_user_for_balance)
            ],
            BALANCE_EDIT_ACTION: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Admin Panel"), show_dashboard),
                MessageHandler(filters.TEXT & ~filters.COMMAND, balance_edit_action_handler)
            ],
            RECEIVE_BALANCE_EDIT_AMOUNT: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Admin Panel"), show_dashboard),
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_balance_edit_amount)
            ],
            SEARCH_USER_FOR_EDIT: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Admin Panel"), show_dashboard),
                MessageHandler(filters.TEXT & ~filters.COMMAND, search_user_for_edit)
            ],
            EDIT_USER_FIELD: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Admin Panel"), show_dashboard),
                MessageHandler(filters.TEXT & ~filters.COMMAND, edit_user_field_handler)
            ],
            RECEIVE_USER_EDIT_VALUE: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Admin Panel"), show_dashboard),
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_user_edit_value)
            ],
            MANAGE_CATEGORY: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Admin Panel"), show_dashboard),
                MessageHandler(filters.TEXT & ~filters.COMMAND, manage_category_handler)
            ],
            MANAGE_SUB_CATEGORY: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Admin Panel"), show_dashboard),
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Manage Main Categories"), manage_category_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, manage_sub_category_handler)
            ],
            ADD_MAIN_CAT: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Admin Panel"), show_dashboard),
                MessageHandler(filters.TEXT & ~filters.COMMAND, add_main_category)
            ],
            REMOVE_MAIN_CAT: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Admin Panel"), show_dashboard),
                MessageHandler(filters.TEXT & ~filters.COMMAND, remove_main_category)
            ],
            ADD_SUB_CAT: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Admin Panel"), show_dashboard),
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Manage Main Categories"), manage_category_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, add_sub_category)
            ],
            REMOVE_SUB_CAT: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Admin Panel"), show_dashboard),
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Manage Main Categories"), manage_category_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, remove_sub_category)
            ],
            ADD_ITEMS: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Admin Panel"), show_dashboard),
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Manage Main Categories"), manage_category_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, add_item_line)
            ],
            EDIT_PAYMENT: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Admin Panel"), show_dashboard),
                MessageHandler(filters.TEXT & ~filters.COMMAND, edit_payment_info)
            ],
            EDIT_PRICE_MAIN: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Admin Panel"), show_dashboard),
                MessageHandler(filters.TEXT & ~filters.COMMAND, edit_price_main_handler)
            ],
            EDIT_PRICE_SUB: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Admin Panel"), show_dashboard),
                MessageHandler(filters.TEXT & ~filters.COMMAND, edit_price_sub_handler)
            ],
            RECEIVE_NEW_PRICE: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Admin Panel"), show_dashboard),
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_price)
            ],
            MANAGE_PAYMENT_CATEGORIES: [
                CallbackQueryHandler(toggle_payment_method),
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Admin Panel"), show_dashboard)
            ]
        },
        fallbacks=[CommandHandler("start", start)]
    )

    # Handler যোগ করুন
    application.add_handler(conv)
    application.add_handler(CallbackQueryHandler(admin_order_action, pattern="^(confirm_manual:|cancel_manual:|force_confirm:)"))
    application.add_handler(CallbackQueryHandler(admin_deposit_action, pattern="^(deposit_confirm:|deposit_cancel:)"))
    application.add_handler(CallbackQueryHandler(toggle_payment_method, pattern="^toggle_payment:"))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND & filters.Regex("🔙 Back to Main Menu"), back_to_main))
    
    logger.info("🤖 Bot running...")
    
    # Webhook এর জন্য application রিটার্ন করুন
    return application

# ================ APPLICATION START ================
if __name__ == "__main__":
    application = main()
    # Webhook mode এ চালানোর জন্য application এক্সপোর্ট করুন
else:
    # Import করার জন্য application তৈরি করুন
    application = main()
