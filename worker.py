import os
import sys
import asyncio
import logging
import warnings
warnings.filterwarnings("ignore")

# Python 3.11 এ চলছে, তাই কোনো Event Loop Fix লাগবে না

from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler, CallbackQueryHandler
from bot import BOT_TOKEN, start, menu_handler, back_to_main, user_choose_category, user_choose_subcategory, back_to_categories, back_to_subcategories, receive_quantity, user_send_screenshot, deposit_handler, receive_deposit_screenshot, admin_panel_handler, show_dashboard, send_notice_text, search_and_show_user_profile, search_user_for_balance, balance_edit_action_handler, receive_balance_edit_amount, search_user_for_edit, edit_user_field_handler, receive_user_edit_value, manage_category_handler, manage_sub_category_handler, add_main_category, remove_main_category, add_sub_category, remove_sub_category, add_item_line, edit_payment_info, edit_price_main_handler, edit_price_sub_handler, receive_price, manage_payment_categories_handler, toggle_payment_method, admin_order_action, admin_deposit_action

# ================ LOGGER ================
logging.basicConfig(format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ================ MAIN ================
def main():
    # Application তৈরি করুন
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    
    # Conversation Handler তৈরি করুন
    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            MessageHandler(filters.TEXT & ~filters.COMMAND, menu_handler)
        ],
        states={
            # ... আপনার সব states এখানে যোগ করুন (আমি সংক্ষিপ্ত করছি)
        },
        fallbacks=[CommandHandler("start", start)]
    )

    # Handler যোগ করুন
    app.add_handler(conv)
    app.add_handler(CallbackQueryHandler(admin_order_action, pattern="^(confirm_manual:|cancel_manual:|force_confirm:)"))
    app.add_handler(CallbackQueryHandler(admin_deposit_action, pattern="^(deposit_confirm:|deposit_cancel:)"))
    app.add_handler(CallbackQueryHandler(toggle_payment_method, pattern="^toggle_payment:"))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND & filters.Regex("🔙 Back to Main Menu"), back_to_main))
    
    logger.info("🤖 Bot running in Worker mode...")
    
    # Polling শুরু করুন (Worker এর জন্য পারফেক্ট)
    app.run_polling()

if __name__ == "__main__":
    main()