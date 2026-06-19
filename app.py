import os
import sys
import json
import warnings
warnings.filterwarnings("ignore")

from flask import Flask, request, jsonify
from telegram import Update
from bot import application, BOT_TOKEN

app = Flask(__name__)

# Environment variables
WEBHOOK_URL = os.environ.get('RENDER_EXTERNAL_URL', 'https://mail-sell-bot.onrender.com')

@app.route('/')
def home():
    return jsonify({"status": "Bot is running!", "message": "Webhook mode active"})

@app.route('/health')
def health():
    return jsonify({"status": "healthy"})

@app.route('/ping')
def ping():
    return jsonify({"status": "pong"})

@app.route(f'/{BOT_TOKEN}', methods=['POST'])
def webhook():
    try:
        # Update ডেটা গ্রহণ করুন
        update_data = request.get_json(force=True)
        update = Update.de_json(update_data, application.bot)
        
        # Update প্রসেস করুন
        application.update_queue.put(update)
        return 'ok', 200
    except Exception as e:
        print(f"Webhook error: {e}")
        return 'error', 500

if __name__ == '__main__':
    # Webhook সেটআপ করুন
    try:
        application.bot.set_webhook(url=f'{WEBHOOK_URL}/{BOT_TOKEN}')
        print("Webhook set successfully!")
        print(f"Webhook URL: {WEBHOOK_URL}/{BOT_TOKEN}")
    except Exception as e:
        print(f"Failed to set webhook: {e}")
    
    # Flask সার্ভার চালান
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
