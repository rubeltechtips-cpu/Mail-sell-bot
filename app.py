import os
import threading
from your_bot import main as run_bot  # আপনার ফাইলের নাম দিন (যেমন: your_bot.py)

from flask import Flask, jsonify

app = Flask(__name__)

@app.route('/')
def home():
    return jsonify({"status": "Bot is running!", "message": "Telegram bot is active"})

@app.route('/health')
def health():
    return jsonify({"status": "healthy"})

@app.route('/ping')
def ping():
    return jsonify({"status": "pong"})

def run_bot_thread():
    """বটকে আলাদা থ্রেডে চালান"""
    run_bot()

if __name__ == '__main__':
    # বট থ্রেড শুরু করুন
    bot_thread = threading.Thread(target=run_bot_thread)
    bot_thread.daemon = True
    bot_thread.start()
    
    # Flask সার্ভার চালান
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)