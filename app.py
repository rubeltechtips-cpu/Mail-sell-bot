import os
import threading
import sys
import asyncio
import warnings
warnings.filterwarnings("ignore")

from flask import Flask, jsonify

# Python 3.14 এর জন্য Event Loop Fix
if sys.version_info >= (3, 14):
    try:
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    except:
        pass

# আপনার বট ফাইল থেকে import
from bot import main as run_bot

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
    try:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        run_bot()
    except Exception as e:
        print(f"Bot error: {e}")

if __name__ == '__main__':
    print("Starting bot thread...")
    bot_thread = threading.Thread(target=run_bot_thread)
    bot_thread.daemon = True
    bot_thread.start()
    
    print("Starting Flask server...")
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
