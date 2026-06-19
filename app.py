import os
import threading
import sys
import warnings
warnings.filterwarnings("ignore")

# Python 3.14 এর জন্য patch
if sys.version_info >= (3, 14):
    try:
        import telegram.ext._updater
        if not hasattr(telegram.ext._updater.Updater, '_Updater__polling_cleanup_cb'):
            telegram.ext._updater.Updater._Updater__polling_cleanup_cb = None
    except:
        pass

from your_bot import main as run_bot  # আপনার ফাইলের নাম অনুযায়ী

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
    try:
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
