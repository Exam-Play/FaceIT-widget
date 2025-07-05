<<<<<<< HEAD
from flask import Flask, request, jsonify
from flask_cors import CORS
import asyncio
import json
import time
import os
from threading import Thread
from faceit import get_full_stats

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
stats_file = os.path.join(BASE_DIR, "stats.json")

app = Flask(__name__)
CORS(app)

current_username = "_ExamPlay_"

def get_username():
    global current_username
    return current_username

def write_stats(username, stats):
    with open(stats_file, "w") as f:
        json.dump({"username": username, **stats}, f, indent=2)

@app.route("/stats.json", methods=["GET", "POST"])
def stats():
    global current_username
    if request.method == "POST":
        username = request.json.get("username", "_ExamPlay_")
        current_username = username

        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        stats_data = loop.run_until_complete(get_full_stats(username))
        loop.close()

        if stats_data:
            write_stats(username, stats_data)
            return jsonify(stats_data)
        else:
            return jsonify({"ok": False, "error": "Статистика не найдена"}), 400

    else:
        try:
            with open(stats_file, "r") as f:
                return f.read()
        except:
            return jsonify({"username": "_ExamPlay_", "elo": "-", "kd": "-", "winrate": "-"})


def update_stats_loop():
    while True:
        username = get_username()

        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        stats_data = loop.run_until_complete(get_full_stats(username))
        loop.close()

        if stats_data:
            print(f"[✅] {username} | Elo: {stats_data['elo']} | K/D: {stats_data.get('kd')} | ADR: {stats_data.get('adr')} | WinRate: {stats_data.get('win_rate')}%")
            write_stats(username, stats_data)
        else:
            print(f"[⚠️] Не удалось обновить статистику для {username}")

        time.sleep(10)

if __name__ == "__main__":
    Thread(target=update_stats_loop, daemon=True).start()
    app.run()
=======
from flask import Flask, request, jsonify
from flask_cors import CORS
import asyncio
import json
import time
from threading import Thread
from faceit import get_full_stats

app = Flask(__name__)
CORS(app)

stats_file = "stats.json"
current_username = "_ExamPlay_"

def get_username():
    global current_username
    return current_username

def write_stats(username, stats):
    with open(stats_file, "w") as f:
        json.dump({"username": username, **stats}, f, indent=2)

@app.route("/stats.json", methods=["GET", "POST"])
def stats():
    global current_username
    if request.method == "POST":
        username = request.json.get("username", "_ExamPlay_")
        current_username = username

        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        stats_data = loop.run_until_complete(get_full_stats(username))
        loop.close()

        if stats_data:
            write_stats(username, stats_data)
            return jsonify(stats_data)
        else:
            return jsonify({"ok": False, "error": "Статистика не найдена"}), 400

    else:
        try:
            with open(stats_file, "r") as f:
                return f.read()
        except:
            return jsonify({"username": "_ExamPlay_", "elo": "-", "kd": "-", "winrate": "-"})


def update_stats_loop():
    while True:
        username = get_username()

        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        stats_data = loop.run_until_complete(get_full_stats(username))
        loop.close()

        if stats_data:
            print(f"[✅] {username} | Elo: {stats_data['elo']} | K/D: {stats_data.get('kd')} | ADR: {stats_data.get('adr')} | WinRate: {stats_data.get('win_rate')}%")
            write_stats(username, stats_data)
        else:
            print(f"[⚠️] Не удалось обновить статистику для {username}")

        time.sleep(10)


if __name__ == "__main__":
    Thread(target=update_stats_loop, daemon=True).start()
    app.run()
>>>>>>> 42252af700a393122e0190ef25bf9e64dc5c778f
