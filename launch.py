"""Microjet Agent 打包入口：啟動 Flask 伺服器並自動開啟瀏覽器。"""
import os
import sys
import threading
import time
import webbrowser

# PyInstaller --onefile：資源解壓至 sys._MEIPASS；切換 CWD 以確保相對路徑與日誌輸出正常
if getattr(sys, 'frozen', False):
    os.chdir(sys._MEIPASS)

from app import app  # noqa: E402

HOST = '127.0.0.1'
PORT = 5001
URL = f'http://{HOST}:{PORT}'


def _open_browser():
    time.sleep(1.5)
    try:
        webbrowser.open(URL)
    except Exception:
        pass


def _banner():
    bar = '=' * 62
    print(bar)
    print('   Microjet AIP Agent  |  Palantir-style Orchestration')
    print(bar)
    print(f'   [READY]   Server listening on {URL}')
    print(f'   [BROWSER] 自動開啟；若未彈出請手動前往：{URL}')
    print('   [EXIT]    關閉此視窗即停止服務 (Ctrl+C)')
    print(bar)


if __name__ == '__main__':
    _banner()
    threading.Thread(target=_open_browser, daemon=True).start()
    # 打包環境關閉 reloader / debug，避免重複產生子行程
    app.run(host=HOST, port=PORT, debug=False, use_reloader=False, threaded=True)
