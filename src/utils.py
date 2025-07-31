# -*- coding: utf-8 -*-
import os
import shutil
import tempfile
import logging
from datetime import datetime
from rich.console import Console
from rich.table import Table
from rich.logging import RichHandler

from src.constants import DIFF_REPORT_PATH, LOG_DIR, LOG_FILE_NAME_ALL

# --- 定数 ---
TABLE_DIFF_SHEET = "テーブル差分"
VBA_DIFF_SHEET = "VBA・フォーム差分"
TABLE_DIFF_HEADERS = ["差分タイプ", "テーブル名", "列1", "列2", "列3", "列4", "列5", "列6", "列7", "列8", "列9", "列10"]
VBA_DIFF_HEADERS = ["差分マーカー", "差分内容"]
MAX_DIFF_ROWS = 100

# --- Utility Functions ---
def is_file_locked(db_path):
    """
    Checks if the Access database file is locked.
    First, it checks for the existence of a lock file (.laccdb or .ldb).
    If not found, it tries to rename the file to itself, which fails if the file is open.
    """
    if not os.path.exists(db_path):
        return False
    
    lock_file = None
    if db_path.lower().endswith('.accdb'):
        lock_file = db_path[:-6] + '.laccdb'
    elif db_path.lower().endswith('.mdb'):
        lock_file = db_path[:-4] + '.ldb'

    if lock_file and os.path.exists(lock_file):
        return True

    try:
        # On Windows, renaming a file to itself fails if it's open.
        os.rename(db_path, db_path)
        return False
    except PermissionError:
        return True

def sanitize_for_excel(text, is_sheet_name=False):
    if not isinstance(text, str):
        return text
    invalid_chars = ['', '/', '*', '[', ']', ':', '?']
    for char in invalid_chars:
        text = text.replace(char, '_')
    if is_sheet_name:
        return text[:31]
    return text

def handle_com_error(e):
    console = Console()
    console.print("[bold red]Microsoft Accessとの連携中にエラーが発生しました。[/bold red]")
    if hasattr(e, 'excepinfo'):
        excepinfo = e.excepinfo
        message = excepinfo[2] if excepinfo and len(excepinfo) > 2 else "詳細不明"
        source = excepinfo[3] if excepinfo and len(excepinfo) > 3 else "不明"
        
        table = Table(title="COMエラー詳細", title_justify="left")
        table.add_column("項目", style="cyan", no_wrap=True)
        table.add_column("内容")
        table.add_row("メッセージ", message)
        table.add_row("ソース", source)
        console.print(table)
    else:
        console.print(f"[red]エラー詳細: {e}[/red]")

def setup_logging(log_level: int, console: Console):
    os.makedirs(LOG_DIR, exist_ok=True)
    log_file_name = LOG_FILE_NAME_ALL.format(datetime=datetime.now().strftime('%Y%m%d_%H%M%S'))
    log_file_path = os.path.join(LOG_DIR, log_file_name)

    # ルートロガーを取得
    root_logger = logging.getLogger()
    root_logger.setLevel(log_level)

    # 既存のハンドラをクリア (重複を避けるため)
    if root_logger.hasHandlers():
        root_logger.handlers.clear()

    # ファイルハンドラ
    file_handler = logging.FileHandler(log_file_path, encoding='utf-8')
    file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(name)s - %(message)s'))
    root_logger.addHandler(file_handler)

    

    return root_logger
