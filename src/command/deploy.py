# -*- coding: utf-8 -*-
import os
import shutil
import time
import threading
import hashlib
import logging

import typer
from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn

from src.utils import handle_com_error
from src.constants import LOG_DIR

# --- コンソールとロガー設定 ---
console = Console(record=True)
logger = logging.getLogger(__name__)

# --- グローバル変数 ---
overwrite_status = {}
lock = threading.Lock()
stop_event = threading.Event()

def get_file_hash(filepath, block_size=65536):
    if not os.path.exists(filepath) or stop_event.is_set():
        return None
    sha256 = hashlib.sha256()
    try:
        with open(filepath, 'rb') as f:
            for block in iter(lambda: f.read(block_size), b''):
                if stop_event.is_set(): return None
                sha256.update(block)
        return sha256.hexdigest()
    except (IOError, OSError) as e:
        logger.warning(f"ハッシュ値の計算に失敗: {filepath} - {e}")
        return None

def _try_overwrite(source_file, target_path, progress, task):
    if stop_event.is_set():
        return False

    source_hash = get_file_hash(source_file)
    if stop_event.is_set(): return False
    target_hash = get_file_hash(target_path)
    if stop_event.is_set(): return False

    if source_hash and source_hash == target_hash:
        logger.info(f"[green]✓[/green] スキップ (同一ファイル): {target_path}")
        progress.update(task, advance=1)
        return True

    try:
        temp_path = target_path + ".tmp"
        shutil.copy2(source_file, temp_path)
        if os.path.exists(target_path):
            os.remove(target_path)
        os.rename(temp_path, target_path)
        logger.info(f"[green]✓[/green] 上書き成功: {target_path}")
        progress.update(task, advance=1)
        return True
    except (IOError, OSError) as e:
        logger.warning(f"[yellow]⚠️[/yellow] 上書き失敗 (リトライ待ち): {target_path} - {e}")
        return False
    except Exception as e:
        logger.error(f"[bold red]❌[/bold red] 予期せぬエラー: {target_path} - {e}", exc_info=True)
        progress.update(task, advance=1)
        return True # 致命的なエラーとしてリトライ対象から外す

def _retry_overwrite_worker(progress, task):
    while not stop_event.is_set():
        with lock:
            pending_files = [(s, t) for (s, t), status in overwrite_status.items() if not status]

        if not pending_files:
            logger.info("全ての上書きが完了しました。リトライ処理を終了します。")
            break

        logger.info(f"[bold yellow]-- {len(pending_files)}件の上書きを60秒後にリトライします (Ctrl+Cで中止) --[/bold yellow]")
        if stop_event.wait(timeout=60):
            break

        with lock:
            for source, target in [(s, t) for (s, t), status in overwrite_status.items() if not status]:
                if stop_event.is_set(): break
                if _try_overwrite(source, target, progress, task):
                    overwrite_status[(source, target)] = True

def deploy(source_file: str = typer.Argument(..., help="展開元となるAccessファイルのパス"), 
           target_dir: str = typer.Argument(..., help="展開先のルートディレクトリパス")):
    """
    指定されたAccessファイル（.accdbまたは.mdb）を、対象ディレクトリ内の同名ファイルに展開（上書き）します。

    このコマンドは、主に開発環境から本番環境へのAccessファイルのデプロイを想定しています。
    - **ファイル検索**: `target_dir`以下を再帰的に検索し、`source_file`と同名のAccessファイルを見つけます。
    - **ハッシュ比較**: 展開元と展開先のファイルのハッシュ値を比較し、内容が異なる場合のみ上書きを実行します。
    - **自動リトライ**: ファイルが他のプロセスによってロックされているなど、上書きに失敗した場合は、
      ファイルが解放されるまで自動的にリトライを試みます。これにより、手動での介入なしにデプロイを完了できます。
    - **進捗表示**: 展開の進捗状況と結果をコンソールに表示します。

    **注意**: この操作は既存のファイルを上書きするため、実行前に必ずバックアップを取ることを推奨します。
    """
    source_file = os.path.abspath(source_file)
    target_dir = os.path.abspath(target_dir)
    global overwrite_status, stop_event
    overwrite_status = {}
    stop_event = threading.Event()

    console.rule("[bold blue]ファイル展開[/bold blue]")
    try:
        if not os.path.exists(source_file):
            logger.error(f"[bold red]❌ 展開元ファイルが見つかりません:[/bold red] {source_file}")
            return
        if not os.path.isdir(target_dir):
            logger.error(f"[bold red]❌ 展開先ディレクトリが見つかりません:[/bold red] {target_dir}")
            return

        source_filename = os.path.basename(source_file)
        target_files = []
        with console.status("[bold green]展開先ファイルを検索中...[/]"):
            for root, _, files in os.walk(target_dir):
                for file in files:
                    if file.lower() == source_filename.lower() and not file.startswith('~$'):
                        target_files.append(os.path.join(root, file))

        if not target_files:
            logger.warning(f"[yellow]展開先に同名のファイルが見つかりませんでした: {source_filename}[/yellow]")
            target_path = os.path.join(target_dir, source_filename)
            logger.info(f"代わりにディレクトリ直下にファイルをコピーします: {target_path}")
            try:
                shutil.copy2(source_file, target_path)
                console.print(f"[green]✅ ファイルをコピーしました: {target_path}[/green]")
            except Exception as e:
                logger.error(f"[bold red]❌ ファイルのコピーに失敗しました: {e}[/bold red]", exc_info=True)
            return

        table = Table(title="展開対象ファイル", title_justify="left", show_header=True, header_style="bold ")
        table.add_column("No.", style="dim", width=5)
        table.add_column("パス")
        for i, path in enumerate(target_files, 1):
            table.add_row(str(i), path)
        console.print(table)

        with lock:
            overwrite_status = {(source_file, T): False for T in target_files}

        with Progress(SpinnerColumn(), TextColumn("[progress.description]{task.description}"), BarColumn(), TextColumn("({task.completed}/{task.total})")) as progress:
            task = progress.add_task("[cyan]上書き処理中...[/cyan]", total=len(target_files))
            
            with lock:
                for source, target in list(overwrite_status.keys()):
                    if stop_event.is_set(): break
                    if _try_overwrite(source, target, progress, task):
                        overwrite_status[(source, target)] = True
            
            if not all(overwrite_status.values()) and not stop_event.is_set():
                retry_thread = threading.Thread(target=_retry_overwrite_worker, args=(progress, task), daemon=True)
                retry_thread.start()
                while retry_thread.is_alive():
                    retry_thread.join(timeout=0.5)

    except KeyboardInterrupt:
        logger.info("[bold yellow]🛑 中断シグナルを受信しました。処理を停止します...[/bold yellow]")
        stop_event.set()
    finally:
        successful = sum(1 for v in overwrite_status.values() if v)
        total = len(overwrite_status)
        
        summary_panel = Panel(
            f"[bold]結果: {successful} / {total} 件成功[/bold]", 
            title="[bold]展開サマリー[/bold]", 
            border_style="green" if successful == total else "red"
        )
        console.print(summary_panel)
        
        # ログファイルへの保存
        # setup_logging関数内でファイルハンドラが設定されているため、ここでは不要
        # console.print(f"[dim]詳細ログは {os.path.join(LOG_DIR, log_file_name)} を確認してください。[/dim]")