# -*- coding: utf-8 -*-
import os
import difflib
import pyodbc
import typer
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn
from datetime import datetime
import webbrowser
import logging

from src.core.access_handler import temporary_access_copy, access_application, export_objects
from src.core.db_operations import db_connection, get_table_names, get_table_data
from src.core.reporting import ReportGenerator
from src.utils import handle_com_error, sanitize_for_excel
from src.constants import DIFF_REPORT_PATH

console = Console()
logger = logging.getLogger(__name__)

def diff_text_files(file1_path, file2_path):
    try:
        with open(file1_path, 'r', encoding='utf-8', errors='ignore') as f1, \
             open(file2_path, 'r', encoding='utf-8', errors='ignore') as f2:
            diff = difflib.unified_diff(f1.readlines(), f2.readlines(), fromfile=os.path.basename(file1_path), tofile=os.path.basename(file2_path))
            return list(diff)
    except IOError as e:
        logger.error(f"テキストファイルの比較中にエラーが発生しました: {file1_path}, {file2_path} - {e}", exc_info=True)
        return [f"Error comparing files: {e}"]

def diff_exported_objects(dir1, dir2):
    diffs = {}
    files1 = set(os.listdir(dir1))
    files2 = set(os.listdir(dir2))
    all_files = sorted(list(files1 | files2))

    for filename in all_files:
        path1 = os.path.join(dir1, filename)
        path2 = os.path.join(dir2, filename)

        if filename in files1 and filename in files2:
            diff_content = diff_text_files(path1, path2)
            if diff_content:
                diffs[filename] = diff_content
        elif filename in files1:
            diffs[filename] = [f"--- {filename}", "+++ /dev/null", "@@ -1 +0,0 @@", "-Object only exists in the first file."]
        else:
            diffs[filename] = ["--- /dev/null", f"+++ {filename}", "@@ -0,0 +1 @@", "+Object only exists in the second file."]
    return diffs

def diff_tables(conn1, conn2):
    tables1 = set(get_table_names(conn1))
    tables2 = set(get_table_names(conn2))
    all_tables = sorted(list(tables1 | tables2))
    diffs = {}

    with Progress(SpinnerColumn(), TextColumn("[progress.description]{task.description}"), BarColumn(), TextColumn("[progress.percentage]{task.percentage:>3.0f}%")) as progress:
        task = progress.add_task("[cyan]テーブル比較中...[/cyan]", total=len(all_tables))
        for table in all_tables:
            progress.update(task, advance=1, description=f"[cyan]テーブル比較中...[/cyan] {table}")
            if table in tables1 and table in tables2:
                data1 = get_table_data(conn1, table)
                data2 = get_table_data(conn2, table)
                only_in_1 = data1 - data2
                only_in_2 = data2 - data1
                if only_in_1 or only_in_2:
                    diffs[table] = (only_in_1, only_in_2)
            elif table in tables1:
                diffs[table] = ({"Table only exists in file 1"}, set())
            else:
                diffs[table] = (set(), {"Table only exists in file 2"})
    return diffs

def diff_vba_objects(file1_path, file2_path, temp_dir1, temp_dir2):
    if os.path.splitext(file1_path)[1].lower() == ".accde" or os.path.splitext(file2_path)[1].lower() == ".accde":
        console.print("[yellow]⚠️ .accde ファイルのため、VBA/フォームの比較はスキップされます。[/yellow]")
        logger.warning(".accde ファイルのため、VBA/フォームの比較はスキップされます。")
        return {}

    export_dir1 = os.path.join(temp_dir1, "export")
    export_dir2 = os.path.join(temp_dir2, "export")

    os.makedirs(export_dir1, exist_ok=True)
    os.makedirs(export_dir2, exist_ok=True)

    try:
        with console.status("[bold green]VBA/フォーム/マクロをエクスポート中...[/]"):
            with access_application(file1_path) as app1:
                export_objects(app1, export_dir1)
            with access_application(file2_path) as app2:
                export_objects(app2, export_dir2)
    except Exception as e:
        console.print(f"[bold red]❌ オブジェクトのエクスポート中にエラーが発生しました: {e}[/bold red]")
        logger.error(f"オブジェクトのエクスポート中にエラーが発生しました: {e}", exc_info=True)
        return {"ExportError": f"Failed to export objects: {e}"}

    return diff_exported_objects(export_dir1, export_dir2)

def diff(file1_path: str = typer.Argument(..., help="比較元のAccessファイルのパス"),          file2_path: str = typer.Argument(..., help="比較先のAccessファイルのパス")):
    """
    2つのAccessデータベース（.accdb, .mdb）の差分を詳細に比較し、結果をExcelファイルに出力します。

    このコマンドは、以下の要素を比較します。
    - **テーブルデータ**: 各テーブルのレコードを比較し、追加・削除された行を特定します。
    - **VBAオブジェクト**: フォーム、レポート、モジュール、マクロ、クエリのソースコードや定義を比較し、変更点を明らかにします。

    比較結果は、見やすいように色分けされたExcelレポートとして `reports/` ディレクトリに保存され、完了後に自動で開かれます。
    .accdeファイルはVBAの比較がスキップされます。
    """
    file1_path = os.path.abspath(file1_path)
    file2_path = os.path.abspath(file2_path)
    logger.info(f"diff コマンドが実行されました。ファイル1: {file1_path}, ファイル2: {file2_path}")
    if not os.path.exists(file1_path):
        console.print(f"[bold red]エラー: ファイルが見つかりません: {file1_path}[/bold red]")
        logger.error(f"ファイルが見つかりません: {file1_path}")
        return
    if not os.path.exists(file2_path):
        console.print(f"[bold red]エラー: ファイルが見つかりません: {file2_path}[/bold red]")
        logger.error(f"ファイルが見つかりません: {file2_path}")
        return

    report_generator = ReportGenerator()

    console.rule("[bold blue]ファイル差分比較[/bold blue]")
    console.print(f"[cyan]ファイル1:[/cyan] {os.path.basename(file1_path)}")
    console.print(f"[cyan]ファイル2:[/cyan] {os.path.basename(file2_path)}")

    try:
        with temporary_access_copy(file1_path) as (f1_copy, temp1), \
             temporary_access_copy(file2_path) as (f2_copy, temp2):

            console.rule("[bold]テーブル比較[/bold]")
            logger.info("テーブル比較を開始します。")
            try:
                with db_connection(f1_copy) as conn1, db_connection(f2_copy) as conn2:
                    table_diffs = diff_tables(conn1, conn2)
                    logger.info("テーブル比較が完了しました。")
            except pyodbc.Error as e:
                console.print(f"[bold red]❌ DB接続に失敗したため、テーブル比較を中止します。: {e}[/bold red]")
                logger.error(f"DB接続に失敗したため、テーブル比較を中止します。: {e}", exc_info=True)
                table_diffs = {"DB Connection Error": (["Failed to connect to one or both databases."], [])}

            console.rule("[bold]VBA/フォーム/マクロ比較[/bold]")
            logger.info("VBA/フォーム/マクロ比較を開始します。")
            vba_diffs = diff_vba_objects(f1_copy, f2_copy, temp1, temp2)
            logger.info("VBA/フォーム/マクロ比較が完了しました。")

            console.rule("[bold]レポート作成[/bold]")
            report_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            report_generator.create_diff_report(table_diffs, vba_diffs, DIFF_REPORT_PATH, report_datetime, file1_path, file2_path)
            console.print(f"\n[bold green]✅ 比較結果を '{DIFF_REPORT_PATH}' に出力しました。[/bold green]")
            logger.info(f"比較結果を '{DIFF_REPORT_PATH}' に出力しました。")
            webbrowser.open(os.path.abspath(DIFF_REPORT_PATH))
    except Exception as e:
        handle_com_error(e)
        logger.error(f"diff コマンドの実行中に予期せぬエラーが発生しました: {e}", exc_info=True)
