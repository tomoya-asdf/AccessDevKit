# -*- coding: utf-8 -*-
import os
import typer
from rich.console import Console
from rich.table import Table
import logging

from src.utils import handle_com_error
from src.core.access_handler import search_all_access_content

console = Console()
logger = logging.getLogger(__name__)

def search(file_path: str = typer.Argument(..., help="検索対象のAccessファイルのパス"), 
           pattern: str = typer.Argument(..., help="検索するキーワードまたは正規表現パターン")):
    """
    指定されたAccessファイル（.accdbまたは.mdb）内の全てのオブジェクトからキーワードを検索します。

    このコマンドは、以下の要素を横断的に検索し、開発者が特定の情報やコードの場所を迅速に見つけるのに役立ちます。
    - **VBAコード**: モジュール、フォーム、レポート、マクロ内のVBAコード。
    - **オブジェクト名**: テーブル、クエリ、フォーム、レポート、マクロ、モジュールの名前。
    - **テーブルデータ**: データベース内の各テーブルのデータ。

    検索結果は、オブジェクトの種類、名前、一致した行番号や列名、そして一致した内容とともにコンソールに表示されます。
    大文字・小文字は区別されません。
    """
    file_path = os.path.abspath(file_path)
    logger.info(f"search コマンドが実行されました。ファイルパス: {file_path}, 検索パターン: {pattern}")
    if not os.path.exists(file_path):
        console.print(f"[bold red]エラー: ファイルが見つかりません: {file_path}[/bold red]")
        logger.error(f"ファイルが見つかりません: {file_path}")
        raise typer.Exit(code=1)

    lock_file = file_path.replace('.accdb', '.laccdb').replace('.mdb', '.ldb')
    if os.path.exists(lock_file):
        console.print(f"[bold red]エラー: 対象ファイルは現在開かれているため、検索を実行できません。ファイルを閉じてから再度お試しください。: {file_path}[/bold red]")
        logger.error(f"対象ファイルは現在開かれています: {file_path}")
        raise typer.Exit(code=1)

    try:
        console.print(f"[cyan]検索を実行中... (キーワード: '{pattern}')[/cyan]")
        logger.info(f"検索を実行中... (キーワード: '{pattern}')")
        results = search_all_access_content(file_path, pattern)

        if not results:
            console.print("[yellow]キーワードに一致するオブジェクトは見つかりませんでした。[/yellow]")
            logger.info("キーワードに一致するオブジェクトは見つかりませんでした。")
            return

        table = Table(title=f"「{pattern}」の検索結果", title_justify="left", show_header=True, header_style="bold ")
        table.add_column("オブジェクト種類", style="cyan")
        table.add_column("オブジェクト名", style="green")
        table.add_column("行/レコード番号", style="dim")
        table.add_column("列名", style="dim")
        table.add_column("一致した内容")

        for result in results:
            table.add_row(
                result["type"],
                result["name"],
                str(result.get("line_num", "-")),
                result.get("column_name", "-"),
                result["line_content"]
            )
            logger.info(f"検索結果: 種類={result['type']}, 名前={result['name']}, 行/レコード番号={result.get('line_num', 'N/A')}, 列名={result.get('column_name', 'N/A')}, 内容={result['line_content']}")
        
        console.print(table)

    except Exception as e:
        handle_com_error(e)
        logger.error(f"search コマンドの実行中にエラーが発生しました: {e}", exc_info=True)