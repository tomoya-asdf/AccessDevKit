# -*- coding: utf-8 -*-
import os
import typer
from rich.console import Console
from rich.tree import Tree
import logging

from src.utils import handle_com_error
from src.core.access_handler import access_application, import_objects
from src.constants import BASE_APP_DIR

console = Console()
logger = logging.getLogger(__name__)

def load(file_path: str = typer.Argument(..., help="インポート対象のAccessファイルのパス"), 
          input_dir: str = typer.Option(os.path.join(BASE_APP_DIR, "output", "export"), "--input", "-i", help="インポートするオブジェクトが格納されているディレクトリ。デフォルトは `./output/export` です。")):
    """
    エクスポートされたオブジェクトを、指定のAccessファイル（*.accdbまたは.mdb）にインポートします。

    このコマンドは、`export`コマンドでエクスポートされたファイル形式に対応しています。
    - フォーム (.frm)
    - レポート (.rpt)
    - マクロ (.mcr)
    - モジュール (.bas)
    - クエリ (.qry)

    これにより、バージョン管理システムで管理されたオブジェクトをAccessファイルに反映させたり、
    異なるAccessファイル間でオブジェクトを共有したりすることが可能になります。

    **注意**: 既存のオブジェクトは上書きされます。実行前にAccessファイルのバックアップを取ることを推奨します。
    """
    file_path = os.path.abspath(file_path)
    input_dir = os.path.abspath(input_dir)
    logger.info(f"load コマンドが実行されました。ファイルパス: {file_path}, 入力ディレクトリ: {input_dir}")
    if not os.path.exists(file_path):
        console.print(f"[bold red]エラー: ファイルが見つかりません: {file_path}[/bold red]")
        logger.error(f"ファイルが見つかりません: {file_path}")
        raise typer.Exit(code=1)
    if not os.path.exists(input_dir):
        console.print(f"[bold red]エラー: 入力ディレクトリが見つかりません: {input_dir}[/bold red]")
        logger.error(f"入力ディレクトリが見つかりません: {input_dir}")
        raise typer.Exit(code=1)

    try:
        with access_application(file_path) as app:
            tree = Tree(f"[bold cyan]📦 {os.path.basename(file_path)}[/bold cyan] へのインポート結果", guide_style="bold bright_blue")
            imported_files = import_objects(app, input_dir)

            for category, files in imported_files.items():
                if files:
                    branch = tree.add(f"[green]{category}[/green] ({len(files)}件)")
                    for file in files:
                        branch.add(f"[white]{file}[/white]")
                        logger.info(f"インポート済み: カテゴリ={category}, ファイル={file}")
            
            console.print(tree)
            console.print(f"\n[bold green]✅ インポートが完了しました。[/bold green]")
            logger.info("インポートが完了しました。")

    except Exception as e:
        handle_com_error(e)
        logger.error(f"load コマンドの実行中にエラーが発生しました: {e}", exc_info=True)
        raise typer.Exit(code=1)