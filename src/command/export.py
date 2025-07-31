# -*- coding: utf-8 -*-
import os
import typer
from rich.console import Console
from rich.tree import Tree
import logging

from src.utils import handle_com_error
from src.core.access_handler import access_application, export_objects
from src.constants import BASE_APP_DIR

console = Console()
logger = logging.getLogger(__name__)

def export(file_path: str = typer.Argument(..., help="エクスポート対象のAccessファイルのパス"), 
           output_dir: str = typer.Option(os.path.join(BASE_APP_DIR, "output", "export"), "--output", "-o", help="エクスポートされたオブジェクトの保存先ディレクトリ。デフォルトは `./output/export` です。")):
    """
    指定されたAccessファイル（.accdbまたは.mdb）から、オブジェクトをテキストファイルとしてエクスポートします。

    エクスポートされるオブジェクトの種類:
    - フォーム (.frm)
    - レポート (.rpt)
    - マクロ (.mcr)
    - モジュール (.bas)
    - クエリ (.qry)

    これらのオブジェクトは、指定された出力ディレクトリにそれぞれのファイルとして保存されます。
    これにより、バージョン管理システムでの管理や、他のAccessファイルへのインポートが容易になります。
    """
    file_path = os.path.abspath(file_path)
    output_dir = os.path.abspath(output_dir)
    logger.info(f"export コマンドが実行されました。ファイルパス: {file_path}, 出力ディレクトリ: {output_dir}")
    if not os.path.exists(file_path):
        console.print(f"[bold red]エラー: ファイルが見つかりません: {file_path}[/bold red]")
        logger.error(f"ファイルが見つかりません: {file_path}")
        raise typer.Exit(code=1)

    if os.path.exists(output_dir):
        import shutil
        shutil.rmtree(output_dir)
        logger.info(f"既存の出力ディレクトリをクリアしました: {output_dir}")
    os.makedirs(output_dir)
    logger.info(f"出力ディレクトリを作成しました: {output_dir}")

    try:
        with access_application(file_path) as app:
            tree = Tree(f"[bold cyan]📦 {os.path.basename(file_path)}[/bold cyan] のエクスポート結果", guide_style="bold bright_blue")
            exported_files = export_objects(app, output_dir)

            for category, files in exported_files.items():
                if files:
                    branch = tree.add(f"[green]{category}[/green] ({len(files)}件)")
                    for file in files:
                        branch.add(f"[white]{file}[/white]")
                        logger.info(f"エクスポート済み: カテゴリ={category}, ファイル={file}")
            
            console.print(tree)
            console.print(f"\n[bold green]✅ エクスポートが完了しました: {os.path.abspath(output_dir)}[/bold green]")
            logger.info(f"エクスポートが完了しました: {os.path.abspath(output_dir)}")

    except Exception as e:
        handle_com_error(e)
        logger.error(f"export コマンドの実行中にエラーが発生しました: {e}", exc_info=True)
        raise typer.Exit(code=1)