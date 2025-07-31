# -*- coding: utf-8 -*-
import os
import shutil
import typer
from rich.console import Console
import logging

from src.utils import handle_com_error
from src.core.access_handler import access_application, release_prepare, update_linked_table_paths

console = Console()
logger = logging.getLogger(__name__)

def prepare_release(
    file_path: str = typer.Argument(..., help="リリース準備を行うAccessファイルのパス（開発用ファイル）"), 
    output_file: str = typer.Argument(..., help="リリース用に最適化されたAccessファイルの出力パス"),
    test_conn_str: str = typer.Option(..., "--test-conn", help="VBAコード内で置換対象となるテスト環境の接続文字列"),
    prod_conn_str: str = typer.Option(..., "--prod-conn", help="テスト接続文字列が置換される本番環境の接続文字列"),
    old_linked_path: str = typer.Option(None, "--old-linked-path", help="リンクテーブルの古いパス"),
    new_linked_path: str = typer.Option(None, "--new-linked-path", help="リンクテーブルの新しいパス"),
):
    """
    Accessファイルを配布用に最適化し、本番環境へのデプロイを容易にするための準備を行います。

    このコマンドは以下の処理を実行します。
    1.  **ファイルのコピー**: 元のAccessファイルを変更せず、指定された出力パスに新しいファイルをコピーします。
    2.  **接続文字列の置換**: VBAモジュールやフォーム内の指定されたテスト用接続文字列を、本番用接続文字列に自動的に置換します。
    3.  **デバッグコードの除去**: VBAコード内の`Debug.Print`文などをコメントアウトし、本番環境での不要な出力を防ぎます。
    4.  **リンクテーブルパスの更新**: オプションで指定された場合、リンクテーブルの接続文字列のパスを更新します。
    5.  **データベースの最適化**: Accessデータベースの「最適化と修復」を実行し、ファイルサイズを削減しパフォーマンスを向上させます。

    これにより、開発環境と本番環境で異なる接続情報を使用している場合でも、
    手動でのVBAコード修正なしに、安全かつ効率的にリリース用ファイルを作成できます。

    **注意**: この操作はVBAコードを直接変更します。実行前に必ず元のAccessファイルのバックアップを取ることを推奨します。
    """
    file_path = os.path.abspath(file_path)
    output_file = os.path.abspath(output_file)
    logger.info(f"prepare-release コマンドが実行されました。開発ファイル: {file_path}, 出力ファイル: {output_file}")
    if not os.path.exists(file_path):
        console.print(f"[bold red]エラー: ファイルが見つかりません: {file_path}[/bold red]")
        logger.error(f"ファイルが見つかりません: {file_path}")
        raise typer.Exit(code=1)

    console.print(f"[cyan]リリース準備を開始します: {os.path.basename(file_path)} -> {os.path.basename(output_file)}[/cyan]")
    logger.info(f"リリース準備を開始します: {file_path} -> {output_file}")
    
    # 1. ファイルをコピー
    try:
        shutil.copy2(file_path, output_file)
        console.print(f"[green]✓[/green] ファイルをコピーしました。")
        logger.info("ファイルをコピーしました。")
    except Exception as e:
        console.print(f"[bold red]エラー: ファイルのコピーに失敗しました: {e}[/bold red]")
        logger.error(f"ファイルのコピーに失敗しました: {e}", exc_info=True)
        raise typer.Exit(code=1)

    # 2. 最適化処理
    try:
        with access_application(output_file) as app:
            with console.status("[bold green]VBAコードの最適化中...[/]"):
                modifications = release_prepare(app, test_conn_str, prod_conn_str)
            
            console.print("[green]✓[/green] 最適化処理が完了しました。")
            logger.info("最適化処理が完了しました。")
            if modifications:
                console.print("  [dim]- 接続文字列を置換しました。[/dim]")
                console.print("  [dim]- デバッグコードをコメントアウトしました。[/dim]")
                logger.info("接続文字列を置換し、デバッグコードをコメントアウトしました。")
            else:
                console.print("  [dim]- コードに変更はありませんでした。[/dim]")
                logger.info("コードに変更はありませんでした。")

            # 3. リンクテーブルパスの更新
            if old_linked_path and new_linked_path:
                with console.status("[bold green]リンクテーブルパスを更新中...[/]"):
                    updated_links = update_linked_table_paths(app, old_linked_path, new_linked_path)
                console.print(f"[green]✓[/green] リンクテーブルパスを更新しました ({updated_links}件)。")
                logger.info(f"リンクテーブルパスを更新しました ({updated_links}件)。")
            elif old_linked_path or new_linked_path:
                console.print("[yellow]警告: リンクテーブルパスの更新には --old-linked-path と --new-linked-path の両方が必要です。スキップしました。[/yellow]")
                logger.warning("リンクテーブルパスの更新には --old-linked-path と --new-linked-path の両方が必要です。スキップしました。")

            # 4. コンパクト化と修復
            with console.status("[bold green]データベースの最適化中...[/]"):
                app.DoCmd.RunCommand(9) # acCmdCompactDatabase
            console.print("[green]✓[/green] データベースを最適化しました。")
            logger.info("データベースを最適化しました。")

        console.print(f"\n[bold green]✅ リリース準備が完了しました: {output_file}[/bold green]")
        logger.info(f"リリース準備が完了しました: {output_file}")

    except Exception as e:
        handle_com_error(e)
        logger.error(f"prepare-release コマンドの実行中にエラーが発生しました: {e}", exc_info=True)
        # エラーが発生した場合、不完全な出力ファイルを削除
        if os.path.exists(output_file):
            os.remove(output_file)
            logger.info(f"エラー発生のため、不完全な出力ファイル {output_file} を削除しました。")
        raise typer.Exit(code=1)