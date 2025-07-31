# -*- coding: utf-8 -*>
import os
import typer
from rich.console import Console
from rich.table import Table
from datetime import datetime
import webbrowser
import logging

from src.utils import handle_com_error
from src.core.access_handler import access_application, analyze_usage as core_analyze_usage
from src.core.reporting import ReportGenerator
from src.constants import UNUSED_OBJECTS_REPORT_PATH

console = Console()
logger = logging.getLogger(__name__)

def analyze_usage(file_path: str = typer.Argument(..., help="分析対象のAccessファイルのパス")):
    """
    指定されたAccessファイル（.accdbまたは.mdb）内の未使用の可能性のあるオブジェクトを分析し、HTMLレポートを生成します。

    この機能は、VBAコード、フォーム、レポート、マクロ、クエリなどのオブジェクト間で、
    明示的に参照されていないオブジェクトを検出するのに役立ちます。
    これにより、データベースの整理や最適化に貢献します。

    **注意点**:
    - この機能は実験的なものであり、動的な参照（例: `CallByName`、`Eval`、ナビゲーションフォームからの参照）や
      外部からの参照は検出できません。そのため、検出されたオブジェクトが必ずしも未使用であるとは限りません。
    - .accdeファイルは分析できません。

    分析結果は `reports/unused_objects_report.html` にHTML形式で出力され、完了後に自動で開かれます。
    """
    file_path = os.path.abspath(file_path)
    logger.info(f"analyze-usage コマンドが実行されました。ファイルパス: {file_path}")
    if not os.path.exists(file_path):
        console.print(f"[bold red]エラー: ファイルが見つかりません: {file_path}[/bold red]")
        logger.error(f"ファイルが見つかりません: {file_path}")
        raise typer.Exit(code=1)

    # 固定のHTML出力パス
    html_output_path = UNUSED_OBJECTS_REPORT_PATH

    report_generator = ReportGenerator()

    try:
        with access_application(file_path) as app:
            console.print(f"[cyan]オブジェクトの参照状況を分析中...（時間がかかる場合があります）[/cyan]")
            logger.info("オブジェクトの参照状況を分析中...")
            with console.status("[bold green]分析中...[/]"):
                unused_objects = core_analyze_usage(app)

            if not unused_objects:
                console.print("[bold green]✅ 未使用の可能性が高いオブジェクトは見つかりませんでした。[/bold green]")
                logger.info("未使用の可能性が高いオブジェクトは見つかりませんでした。")
                # 変更がない場合でもHTMLレポートを生成
                report_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                report_generator.create_unused_objects_report([], html_output_path, report_datetime, file_path)
                console.print(f"\n[bold green]✅ 未使用オブジェクトレポートを '{html_output_path}' に出力しました。[/bold green]")
                logger.info(f"未使用オブジェクトレポートを '{html_output_path}' に出力しました。")
                webbrowser.open(os.path.abspath(html_output_path))
                return

            console.print("\n[yellow]⚠️ 以下のオブジェクトは、どこからも参照されていない可能性があります。[/yellow]")
            console.print("[dim]（動的な呼び出しや、ナビゲーションフォームからの参照は検出できません）[/dim]")
            logger.warning("未使用の可能性があるオブジェクトが見つかりました。")

            table = Table(title="未使用の可能性があるオブジェクト", title_justify="left", show_header=True, header_style="bold ")
            table.add_column("オブジェクト種類", style="cyan")
            table.add_column("オブジェクト名", style="green")

            for obj_type, obj_name in unused_objects:
                table.add_row(obj_type, obj_name)
                logger.info(f"未使用オブジェクト: 種類={obj_type}, 名前={obj_name}")
            
            console.print(table)

            report_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            report_generator.create_unused_objects_report(unused_objects, html_output_path, report_datetime, file_path)
            console.print(f"\n[bold green]✅ 未使用オブジェクトレポートを '{html_output_path}' に出力しました。[/bold green]")
            logger.info(f"未使用オブジェクトレポートを '{html_output_path}' に出力しました。")
            webbrowser.open(os.path.abspath(html_output_path))

    except Exception as e:
        handle_com_error(e)
        logger.error(f"analyze-usage コマンドの実行中にエラーが発生しました: {e}", exc_info=True)
        raise typer.Exit(code=1)
