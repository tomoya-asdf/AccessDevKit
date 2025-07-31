# -*- coding: utf-8 -*>
import os
import typer
import time
from rich.console import Console
from rich.table import Table
from datetime import datetime
import webbrowser
import logging

from src.utils import handle_com_error
from src.core.db_operations import db_connection, run_benchmark as core_run_benchmark
from src.core.reporting import ReportGenerator
from src.constants import BENCHMARK_REPORT_PATH
from src.core.access_handler import access_application, get_access_query_names

console = Console()
logger = logging.getLogger(__name__)

def benchmark(
    file_path: str = typer.Argument(..., help="ベンチマーク対象のAccessファイルのパス"), 
    queries: str = typer.Option(None, "--query", "-q", help="測定対象のクエリ名（カンマ区切りで複数指定可）。指定しない場合、Accessファイル内の全てのクエリを測定します。"),
    runs: int = typer.Option(5, "--runs", "-r", help="各クエリの実行回数。デフォルトは5回です。")
):
    """
    指定されたAccessファイル（.accdbまたは.mdb）内のクエリの実行パフォーマンスを計測し、HTMLレポートを生成します。

    このコマンドは、データベースの最適化やパフォーマンスチューニングの際に役立ちます。
    - **クエリ指定**: 特定のクエリを指定してその実行時間を測定できます。複数のクエリをカンマ区切りで指定することも可能です。
    - **全クエリ測定**: クエリを指定しない場合、Accessファイル内の全てのクエリを自動的に検出し、それぞれの実行時間を測定します。
    - **複数回実行**: 各クエリを複数回実行し、平均実行時間を算出することで、より信頼性の高いパフォーマンスデータを提供します。

    ベンチマーク結果は、クエリ名、平均実行時間、合計実行時間を含む表形式で `reports/benchmark_report.html` にHTML形式で出力され、完了後に自動で開かれます。
    """
    file_path = os.path.abspath(file_path)
    logger.info(f"benchmark コマンドが実行されました。ファイルパス: {file_path}, クエリ: {queries}, 実行回数: {runs}")
    if not os.path.exists(file_path):
        console.print(f"[bold red]エラー: ファイルが見つかりません: {file_path}[/bold red]")
        logger.error(f"ファイルが見つかりません: {file_path}")
        raise typer.Exit(code=1)

    # 固定のHTML出力パス
    html_output_path = BENCHMARK_REPORT_PATH

    report_generator = ReportGenerator()

    results = []
    try:
        # クエリが指定されていない場合、Access COMオブジェクト経由でクエリ名を取得
        if queries and queries.lower() != "none":
            queries_to_benchmark = [q.strip() for q in queries.split(',')]
        else:
            console.print("[cyan]クエリが指定されていません。全てのクエリを測定します。[/cyan]")
            logger.info("クエリが指定されていません。全てのクエリを測定します。")
            with access_application(file_path) as app:
                queries_to_benchmark = get_access_query_names(app)
            
            if not queries_to_benchmark:
                console.print("[yellow]警告: データベース内に測定可能なクエリが見つかりませんでした。[/yellow]")
                logger.warning("データベース内に測定可能なクエリが見つかりませんでした。")
                return

        with db_connection(file_path) as conn:
            console.print(f"[cyan]ベンチマークを開始します（実行回数: {runs}回）[/cyan]")
            logger.info(f"ベンチマークを開始します（実行回数: {runs}回）")

            table = Table(title="ベンチマーク結果", title_justify="left", show_header=True, header_style="bold ")
            table.add_column("クエリ名", style="green")
            table.add_column("平均実行時間 (秒)", style="yellow", justify="right")
            table.add_column("合計実行時間 (秒)", style="dim", justify="right")

            for query_name in queries_to_benchmark:
                with console.status(f"[bold green]クエリ '{query_name}' を実行中...[/]"):
                    try:
                        timings = core_run_benchmark(conn, query_name, runs)
                        
                        avg_time = sum(timings) / len(timings)
                        total_time = sum(timings)
                        results.append((query_name, avg_time, total_time))
                        table.add_row(query_name, f"{avg_time:.4f}", f"{total_time:.4f}")
                        logger.info(f"クエリ '{query_name}': 平均実行時間={avg_time:.4f}秒, 合計実行時間={total_time:.4f}秒")

                    except Exception as e:
                        console.print(f"[bold red]クエリ '{query_name}' の実行中にエラーが発生しました: {e}[/bold red]")
                        logger.error(f"クエリ '{query_name}' の実行中にエラーが発生しました: {e}", exc_info=True)
            
            console.print(table)

            report_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            report_generator.create_benchmark_report(results, html_output_path, report_datetime, file_path)
            console.print(f"\n[bold green]✅ ベンチマークレポートを '{html_output_path}' に出力しました。[/bold green]")
            logger.info(f"ベンチマークレポートを '{html_output_path}' に出力しました。")
            webbrowser.open(os.path.abspath(html_output_path))

    except Exception as e:
        handle_com_error(e)
        logger.error(f"benchmark コマンドの実行中にエラーが発生しました: {e}", exc_info=True)
        raise typer.Exit(code=1)