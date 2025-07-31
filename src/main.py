# -*- coding: utf-8 -*-
import sys
import os

# Add the parent directory of 'src' to sys.path
# This allows imports like 'from src.command.diff import diff' to work
# when main.py is run directly from various locations.
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.abspath(os.path.join(current_dir, os.pardir))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

import typer
import inspect
import sys
import logging
import traceback
from datetime import datetime
from rich.console import Console
from rich.prompt import Prompt, Confirm
from rich.text import Text
from rich.table import Table
from rich.align import Align

# --- コマンドのインポート ---
from src.command.diff import diff
from src.command.deploy import deploy
from src.command.export import export
from src.command.load import load
from src.command.analyze_usage import analyze_usage
from src.command.benchmark import benchmark
from src.command.prepare_release import prepare_release
from src.command.search import search

# --- アプリケーションのセットアップ ---
app = typer.Typer(
    name="Access Tool",
    help="Microsoft Access 開発を支援するためのCLIツールセットです。",
    add_completion=False,
    rich_markup_mode="markdown"
)
console = Console()

# --- ロギング設定 ---
from src.constants import LOG_DIR, LOG_FILE_NAME_ALL
from src.utils import setup_logging

# --- ロギング設定 ---
logger = setup_logging(logging.DEBUG, console)
logger.debug("Application started.")
logger.debug(f"Log file path: {os.path.join(LOG_DIR, LOG_FILE_NAME_ALL.format(datetime=datetime.now().strftime('%Y%m%d_%H%M%S')))}")

# --- コマンド登録 ---
app.command(name="diff")(diff)
app.command(name="deploy")(deploy)
app.command(name="export")(export)
app.command(name="load")(load)
app.command(name="analyze-usage")(analyze_usage)
app.command(name="benchmark")(benchmark)
app.command(name="prepare-release")(prepare_release)
app.command(name="search")(search)

# --- 対話モード ---
def run_interactive_mode(ctx: typer.Context):
    console.rule("[bold blue]対話モード[/bold blue]")
    ascii_art = """
    _                         ____             _  ___ _   
   / \   ___ ___ ___  ___ ___|  _ \  _____   _| |/ (_) |_ 
  / _ \ / __/ __/ _ \/ __/ __| | | |/ _ \ \ / / ' /| | __|
 / ___ \ (_| (_|  __/\__ \__ \ |_| |  __/\ V /| . \| | |_ 
/_/   \_\___\___\___||___/___/____/ \___| \_/ |_|\_\_|\__|
                                                          
"""
    all_commands = [cmd for cmd in app.registered_commands if cmd.name != 'interactive']
    sorted_commands = sorted(all_commands, key=lambda cmd: cmd.name)

    clear_command = 'cls' if os.name == 'nt' else 'clear'

    try:
        while True:
            os.system(clear_command)

            title = Align(Text(ascii_art, style="bold"), align="left")
            table = Table(title=" 利用可能なコマンド", title_justify="left", show_header=True, header_style="bold")
            table.add_column("番号", style="cyan", justify="right")
            table.add_column("コマンド", style="cyan bold")
            table.add_column("説明")

            for i, cmd_info in enumerate(sorted_commands, 1):
                doc = inspect.getdoc(cmd_info.callback) or ""
                brief_doc = doc.strip().split('\n')[0]
                table.add_row(str(i), cmd_info.name, brief_doc)

            console.print(title)
            console.print(table)
            choice = Prompt.ask("\n実行するコマンド番号を入力 ('q'で終了)")

            if choice.lower() == 'q':
                break
            if not choice.strip(): # 空の入力の場合
                console.print("[bold red]エラー: コマンド番号を入力してください。[/bold red]")
                continue # 再度プロンプトを表示

            try:
                idx = int(choice) - 1
                if not (0 <= idx < len(sorted_commands)): 
                    raise ValueError("無効なコマンド番号です。")
                selected_cmd = sorted_commands[idx]

                # CommandInfo オブジェクトから TyperCommand オブジェクトを取得し、パラメータにアクセス
                collected_args = collect_args(selected_cmd.name, inspect.signature(selected_cmd.callback).parameters.values())
                if collected_args is not None:
                    # 破壊的な操作に対する確認
                    if selected_cmd.name == "deploy":
                        if not Confirm.ask(f"[bold yellow]警告: コマンド '{selected_cmd.name}' は、展開先のディレクトリにある同名のファイルを上書きします。続行しますか？[/bold yellow]"):
                            console.print("[bold red]コマンドの実行がキャンセルされました。[/bold red]")
                            continue # 次のループへ
                    elif selected_cmd.name == "prepare-release":
                        if not Confirm.ask(f"[bold yellow]警告: コマンド '{selected_cmd.name}' は、指定されたファイルに接続文字列の置換やデバッグコードの除去を行い、新しいファイルを生成します。続行しますか？[/bold yellow]"):
                            console.print("[bold red]コマンドの実行がキャンセルされました。[/bold red]")
                            continue # 次のループへ
                    
                    console.print("") # コマンド実行前に改行
                    ctx.invoke(selected_cmd.callback, **collected_args)
                    console.print(f"[bold green]コマンド '{selected_cmd.name}' が実行されました。[/bold green]")
            except (ValueError, TypeError) as e:
                console.print(f"[bold red]エラー: {e}[/bold red]")
                logger.error(f"対話モードでのエラー: {e}", exc_info=True)
            except Exception as e:
                console.print(f"[bold red]予期せぬエラー: {e}[/bold red]")
                logger.error(f"対話モードでの予期せぬエラー: {e}", exc_info=True)

            if not Confirm.ask("\n続けて他のコマンドを実行しますか？"):
                break
    except KeyboardInterrupt:
        #os.system(clear_command)
        console.print("[bold yellow]Ctrl+Cが押されました。[/bold yellow]")

    console.print("対話モードを終了します。")

def collect_args(command_name: str, params: list[typer.models.ParameterInfo]) -> dict:
    kwargs = {}
    console.print(f"[bold]コマンド '[cyan]{command_name}[/cyan]' のパラメータを入力:[/bold]")
    
    for param_info in params:
        param_name = param_info.name
        param_type = param_info.annotation if param_info.annotation != inspect.Parameter.empty else str
        default_val = param_info.default if param_info.default != inspect.Parameter.empty else ...
        param_help = ""
        if isinstance(param_info.default, (typer.models.OptionInfo, typer.models.ArgumentInfo)):
            param_help = param_info.default.help or ""

        prompt_parts = [f"  [yellow]{param_name}[/yellow]"]
        if param_type != str: # Add type hint if not string
            prompt_parts.append(f"[dim]({param_type.__name__})[/dim]")
        if param_help:
            prompt_parts.append(f"[dim]（{param_help}）[/dim]")
        display_default = default_val
        if isinstance(default_val, (typer.models.OptionInfo, typer.models.ArgumentInfo)):
            display_default = default_val.default

        if display_default is not ... and display_default is not None:
            prompt_parts.append(f"[dim]（デフォルト: {display_default}）[/dim]")
        prompt_text = " ".join(prompt_parts)

        value_str = Prompt.ask('\n'+prompt_text, default=str(display_default) if display_default is not ... else None)

        try:
            # If the parameter is required and no value was provided
            if display_default is ... and (value_str is None or value_str == ""):
                console.print(f"[bold red]エラー: パラメータ '{param_name}' は必須です。値を入力してください。[/bold red]")
                logger.error(f"必須パラメータ '{param_name}' が入力されませんでした。")
                return None # Indicate failure to collect args

            # If the parameter is optional and the user entered nothing (empty string or None)
            # then apply the default value.
            if display_default is not ... and (value_str is None or value_str == ""):
                kwargs[param_name] = display_default
            elif param_type == bool:
                kwargs[param_name] = value_str.lower() in ['true', '1', 'y', 'yes']
            else:
                # Handle list[str] specifically
                import typing
                origin_type = typing.get_origin(param_type)
                args_type = typing.get_args(param_type)

                if origin_type is list and args_type == (str,):
                    kwargs[param_name] = [s.strip() for s in value_str.split(',') if s.strip()]
                else:
                    kwargs[param_name] = param_type(value_str)
        except (ValueError, TypeError) as e:
            console.print(f"[bold red]エラー: '{value_str}'を型 {param_type.__name__} に変換できませんでした。[/bold red]")
            logger.error(f"パラメータ変換エラー: '{value_str}'を型 {param_type.__name__} に変換できませんでした。", exc_info=True)
            return None # Return None on error
    return kwargs

# --- メインコールバック ---
@app.callback(invoke_without_command=True)
def main(ctx: typer.Context, debug: bool = typer.Option(False, "--debug", help="デバッグモードを有効にし、詳細なエラー情報を表示します。")):
    if debug:
        console.print("[bold yellow]デバッグモードが有効です。[/bold yellow]")
        # You can set a global variable or a context object attribute here
        # e.g., ctx.obj = {"debug": True}

    if ctx.invoked_subcommand is None:
        run_interactive_mode(ctx)

# --- エントリーポイント ---
if __name__ == "__main__":
    try:
        app()
    except KeyboardInterrupt:
        console.print("\n[bold yellow]Ctrl+Cが押されました。アプリケーションを終了します。[/bold yellow]")
    except Exception as e:
        console.print(f"[bold red]予期せぬエラーが発生しました: {e}[/bold red]")
        logger.error(f"トップレベルでの予期せぬエラー: {e}", exc_info=True)
        if "--debug" in sys.argv:
            console.print("[bold yellow]デバッグモードのため、詳細なトレースバックを表示します。[/bold yellow]")
            traceback.print_exc()
