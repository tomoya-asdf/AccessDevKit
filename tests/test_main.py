from typer.testing import CliRunner
from src.main import app
import os

runner = CliRunner()

def test_app_help():
    """
    アプリケーションのヘルプメッセージが正しく表示されることをテストします。
    """
    result = runner.invoke(app, ["--help"])
    assert result.exit_code == 0
    assert "Usage: main [OPTIONS] COMMAND [ARGS]..." in result.stdout
    assert "Microsoft Access 開発を支援するためのCLIツールセットです。" in result.stdout

def test_commands_listed_in_help():
    """
    登録されているすべてのコマンドがヘルプメッセージに表示されることをテストします。
    """
    result = runner.invoke(app, ["--help"])
    assert result.exit_code == 0
    commands = [
        "diff",
        "deploy",
        "export",
        "load",
        "analyze-usage",
        "benchmark",
        "prepare-release",
        "search",
    ]
    for command in commands:
        assert command in result.stdout

# --- diff コマンドのテスト ---
def test_diff_command_help():
    """
    diffコマンドのヘルプメッセージが正しく表示されることをテストします。
    """
    result = runner.invoke(app, ["diff", "--help"])
    assert result.exit_code == 0
    assert "Usage: main diff [OPTIONS] FILE1_PATH FILE2_PATH" in result.stdout
    assert "2つのAccessファイルの差分を比較し、Excelレポートを生成します。" in result.stdout

def test_diff_command_missing_args():
    """
    diffコマンドが必須引数なしで呼び出されたときにエラーを報告することをテストします。
    """
    result = runner.invoke(app, ["diff"])
    assert result.exit_code != 0 # 必須引数不足は通常エラー終了
    assert "Missing argument 'FILE1_PATH'" in result.stderr or \
           "Missing argument 'FILE2_PATH'" in result.stderr

# --- deploy コマンドのテスト ---
def test_deploy_command_help():
    """
    deployコマンドのヘルプメッセージが正しく表示されることをテストします。
    """
    result = runner.invoke(app, ["deploy", "--help"])
    assert result.exit_code == 0
    assert "Usage: main deploy [OPTIONS] SOURCE_FILE TARGET_DIR" in result.stdout
    assert "Accessファイルを指定ディレクトリ内の同名ファイルに展開（上書き）します。" in result.stdout

def test_deploy_command_missing_args():
    """
    deployコマンドが必須引数なしで呼び出されたときにエラーを報告することをテストします。
    """
    result = runner.invoke(app, ["deploy"])
    assert result.exit_code != 0
    assert "Missing argument 'SOURCE_FILE'" in result.stderr or \
           "Missing argument 'TARGET_DIR'" in result.stderr

# --- export コマンドのテスト ---
def test_export_command_help():
    """
    exportコマンドのヘルプメッセージが正しく表示されることをテストします。
    """
    result = runner.invoke(app, ["export", "--help"])
    assert result.exit_code == 0
    assert "Usage: main export [OPTIONS] FILE_PATH" in result.stdout
    assert "Accessファイルから全ての主要オブジェクトをエクスポートします。" in result.stdout

def test_export_command_missing_args():
    """
    exportコマンドが必須引数なしで呼び出されたときにエラーを報告することをテストします。
    """
    result = runner.invoke(app, ["export"])
    assert result.exit_code != 0
    assert "Missing argument 'FILE_PATH'" in result.stderr

# --- load コマンドのテスト ---
def test_load_command_help():
    """
    loadコマンドのヘルプメッセージが正しく表示されることをテストします。
    """
    result = runner.invoke(app, ["load", "--help"])
    assert result.exit_code == 0
    assert "Usage: main load [OPTIONS] FILE_PATH" in result.stdout
    assert "ファイルからAccessへ全ての主要オブジェクトをインポートします。" in result.stdout

def test_load_command_missing_args():
    """
    loadコマンドが必須引数なしで呼び出されたときにエラーを報告することをテストします。
    """
    result = runner.invoke(app, ["load"])
    assert result.exit_code != 0
    assert "Missing argument 'FILE_PATH'" in result.stderr

# --- analyze-usage コマンドのテスト ---
def test_analyze_usage_command_help():
    """
    analyze-usageコマンドのヘルプメッセージが正しく表示されることをテストします。
    """
    result = runner.invoke(app, ["analyze-usage", "--help"])
    assert result.exit_code == 0
    assert "Usage: main analyze-usage [OPTIONS] FILE_PATH" in result.stdout
    assert "Accessファイル内の未使用オブジェクトを分析します。（実験的機能）" in result.stdout

def test_analyze_usage_command_missing_args():
    """
    analyze-usageコマンドが必須引数なしで呼び出されたときにエラーを報告することをテストします。
    """
    result = runner.invoke(app, ["analyze-usage"])
    assert result.exit_code != 0
    assert "Missing argument 'FILE_PATH'" in result.stderr

# --- benchmark コマンドのテスト ---
def test_benchmark_command_help():
    """
    benchmarkコマンドのヘルプメッセージが正しく表示されることをテストします。
    """
    result = runner.invoke(app, ["benchmark", "--help"])
    assert result.exit_code == 0
    assert "Usage: main benchmark [OPTIONS] FILE_PATH" in result.stdout
    assert "指定されたクエリの実行時間を計測します。" in result.stdout

def test_benchmark_command_missing_args():
    """
    benchmarkコマンドが必須引数なしで呼び出されたときにエラーを報告することをテストします。
    """
    result = runner.invoke(app, ["benchmark"])
    assert result.exit_code != 0
    assert "Missing argument 'FILE_PATH'" in result.stderr or \
           "Missing option '--query'" in result.stderr

# --- prepare-release コマンドのテスト ---
def test_prepare_release_command_help():
    """
    prepare-releaseコマンドのヘルプメッセージが正しく表示されることをテストします。
    """
    result = runner.invoke(app, ["prepare-release", "--help"])
    assert result.exit_code == 0
    assert "Usage: main prepare-release [OPTIONS] FILE_PATH OUTPUT_FILE" in result.stdout
    assert "Accessファイルを配布用に最適化し、接続文字列やデバッグコードを処理します。" in result.stdout

def test_prepare_release_command_missing_args():
    """
    prepare-releaseコマンドが必須引数なしで呼び出されたときにエラーを報告することをテストします。
    """
    result = runner.invoke(app, ["prepare-release"])
    assert result.exit_code != 0
    assert "Missing argument 'FILE_PATH'" in result.stderr or \
           "Missing argument 'OUTPUT_FILE'" in result.stderr or \
           "Missing option '--test-conn'" in result.stderr or \
           "Missing option '--prod-conn'" in result.stderr

# --- search コマンドのテスト ---
def test_search_command_help():
    """
    searchコマンドのヘルプメッセージが正しく表示されることをテストします。
    """
    result = runner.invoke(app, ["search", "--help"])
    assert result.exit_code == 0
    assert "Usage: main search [OPTIONS] FILE_PATH PATTERN" in result.stdout
    assert "Accessファイル内の全オブジェクトからキーワードを検索します。" in result.stdout

def test_search_command_missing_args():
    """
    searchコマンドが必須引数なしで呼び出されたときにエラーを報告することをテストします。
    """
    result = runner.invoke(app, ["search"])
    assert result.exit_code != 0
    assert "Missing argument 'FILE_PATH'" in result.stderr or \
           "Missing argument 'PATTERN'" in result.stderr
