[![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)](https://github.com/tomoya-asdf/AccessDevKit/blob/main/LICENSE)

# AccessDevKit

Microsoft Access 開発を支援するためのCLIツールセットです。差分比較、ファイル展開、オブジェクトのエクスポート/インポート、未使用オブジェクト分析、パフォーマンス測定、リリース準備、オブジェクト内検索など、Access開発の様々なタスクを効率化します。

**主な特徴**

*   **ファイルロックの自動検出**: コマンド実行時にAccessファイルが開かれている場合、エラーメッセージを表示して処理を中断します。
*   **対話モード**: 初心者でも簡単に操作できる対話モードを搭載しています。
*   **実行ファイル（exe）のビルド**: `pyinstaller` を使って、Python環境がないPCでも実行できるexeファイルを簡単にビルドできます。

## フォルダ構造

```
. 
├── src/
│   ├── main.py             # CLIのエントリーポイントとコマンド定義
│   ├── command/            # 各CLIコマンドの実装
│   │   ├── diff.py         # Accessファイルの差分比較
│   │   ├── deploy.py       # Accessファイルの展開（上書き）
│   │   ├── export.py       # Accessオブジェクトのエクスポート
│   │   ├── load.py         # Accessオブジェクトのインポート
│   │   ├── analyze_usage.py# 未使用オブジェクトの分析
│   │   ├── benchmark.py    # クエリ/フォームのパフォーマンス測定
│   │   ├── prepare_release.py # リリース準備（接続文字列置換、デバッグコード除去など）
│   │   └── search.py       # Accessオブジェクト内のキーワード検索
│   ├── core/               # コアロジック（Access COM操作、DB操作、レポート生成など）
│   │   ├── access_handler.py # AccessアプリケーションとのCOM連携
│   │   ├── db_operations.py  # データベース操作（pyodbc）
│   │   └── reporting.py    # レポート生成（Excel）
│   ├── logs/                   # ログファイル出力ディレクトリ
│   ├── reports/                # レポートなどの出力ディレクトリ
│   ├── templates/              # テンプレートファイルディレクトリ
│   ├── constants.py        # 定数定義
│   └── utils.py            # 共通ユーティリティ関数（エラーハンドリング、Excel用サニタイズなど）
├── build.bat               # 実行ファイル（exe）をビルドするためのバッチファイル
├── requirements.txt        # 依存ライブラリ
└── README.md               # このファイル
```

## インストール

1.  **Pythonのインストール**: Python 3.8以上がインストールされていることを確認してください。
2.  **仮想環境の作成と有効化（推奨）**:
    ```bash
    python -m venv env
    # Windows
    .\env\Scripts\activate
    # macOS/Linux
    source env/bin/activate
    ```
3.  **依存ライブラリのインストール**:
    ```bash
    pip install -r requirements.txt
    ```
4.  **pywin32のインストール**: `pywin32`は特殊なインストールが必要です。
    ```bash
    pip install pywin32
    python -m win32com.client.makepy -i "Microsoft Access 16.0 Object Library"
    ```

## 使い方

ツールは対話モードと直接コマンド実行の両方に対応しています。

### 対話モード

引数なしで`main.py`を実行すると、対話モードが起動し、利用可能なコマンドが一覧表示されます。番号を選択して実行できます。

```bash
python src/main.py
```

**実行例:**

```
╔══════════════════════════════════════════════════════════════════════════════╗
║                Access Development Tool - コマンド選択                         ║
╠══════════════════════════════════════════════════════════════════════════════╣
║                                                                              ║
║   [1] diff: 2つのAccessファイルの差分を比較し、HTMLレポートを生成します。         ║
║   [2] deploy: Accessファイルを指定ディレクトリ内の同名ファイルに展開します。      ║
║   [3] export: Accessファイルから全ての主要オブジェクトをテキストファイルとしてエクスポートします。 ║
║   [4] load: テキストファイルをAccessファイルにインポートします。                 ║
║   [5] analyze-usage: Accessファイル内の未使用オブジェクトを分析します。          ║
║   [6] benchmark: 指定されたクエリの実行時間を計測します。                        ║
║   [7] prepare-release: Accessファイルを配布用に最適化します。                   ║
║   [8] search: Accessファイル内の全オブジェクトからキーワードを検索します。        ║
║                                                                              ║
╚══════════════════════════════════════════════════════════════════════════════╝
```

### 直接コマンド実行

各コマンドは、`python src/main.py <command> [options]` の形式で直接実行できます。

#### グローバルオプション

*   `--debug`: デバッグモードを有効にし、詳細なエラー情報を表示します。

#### コマンド一覧

##### `diff`

2つのAccessファイルの差分を比較し、HTMLレポートを生成します。

```bash
python src/main.py diff <file1_path> <file2_path>
```

*   `<file1_path>`: 比較対象のAccessファイル1のパス
*   `<file2_path>`: 比較対象のAccessファイル2のパス

**出力**: 比較結果は`reports/access_diff_report.html`にHTML形式で出力され、自動的にブラウザで開かれます。

##### `deploy`

Accessファイルを指定ディレクトリ内の同名ファイルに展開（上書き）します。

```bash
python src/main.py deploy <source_file> <target_dir>
```

*   `<source_file>`: 展開元となるAccessファイルのパス
*   `<target_dir>`: 展開先のディレクトリパス

##### `export`

Accessファイルから全ての主要オブジェクト（フォーム、レポート、マクロ、モジュール、クエリ）をテキストファイルとしてエクスポートします。

```bash
python src/main.py export <file_path> [--output <output_dir>]
```

*   `<file_path>`: エクスポート対象のAccessファイルパス
*   `--output`, `-o` (オプション): オブジェクトの出力先ディレクトリ（デフォルト: `./export`）

##### `load`

指定されたディレクトリから、`export`コマンドでエクスポートされたテキストファイルをAccessファイルにインポート（上書き）します。

```bash
python src/main.py load <file_path> [--input <input_dir>]
```

*   `<file_path>`: インポート対象のAccessファイルパス
*   `--input`, `-i` (オプション): オブジェクトが格納されているディレクトリ（デフォルト: `./export`）

##### `analyze-usage`

Accessファイル内の未使用オブジェクトを分析し、HTMLレポートを生成します。（実験的機能）

```bash
python src/main.py analyze-usage <file_path>
```

*   `<file_path>`: 分析対象のAccessファイルパス

**出力**: 分析結果は`reports/unused_objects_report.html`にHTML形式で出力され、自動的にブラウザで開かれます。

##### `benchmark`

指定されたクエリの実行時間を計測し、HTMLレポートを生成します。

```bash
python src/main.py benchmark <file_path> [--query <query_name_1> ...] [--runs <num_runs>]
```

*   `<file_path>`: 対象のAccessファイルパス
*   `--query`, `-q` (オプション): 測定対象のクエリ名（複数指定可）。指定しない場合、全てのクエリを測定します。
*   `--runs`, `-r` (オプション): 各クエリの実行回数（デフォルト: `5`）

**出力**: ベンチマーク結果は`reports/benchmark_report.html`にHTML形式で出力され、自動的にブラウザで開かれます。

##### `prepare-release`

Accessファイルを配布用に最適化し、接続文字列の置換やデバッグコードの除去を行います。

```bash
python src/main.py prepare-release <file_path> <output_file> --test-conn <test_conn_str> --prod-conn <prod_conn_str>
```

*   `<file_path>`: 開発用のAccessファイルパス
*   `<output_file>`: 出力するリリース用のファイルパス
*   `--test-conn` (必須): 置換前のテスト用接続文字列
*   `--prod-conn` (必須): 置換後の本番用接続文字列

##### `search`

Accessファイル内の全オブジェクト（VBAコード、フォーム、レポート、マクロ、クエリ、テーブルデータ）からキーワードを検索します。

```bash
python src/main.py search <file_path> <pattern>
```

*   `<file_path>`: 検索対象のAccessファイルパス
*   `<pattern>`: 検索キーワード

**出力**: 検索結果はコンソールに表示されます。

## 実行ファイル（exe）のビルド

`pyinstaller` を使用して、このツールを単一の実行ファイル（`.exe`）としてパッケージングできます。これにより、Pythonがインストールされていない環境でもツールを実行できます。

ビルドを実行するには、プロジェクトのルートディレクトリで`build.bat`を実行してください。

```bash
build.bat
```

ビルドが成功すると、`build/` ディレクトリに `AccessTools.exe` が生成されます。

## 開発者向け情報

### エラーハンドリング

COM連携中にエラーが発生した場合、ツールは詳細なエラーメッセージとCOMエラー情報を表示します。`--debug`フラグを使用すると、さらに詳細なPythonのスタックトレースを確認できます。

ファイルが他のプロセスによってロックされている場合、ツールは自動的にそれを検出し、以下の様なエラーメッセージを表示して安全に処理を中断します。

```
IOError: 対象のAccessファイルが開かれているため、処理を中断しました。ファイルを閉じてから再実行してください。
```

### 拡張性

`src/command`ディレクトリに新しいPythonファイルを追加し、`src/main.py`でインポートして`app.command()`で登録することで、新しいコマンドを簡単に追加できます。

`src/core`ディレクトリ内のモジュールは、Access操作、データベース操作、レポート生成といった共通のロジックを提供しており、新しい機能を追加する際に再利用できます。

## ライセンス

MIT License