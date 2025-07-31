# -*- coding: utf-8 -*-
import pyodbc
import contextlib
import time
from rich.console import Console
import os
from src.utils import is_file_locked

console = Console()

@contextlib.contextmanager
def db_connection(db_path):
    if is_file_locked(db_path):
        raise IOError("対象のAccessファイルが開かれているため、処理を中断しました。ファイルを閉じてから再実行してください。")
    conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path};'
    conn = pyodbc.connect(conn_str)
    try:
        yield conn
    finally:
        conn.close()

def get_table_names(conn):
    return [row.table_name for row in conn.cursor().tables(tableType='TABLE')]

def get_table_data(conn, table_name):
    cursor = conn.cursor()
    cursor.execute(f"SELECT * FROM [{table_name}]")
    return {tuple(row) for row in cursor.fetchall()}

def run_benchmark(conn, query_name, runs):
    timings = []
    cursor = conn.cursor()
    for _ in range(runs):
        start = time.perf_counter()
        cursor.execute(f"SELECT * FROM [{query_name}]")
        cursor.fetchall()
        end = time.perf_counter()
        timings.append(end - start)
    return timings

def search_in_tables(conn, pattern):
    results = []
    table_names = get_table_names(conn)
    for table_name in table_names:
        if pattern.lower() in table_name.lower():
            results.append({
                "type": "Table Name", "name": table_name,
                "line_num": None, "column_name": None, "line_content": table_name
            })
        try:
            cursor = conn.cursor()
            cursor.execute(f"SELECT * FROM [{table_name}]")
            columns = [column[0] for column in cursor.description]
            for row_num, row in enumerate(cursor.fetchall(), 1):
                for col_num, cell_value in enumerate(row):
                    if isinstance(cell_value, str) and pattern.lower() in cell_value.lower():
                        results.append({
                            "type": "Table Data",
                            "name": table_name,
                            "line_num": row_num,
                            "column_name": columns[col_num],
                            "line_content": str(row)
                        })
        except pyodbc.Error as e:
            console.print(f"[yellow]警告: テーブル '{table_name}' の読み取り中にエラーが発生しました: {e}[/yellow]")
            continue
    return results
