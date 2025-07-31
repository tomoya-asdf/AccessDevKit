# -*- coding: utf-8 -*-
import os
import win32com.client
import win32com.client.gencache
win32com.client.gencache.EnsureDispatch("Access.Application")
import contextlib
import tempfile
import shutil
from src.utils import handle_com_error, sanitize_for_excel, is_file_locked
from src.core.db_operations import db_connection, search_in_tables

OBJECT_TYPES = {
    "Forms": win32com.client.constants.acForm,
    "Reports": win32com.client.constants.acReport,
    "Macros": win32com.client.constants.acMacro,
    "Modules": win32com.client.constants.acModule,
    "Queries": win32com.client.constants.acQuery,
}
OBJECT_EXTENSIONS = {
    "Forms": ".frm", "Reports": ".rpt", "Macros": ".mcr", "Modules": ".bas", "Queries": ".qry",
}

@contextlib.contextmanager
def temporary_access_copy(original_path):
    if is_file_locked(original_path):
        raise IOError("対象のAccessファイルが開かれているため、処理を中断しました。ファイルを閉じてから再実行してください。")
    temp_dir = tempfile.mkdtemp()
    temp_path = os.path.join(temp_dir, os.path.basename(original_path))
    shutil.copy2(original_path, temp_path)
    try:
        yield temp_path, temp_dir
    finally:
        shutil.rmtree(temp_dir)


@contextlib.contextmanager
def access_application(db_path):
    if is_file_locked(db_path):
        raise IOError("対象のAccessファイルが開かれているため、処理を中断しました。ファイルを閉じてから再実行してください。")
    app = win32com.client.Dispatch("Access.Application")
    app.Visible = False
    app.OpenCurrentDatabase(db_path)
    try:
        yield app
    finally:
        app.CloseCurrentDatabase()
        app.Quit()

def export_objects(app, export_dir):
    exported_files = {category: [] for category in OBJECT_TYPES.keys()}
    for category, obj_type in OBJECT_TYPES.items():
        if category == "Queries":
            collection = app.CurrentDb().QueryDefs
        else:
            collection = getattr(app.CurrentProject, f"All{category}")
        for obj in collection:
            if obj and obj.Name:
                # 一時クエリやシステムクエリを除外
                if category == "Queries" and (obj.Name.startswith("~") or obj.Name.startswith("MSys")):
                    continue
                ext = OBJECT_EXTENSIONS[category]
                filename = f"{obj.Name}{ext}"
                filepath = os.path.join(export_dir, filename)
                app.SaveAsText(obj_type, obj.Name, filepath)
                exported_files[category].append(filename)
    return exported_files

def import_objects(app, import_dir):
    imported_files = {category: [] for category in OBJECT_TYPES.keys()}
    for filename in os.listdir(import_dir):
        filepath = os.path.join(import_dir, filename)
        obj_name, ext = os.path.splitext(filename)
        for category, ext_type in OBJECT_EXTENSIONS.items():
            if ext == ext_type:
                obj_type = OBJECT_TYPES[category]
                app.LoadFromText(obj_type, obj_name, filepath)
                imported_files[category].append(filename)
                break
    return imported_files

def search_exported_objects(app, pattern):
    results = []
    for category, obj_type in OBJECT_TYPES.items():
        if category == "Queries":
            collection = app.CurrentDb().QueryDefs
        else:
            collection = getattr(app.CurrentProject, f"All{category}")
        for obj in collection:
            if not (obj and obj.Name):
                continue
            temp_dir = tempfile.mkdtemp()
            temp_file_path = os.path.join(temp_dir, "temp_export.txt")
            try:
                app.SaveAsText(obj_type, obj.Name, temp_file_path)
                with open(temp_file_path, 'r', encoding='utf-16-le', errors='ignore') as f: #accessで出力されたァイルはutf-16になる
                    for i, line in enumerate(f, 1):
                        if pattern.lower() in line.lower(): # 大文字小文字 区別なし
                            results.append({
                                "type": category, "name": obj.Name,
                                "line_num": i, "line_content": line.strip(),
                            })
            finally:
                shutil.rmtree(temp_dir)
    return results

def search_object_names(app, pattern):
    results = []
    for category, obj_type in OBJECT_TYPES.items():
        if category == "Queries":
            collection = app.CurrentDb().QueryDefs
        else:
            collection = getattr(app.CurrentProject, f"All{category}")
        for obj in collection:
            if obj and obj.Name and pattern.lower() in obj.Name.lower():
                results.append({
                    "type": f"{category} Name", "name": obj.Name,
                    "line_num": None, "line_content": obj.Name,
                })
    return results

def search_all_access_content(db_path, pattern):
    all_results = []
    # Search in exported objects (VBA, Forms, Reports, Macros, Queries)
    with access_application(db_path) as app:
        all_results.extend(search_object_names(app, pattern))
        all_results.extend(search_exported_objects(app, pattern))

    return all_results

def analyze_usage(app):
    all_objects = {}
    for cat in OBJECT_TYPES:
        if cat == "Queries":
            all_objects[cat] = [obj.Name for obj in app.CurrentDb().QueryDefs if obj and obj.Name and not (obj.Name.startswith("~") or obj.Name.startswith("MSys"))]
        else:
            all_objects[cat] = [obj.Name for obj in getattr(app.CurrentProject, f"All{cat}") if obj and obj.Name and not (obj.Name.startswith("~") or obj.Name.startswith("MSys"))]

    full_source_code = ""
    temp_dir = tempfile.mkdtemp()
    try:
        for category, obj_type in OBJECT_TYPES.items():
            if category == "Queries":
                collection = app.CurrentDb().QueryDefs
            else:
                collection = getattr(app.CurrentProject, f"All{category}")
            for obj in collection:
                if obj and obj.Name and not (obj.Name.startswith("~") or obj.Name.startswith("MSys")):
                    temp_file = os.path.join(temp_dir, "temp.txt")
                    app.SaveAsText(obj_type, obj.Name, temp_file)
                    with open(temp_file, 'r', encoding='utf-16-le', errors='ignore') as f: #accessで出力されたァイルはutf-16になる
                        full_source_code += f.read() + '\n'
    finally:
        shutil.rmtree(temp_dir)
    
    unused = [(cat, name) for cat, names in all_objects.items() for name in names if name not in full_source_code]
    return unused

def get_access_query_names(app):
    query_names = []
    for qdef in app.CurrentDb().QueryDefs:
        # システムクエリや一時クエリを除外
        if not (qdef.Name.startswith("~") or qdef.Name.startswith("MSys")):
            query_names.append(qdef.Name)
    return query_names

def release_prepare(app, test_conn, prod_conn):
    modified = False
    for category in ["Modules", "Forms"]:
        for obj in getattr(app.CurrentProject, f"All{category}"):
            if obj and obj.Name:
                module = app.VBE.ActiveVBProject.VBComponents(obj.Name).CodeModule
                code = module.Lines(1, module.CountOfLines)
                new_code = []
                changed = False
                for line in code.splitlines():
                    new_line = line
                    if test_conn in new_line:
                        new_line = new_line.replace(test_conn, prod_conn)
                        changed = True
                    if "debug.print" in new_line.lower() and not new_line.strip().startswith("'"):
                        new_line = "'" + new_line
                        changed = True
                    new_code.append(new_line)
                if changed:
                    modified = True
                    module.DeleteLines(1, module.CountOfLines)
                    module.AddFromString("\n".join(new_code))
    return modified

def update_linked_table_paths(app, old_path_prefix, new_path_prefix):
    updated_count = 0
    for tdf in app.CurrentDb().TableDefs:
        # リンクテーブルかどうかをチェック (dbAttachedTable属性)
        if tdf.Attributes & win32com.client.constants.dbAttachedTable:
            connect_string = tdf.Connect
            # Connect文字列が古いパスプレフィックスで始まるかチェック
            if connect_string.startswith(old_path_prefix):
                new_connect_string = connect_string.replace(old_path_prefix, new_path_prefix, 1)
                if new_connect_string != connect_string:
                    tdf.Connect = new_connect_string
                    tdf.RefreshLink() # リンクを更新
                    updated_count += 1
    return updated_count
