import os
import pytest
from src.command.diff import diff_text_files, diff_exported_objects

# テスト用のダミーファイルを作成するヘルパー関数
@pytest.fixture
def create_dummy_files(tmp_path):
    file1 = tmp_path / "file1.txt"
    file2 = tmp_path / "file2.txt"
    file3 = tmp_path / "file3.txt" # 異なる内容
    file4 = tmp_path / "file4.txt" # 存在しないファイル用

    file1.write_text("Line 1\nLine 2\nLine 3\n")
    file2.write_text("Line 1\nLine 2\nLine 3\n") # file1と同じ内容
    file3.write_text("Line A\nLine B\nLine C\n")

    return str(file1), str(file2), str(file3), str(file4)

def test_diff_text_files_no_diff(create_dummy_files):
    """内容が同じファイル間の差分がないことをテスト"""
    file1, file2, _, _ = create_dummy_files
    diff_result = diff_text_files(file1, file2)
    assert not diff_result # 差分がない場合は空リストが返されることを期待

def test_diff_text_files_with_diff(create_dummy_files):
    """内容が異なるファイル間の差分があることをテスト"""
    file1, _, file3, _ = create_dummy_files
    diff_result = diff_text_files(file1, file3)
    assert diff_result # 差分があることを確認
    assert "--- " + os.path.basename(file1) in diff_result[0]
    assert "+++ " + os.path.basename(file3) in diff_result[1]
    assert "-Line 1" in diff_result[3]
    assert "+Line A" in diff_result[3] # 差分の具体的な内容を一部確認

def test_diff_text_files_missing_file(create_dummy_files):
    """ファイルが見つからない場合のエラーハンドリングをテスト"""
    file1, _, _, file4 = create_dummy_files
    diff_result = diff_text_files(file1, file4)
    assert "Error comparing files" in diff_result[0] # エラーメッセージが含まれることを確認

# diff_exported_objects のテスト
@pytest.fixture
def setup_exported_dirs(tmp_path):
    dir1 = tmp_path / "dir1"
    dir2 = tmp_path / "dir2"
    dir1.mkdir()
    dir2.mkdir()

    # dir1 のみにあるファイル
    (dir1 / "file_only_in_dir1.txt").write_text("content A")

    # dir2 のみにあるファイル
    (dir2 / "file_only_in_dir2.txt").write_text("content B")

    # 両方にあり、内容が同じファイル
    (dir1 / "same_content.txt").write_text("same content")
    (dir2 / "same_content.txt").write_text("same content")

    # 両方にあり、内容が異なるファイル
    (dir1 / "diff_content.txt").write_text("original content")
    (dir2 / "diff_content.txt").write_text("modified content")

    return str(dir1), str(dir2)

def test_diff_exported_objects(setup_exported_dirs):
    dir1, dir2 = setup_exported_dirs
    diffs = diff_exported_objects(dir1, dir2)

    # dir1 のみにあるファイルのテスト
    assert "file_only_in_dir1.txt" in diffs
    assert "Object only exists in the first file." in diffs["file_only_in_dir1.txt"][3]

    # dir2 のみにあるファイルのテスト
    assert "file_only_in_dir2.txt" in diffs
    assert "Object only exists in the second file." in diffs["file_only_in_dir2.txt"][3]

    # 内容が同じファイルのテスト
    assert "same_content.txt" not in diffs # 差分がないので含まれない

    # 内容が異なるファイルのテスト
    assert "diff_content.txt" in diffs
    assert "-original content" in diffs["diff_content.txt"][3]
    assert "+modified content" in diffs["diff_content.txt"][4]