# -*- coding: utf-8 -*>
import os
import sys
import re
import json
from src.constants import (
    DIFF_REPORT_TEMPLATE,
    UNUSED_OBJECTS_REPORT_TEMPLATE,
    BENCHMARK_REPORT_TEMPLATE,
    TEMPLATES_DIR
)

class ReportGenerator:
    def __init__(self):
        pass

    def _get_template_path(self, template_name):
        return os.path.join(TEMPLATES_DIR, template_name)

    def _write_html_report(self, output_path, html_content):
        output_dir = os.path.dirname(output_path)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

    def create_diff_report(self, table_diffs, vba_diffs, output_path, report_datetime, file1_path, file2_path):
        html_template = ""
        template_path = self._get_template_path(DIFF_REPORT_TEMPLATE)
        with open(template_path, 'r', encoding='utf-8') as f:
            html_template = f.read()

        # サマリー情報の計算
        total_table_changes = 0
        for table, (only1, only2) in table_diffs.items():
            total_table_changes += len(only1) + len(only2)
        
        total_vba_changes = len(vba_diffs)

        # テーブル差分のHTML生成
        table_diffs_html = ""
        if not table_diffs:
            table_diffs_html = "<tr><td colspan=\"3\" class=\"no-changes\">テーブルの変更は見つかりませんでした。</td></tr>"
        else:
            for table, (only1, only2) in sorted(table_diffs.items()):
                if only1:
                    for row in list(only1)[:100]:
                        table_diffs_html += f"<tr class=\"removed\"><td>REMOVED</td><td>{table}</td><td>{' '.join(map(str, row))}</td></tr>\n"
                if only2:
                    for row in list(only2)[:100]:
                        table_diffs_html += f"<tr class=\"added\"><td>ADDED</td><td>{table}</td><td>{' '.join(map(str, row))}</td></tr>\n"

        # VBA差分のHTML生成
        vba_diffs_html = ""
        if not vba_diffs:
            vba_diffs_html = "<tr><td colspan=\"2\" class=\"no-changes\">VBA/フォームの変更は見つかりませんでした。</td></tr>"
        else:
            for fname, diff_lines in sorted(vba_diffs.items()):
                formatted_diff_lines = []
                line_num1 = 0
                line_num2 = 0
                for line in diff_lines:
                    cleaned_line = ''.join(char for char in line if char.isprintable() or char == ' ' or char == '\t')
                    
                    # unified_diffのヘッダー行を解析
                    match = re.match(r'@@ -(\d+)(?:,\d+)? \+(\d+)(?:,\d+)? @@', cleaned_line)
                    if match:
                        line_num1 = int(match.group(1))
                        line_num2 = int(match.group(2))
                        formatted_diff_lines.append(f'<span class=\"diff-header\">{cleaned_line}</span>')
                        continue

                    line_content = cleaned_line[1:] # 記号を除いた内容
                    
                    if cleaned_line.startswith('+'):
                        formatted_diff_lines.append(f'<span class=\"diff-line-added\"><span class=\"line-num-old\"></span><span class=\"line-num-new\">{line_num2}</span> {line_content}</span>')
                        line_num2 += 1
                    elif cleaned_line.startswith('-'):
                        formatted_diff_lines.append(f'<span class=\"diff-line-removed\"><span class=\"line-num-old\">{line_num1}</span><span class=\"line-num-new\"></span> {line_content}</span>')
                        line_num1 += 1
                    else: # context line
                        formatted_diff_lines.append(f'<span class=\"diff-line-unchanged\"><span class=\"line-num-old\">{line_num1}</span><span class=\"line-num-new\">{line_num2}</span> {line_content}</span>')
                        line_num1 += 1
                        line_num2 += 1
                
                vba_diffs_html += f"<tr><td>{fname}</td><td><pre class=\"diff-content\">{'<br>'.join(formatted_diff_lines)}</pre></td></tr>\n"

        # テンプレートにデータを埋め込み
        html_report = html_template.replace('{{table_diffs_html}}', table_diffs_html)
        html_report = html_report.replace('{{vba_diffs_html}}', vba_diffs_html)
        html_report = html_report.replace('{{report_datetime}}', report_datetime)
        html_report = html_report.replace('{{file1_path}}', file1_path)
        html_report = html_report.replace('{{file2_path}}', file2_path)
        html_report = html_report.replace('{{total_table_changes}}', str(total_table_changes))
        html_report = html_report.replace('{{total_vba_changes}}', str(total_vba_changes))

        self._write_html_report(output_path, html_report)

    def create_unused_objects_report(self, unused_objects, output_path, report_datetime, file_path):
        html_template = ""
        template_path = self._get_template_path(UNUSED_OBJECTS_REPORT_TEMPLATE)
        with open(template_path, 'r', encoding='utf-8') as f:
            html_template = f.read()

        unused_objects_html = ""
        if not unused_objects:
            unused_objects_html = "<tr><td colspan=\"2\" class=\"no-changes\">未使用のオブジェクトは見つかりませんでした。</td></tr>"
        else:
            for obj_type, obj_name in sorted(unused_objects):
                unused_objects_html += f"<tr><td>{obj_type}</td><td>{obj_name}</td></tr>\n"

        html_report = html_template.replace('{{unused_objects_html}}', unused_objects_html)
        html_report = html_report.replace('{{report_datetime}}', report_datetime)
        html_report = html_report.replace('{{file_path}}', file_path)
        html_report = html_report.replace('{{total_unused_objects}}', str(len(unused_objects)))

        self._write_html_report(output_path, html_report)

    def create_benchmark_report(self, benchmark_results, output_path, report_datetime, file_path):
        html_template = ""
        template_path = self._get_template_path(BENCHMARK_REPORT_TEMPLATE)
        with open(template_path, 'r', encoding='utf-8') as f:
            html_template = f.read()

        benchmark_results_html = ""
        if not benchmark_results:
            benchmark_results_html = "<tr><td colspan=\"3\" class=\"no-changes\">ベンチマーク結果は見つかりませんでした。</td></tr>"
        else:
            for query_name, avg_time, total_time in benchmark_results:
                benchmark_results_html += f"<tr><td>{query_name}</td><td>{avg_time:.4f}</td><td>{total_time:.4f}</td></tr>\n"

        # グラフデータ用にJSON形式でデータを渡す
        chart_labels = json.dumps([res[0] for res in benchmark_results])
        chart_data = json.dumps([res[1] for res in benchmark_results])

        html_report = html_template.replace('{{benchmark_results_html}}', benchmark_results_html)
        html_report = html_report.replace('{{report_datetime}}', report_datetime)
        html_report = html_report.replace('{{file_path}}', file_path)
        html_report = html_report.replace('{{chart_labels}}', chart_labels)
        html_report = html_report.replace('{{chart_data}}', chart_data)

        self._write_html_report(output_path, html_report)