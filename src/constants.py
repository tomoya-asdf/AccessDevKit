# -*- coding: utf-8 -*-

import os
import sys

current_constants_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.abspath(os.path.join(current_constants_dir, os.pardir))

BASE_APP_DIR = os.path.dirname(sys.argv[0])

# Report Output Paths (relative to BASE_APP_DIR)
DIFF_REPORT_PATH = os.path.join(BASE_APP_DIR, "output", "reports", "access_diff_report.html")
UNUSED_OBJECTS_REPORT_PATH = os.path.join(BASE_APP_DIR, "output", "reports", "unused_objects_report.html")
BENCHMARK_REPORT_PATH = os.path.join(BASE_APP_DIR, "output", "reports", "benchmark_report.html")

# Log Output Paths (relative to BASE_APP_DIR)
LOG_DIR = os.path.join(BASE_APP_DIR, "logs")
LOG_FILE_NAME_ALL = "{datetime}.log"

# Determine the base path for resources (templates, etc.)
if getattr(sys, 'frozen', False):
    if hasattr(sys, '_MEIPASS'):
        RESOURCE_BASE_PATH = sys._MEIPASS
    else:
        RESOURCE_BASE_PATH = os.path.dirname(sys.argv[0])
else:
    RESOURCE_BASE_PATH = project_root

# Template File Names (relative to src/templates)
DIFF_REPORT_TEMPLATE = "report_template.html"
UNUSED_OBJECTS_REPORT_TEMPLATE = "unused_objects_report_template.html"
BENCHMARK_REPORT_TEMPLATE = "benchmark_report_template.html"

# Absolute path to the templates directory
TEMPLATES_DIR = os.path.join(RESOURCE_BASE_PATH, 'src', 'templates')
