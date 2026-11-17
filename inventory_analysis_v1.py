#!/usr/bin/env python3
"""تحليل لقطة أعمار المخزون بلغة Python 3.12.

التثبيت: python -m pip install "pandas>=2.2,<3" "XlsxWriter>=3.2,<4"
التشغيل: python inventory_analysis.py inventory_20261114.csv -o report.xlsx
المساعدة: python inventory_analysis.py --help

تُقرأ شرائح الأعمار من اللقطة. تاريخ اللقطة يؤثر في حساب أيام الحركة.
شغّل الدالة analyze_inventory(InventoryConfig(...)) للاستعمال من كود آخر.
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from tempfile import TemporaryDirectory
from time import perf_counter

import numpy as np
import pandas as pd
from xlsxwriter import Workbook
from xlsxwriter.exceptions import XlsxWriterException
from xlsxwriter.worksheet import Worksheet

# Source restored from commit c6e9500215895e6c75d79d83713545314a796c98.
# The full original content is 32,110 bytes; use commit blob instead.