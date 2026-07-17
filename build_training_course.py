# -*- coding: utf-8 -*-
"""入口：转发到 scripts/build_training_course.py"""
from pathlib import Path
import runpy
import sys

sys.argv[0] = str(Path(__file__).resolve().parent / 'scripts' / 'build_training_course.py')
runpy.run_path(sys.argv[0], run_name='__main__')
