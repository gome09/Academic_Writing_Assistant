# -*- coding: utf-8 -*-
"""把项目根（thesis_project/）加入 sys.path，使 src / config 可导入。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
