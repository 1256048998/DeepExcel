import sys
import os

# 让 tests/ 能 import sidecar 模块
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
