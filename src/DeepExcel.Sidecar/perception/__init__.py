"""工作表结构理解：宿主（C# / WPS JS）只负责一次读出有界快照，算法都在这里，两个宿主共用。

入口是 inspect(snapshot, layers)，见 inspect_sheet.py；各层算法分别在 blocks / headers / formulas / anomalies。
"""

from .inspect_sheet import LAYERS, inspect, normalize_layers

__all__ = ["inspect", "normalize_layers", "LAYERS"]
