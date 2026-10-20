"""Structure-preserving data import for .xlsm templates."""

from .api import ControlInfo, ImportReport, SheetEditor, Template
from .spec import apply, extract
from .verify import inspect, verify
from .workbook import Workbook, XlsmError

__all__ = ["ControlInfo", "ImportReport", "SheetEditor", "Template", "Workbook", "XlsmError",
           "apply", "extract", "inspect", "verify"]
