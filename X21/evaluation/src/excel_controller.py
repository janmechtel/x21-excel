"""
Excel automation via pywin32 for headed evaluation.
"""

import os
import time
from typing import Optional

import win32com.client as win32


class ExcelController:
    """Controls Excel via COM automation for evaluation."""

    def __init__(self):
        self.excel = None
        self.workbook = None

    def connect(self) -> bool:
        """Connect to running Excel instance or start a new one."""
        try:
            self.excel = win32.GetActiveObject("Excel.Application")
            print("Connected to existing Excel instance")
            return True
        except Exception:
            try:
                self.excel = win32.Dispatch("Excel.Application")
                self.excel.Visible = True
                print("Started new Excel instance")
                return True
            except Exception as e:
                print(f"Failed to connect to Excel: {e}")
                return False

    def open_workbook(self, filepath: str) -> bool:
        """Open a workbook from the given path."""
        if not self.excel:
            print("Excel not connected")
            return False

        try:
            abs_path = os.path.abspath(filepath)
            self.workbook = self.excel.Workbooks.Open(abs_path)
            # Small delay to let Excel settle
            time.sleep(0.5)
            return True
        except Exception as e:
            print(f"Failed to open workbook: {e}")
            return False

    def get_workbook_name(self) -> Optional[str]:
        """Get the name of the active workbook."""
        if self.workbook:
            return self.workbook.Name
        return None

    def get_workbook_path(self) -> Optional[str]:
        """Get the full path of the active workbook."""
        if self.workbook:
            return self.workbook.FullName
        return None

    def save_workbook(self, filepath: Optional[str] = None) -> bool:
        """Save the workbook, optionally to a new path."""
        if not self.workbook:
            print("No workbook open")
            return False

        try:
            if filepath:
                abs_path = os.path.abspath(filepath)
                # Disable alerts to suppress "file exists" dialog
                self.excel.DisplayAlerts = False
                self.workbook.SaveAs(abs_path)
                self.excel.DisplayAlerts = True
            else:
                self.workbook.Save()
            return True
        except Exception as e:
            self.excel.DisplayAlerts = True  # Re-enable alerts on error
            print(f"Failed to save workbook: {e}")
            return False

    def close_workbook(self, save: bool = False) -> bool:
        """Close the current workbook."""
        if not self.workbook:
            return True

        try:
            self.workbook.Close(SaveChanges=save)
            self.workbook = None
            return True
        except Exception as e:
            print(f"Failed to close workbook: {e}")
            return False

    def quit(self):
        """Quit Excel application."""
        if self.excel:
            try:
                self.excel.Quit()
            except Exception:
                pass
            self.excel = None
