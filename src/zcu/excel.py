"""
This module houses interactions with Excel files via `xlwings`.
"""

import os

import xlwings as xw
import win32com.client as win32
import pandas as pd
import numpy as np

from pathlib import Path


# Configure logging
from .misc import configure_logging
configure_logging()



def init_vba_project(wb_path: os.PathLike | str | Path) -> None:
    """
    Import (overwrite if exists) all .bas files' modules.
    """
    wb_path = Path(wb_path).resolve()

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = True
    
    try:
        workbook = excel.Workbooks.Open(str(wb_path))
        vba_project = workbook.VBProject

        vba_src_dir = Path(__file__).parent / "vba_src"
        for bas_file in list(vba_src_dir.glob("*.bas")):
            module_name = bas_file.stem  # Get the module name without extension
            existing_component = None
            for component in vba_project.VBComponents:
                if component.Name == module_name:
                    existing_component = component
                    break
            
            if existing_component:
                # Remove the existing module to allow overwriting
                vba_project.VBComponents.Remove(existing_component)
                logging.info(f"Removed existing module '{module_name}' to overwrite.")
            
            # Import the new module
            vba_project.VBComponents.Import(str(bas_file.resolve()))
            logging.info(f"Imported {bas_file.name} into VBA project.")

        if wb_path.suffix.lower() != ".xlsm":
            xlsm_path = wb_path.with_suffix(".xlsm")
            workbook.SaveAs(str(xlsm_path), FileFormat=52)
            logging.info(f"Saved workbook as macro-enabled workbook: {xlsm_path}")
        else:
            workbook.Save()
            logging.info(f"Saved workbook: {wb_path}")

    except Exception as e:
        logging.error(f"An error occurred: {e}")


# Test
def main():
    ...


if __name__ == "__main__":
    main()