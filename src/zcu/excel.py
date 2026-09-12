import logging
import xlwings as xw
import pandas as pd
from contextlib import contextmanager

logger = logging.getLogger(__name__)


# VBA
# def init_vba_project(wb_path: os.PathLike | str | Path) -> None:
#     """
#     Import (overwrite if exists) all .bas files' modules.
#     """
#     wb_path = Path(wb_path).resolve()

#     excel = win32.gencache.EnsureDispatch("Excel.Application")
#     excel.Visible = True
    
#     try:
#         workbook = excel.Workbooks.Open(str(wb_path))
#         vba_project = workbook.VBProject

#         vba_src_dir = Path(__file__).parent / "vba_src"
#         for bas_file in list(vba_src_dir.glob("*.bas")):
#             module_name = bas_file.stem  # Get the module name without extension
#             existing_component = None
#             for component in vba_project.VBComponents:
#                 if component.Name == module_name:
#                     existing_component = component
#                     break
            
#             if existing_component:
#                 # Remove the existing module to allow overwriting
#                 vba_project.VBComponents.Remove(existing_component)
#                 logging.info(f"Removed existing module '{module_name}' to overwrite.")
            
#             # Import the new module
#             vba_project.VBComponents.Import(str(bas_file.resolve()))
#             logging.info(f"Imported {bas_file.name} into VBA project.")

#         if wb_path.suffix.lower() != ".xlsm":
#             xlsm_path = wb_path.with_suffix(".xlsm")
#             workbook.SaveAs(str(xlsm_path), FileFormat=52)
#             logging.info(f"Saved workbook as macro-enabled workbook: {xlsm_path}")
#         else:
#             workbook.Save()
#             logging.info(f"Saved workbook: {wb_path}")

#     except Exception as e:
#         logging.error(f"An error occurred: {e}")


# Application
def get_app_state(wb: xw.Book) -> dict:
    app = wb.app
    return {
        "screen_updating"   : app.screen_updating,
        "display_alerts"    : app.display_alerts,
        "enable_events"     : app.enable_events,
        "calculation"       : app.calculation,
        "status_bar"        : app.status_bar,
    }

def set_app_quiet(wb: xw.Book) -> None:
    app = wb.app

    app.screen_updating     = False
    app.display_alerts      = False
    app.enable_events       = False
    app.calculation         = 'manual'
    app.status_bar          = False

def set_app_state(wb: xw.Book, state: dict) -> None:
    app = wb.app
    for key, value in state.items():
        setattr(app, key, value)

@contextmanager
def excel_write_transaction(wb: xw.Book):
    state = get_app_state(wb)
    try:
        set_app_quiet(wb)
        yield
    finally:
        set_app_state(wb, state)


# Book
def find_open_workbook(wb_name: str):
    '''Return an xlwings Book object if found, else None.'''
    for wb in xw.books:
        if wb.name == wb_name:
            return wb

    raise FileNotFoundError(f"Workbook '{wb_name}' not found among open workbooks.")


# Sheet
def find_sheets(wb: xw.Book, sheet_names: list[str]):
    '''Return a list of xlwings Sheet objects if found, else None.'''
    found_sheets = {}
    for sht in wb.sheets:
        if sht.name in sheet_names:
            found_sheets[sht.name] = sht
    return found_sheets

def find_sheet(wb: xw.Book, sheet_name: str):
    '''Return xlwings Sheet object if found, else None.'''
    for sht in wb.sheets:
        if sht.name == sheet_name:
            return sht
    return None

def read_sheet(
        wb: xw.Book,
        sheet_name: str,
        *,
        start_cell: str = 'A1',
        header: bool = True,
) -> pd.DataFrame:
    '''Read the data from a sheet and return it as a pandas DataFrame.'''
    sheet = find_sheet(wb, sheet_name)
    if sheet:
        df = sheet.range(start_cell).options(pd.DataFrame, header=header).value
        return df
    else:
        logger.warning(f"Sheet '{sheet_name}' not found.")
        return pd.DataFrame()


# Table
def find_tables(wb: xw.Book, table_names: list[str]):
    '''Return a list of xlwings Table objects if found, else None.'''
    found_tables = {}
    for sht in wb.sheets:
        for table in sht.tables:
            if table.name in table_names:
                found_tables[table.name] = table
    return found_tables

def find_table(wb: xw.Book, table_name: str):
    '''Return xlwings Table object if found, else None.'''
    for sht in wb.sheets:
        if table_name in [t.name for t in sht.tables]:
            return sht.tables[table_name]

def read_table(wb: xw.Book, table_name: str):
    '''Read the data from a table and return it as a pandas DataFrame.'''
    table = find_table(wb, table_name)
    if table:
        return table.range.options(pd.DataFrame, index=False).value
    else:
        logger.warning(f"Table '{table_name}' not found.")
        return pd.DataFrame()

def resize_table(
        table: xw.main.Table,
        rows: int | None = None,
        cols: int | None = None,
        *,
        clear_outside: bool = True,
) -> None:
    """
    Resize an Excel Table and optionally clear data that falls outside the new range.

    Args:
        table: xlwings Table object to resize.
        rows: New number of data rows (excluding header). If None, keeps the current number of rows.
        cols: New number of columns. If None, keeps the current number of columns.
        clear_outside: If True, clears contents of cells that fall outside the new table range.
    """
    old_range = table.range
    old_rows = max(old_range.rows.count - 1, 0)
    old_cols = old_range.columns.count

    rows = old_rows if rows is None else max(rows, 0)
    cols = old_cols if cols is None else max(cols, 1)

    new_range = old_range.resize(rows + 1, cols)
    table.resize(new_range)

    if not clear_outside:
        return

    sheet = table.parent

    if rows < old_rows:
        below = sheet.range(
            (new_range.last_cell.row + 1, old_range.column),
            (old_range.last_cell.row, old_range.column + old_cols - 1)
        )
        below.clear_contents()

    if cols < old_cols:
        right = sheet.range(
            (old_range.row, new_range.last_cell.column + 1),
            (old_range.last_cell.row, old_range.last_cell.column)
        )
        right.clear_contents()

def update_table(
        *,
        df: pd.DataFrame,
        table: xw.main.Table,
        index: bool = False,
        clear_outside: bool = True,
) -> None:
    '''
    Update an Excel Table with a DataFrame. Wrapper around `xlwings.Table.update()` with additional features.

    Args:
        df: DataFrame to write.
        table: xlwings Table object.
        clear_outside: Clear cells removed from the original table range after resize.
    '''
    if df.columns.duplicated().any():
        duplicated = df.columns[df.columns.duplicated()].tolist()
        raise ValueError(f"Duplicated DataFrame columns: {duplicated}")

    with excel_write_transaction(table.parent.book):
        resize_table(
            table,
            rows=len(df),
            cols=len(df.columns),
            clear_outside=clear_outside,
        )
        if not df.empty:
            table.update(df, index=index)
        else:
            table.header_row_range.value = [list(df.columns)]
        return

def update_table_with_formula(
        *,
        df: pd.DataFrame,
        table: xw.main.Table,
        formula_map: dict[str, str],
        output_col: list[str],
        clear_outside: bool = True,
) -> None:
    '''
    Update an Excel Table with DataFrame values and Excel formulas.

    Args:
        df: DataFrame containing the non-formula columns.
        table: xlwings Table object.
        formula_map: Mapping from formula column name to Excel formula.
            Structured references such as `=[@Qty]*[@Price]` are supported.
        output_col: Complete output schema and column order of the Excel Table. 
            Each column must exist in either `df` or `formula_map`.
        clear_outside: Clear cells removed from the original table range after resize.
    '''
    if df.columns.duplicated().any():
        duplicated = df.columns[df.columns.duplicated()].tolist()
        raise ValueError(f"Duplicated DataFrame columns: {duplicated}")

    if len(output_col) != len(set(output_col)):
        raise ValueError("Duplicated columns in output_col.")

    missing_df = [col for col in df.columns if col not in output_col]

    if missing_df:
        raise ValueError(f"DataFrame columns missing from output_col: {missing_df}")

    missing_formula = [col for col in formula_map if col not in output_col]

    if missing_formula:
        raise ValueError(f"Formula columns missing from output_col: {missing_formula}")

    unknown = [col for col in output_col if col not in df.columns and col not in formula_map]

    if unknown:
        raise ValueError(f"Columns in output_col are neither DataFrame nor formula columns: {unknown}")

    with excel_write_transaction(table.parent.book):
        resize_table(
            table,
            rows=len(df),
            cols=len(output_col),
            clear_outside=clear_outside,
        )

        table.header_row_range.value = [output_col]

        if df.empty:
            return

        data_body = table.data_body_range

        for i, col in enumerate(output_col):
            col_range = data_body[:, i]

            if col in formula_map:
                col_range.formula = formula_map[col]
            else:
                col_range.value = df[col].to_numpy().reshape(-1, 1)


# Test
def main():
    ...


if __name__ == "__main__":
    main()