import io
from typing import Union
import openpyxl
from openpyxl.utils import get_column_letter


class ExcelParser:
    """Parses Excel files (.xlsx) and extracts structured data from all sheets."""

    def parse(self, file_path_or_bytes: Union[str, bytes, io.BytesIO]) -> dict:
        """
        Parse an Excel file and extract all sheet data.

        Args:
            file_path_or_bytes: Path to file, raw bytes, or BytesIO object.

        Returns:
            {
                "sheets": {
                    sheet_name: {
                        "headers": [str, ...],
                        "rows": [[cell_value, ...], ...],
                        "numeric_cells": [
                            {"ref": "B12", "valeur": 1234.5, "sheet": "CA",
                             "row": 12, "col": 2},
                            ...
                        ]
                    }
                },
                "source": filename_or_bytes_label,
                "total_sheets": int
            }
        """
        if isinstance(file_path_or_bytes, bytes):
            source = "uploaded_bytes"
            file_obj = io.BytesIO(file_path_or_bytes)
        elif isinstance(file_path_or_bytes, io.BytesIO):
            source = getattr(file_path_or_bytes, "name", "uploaded_file")
            file_obj = file_path_or_bytes
        else:
            source = str(file_path_or_bytes)
            file_obj = file_path_or_bytes

        wb = openpyxl.load_workbook(file_obj, read_only=True, data_only=True)

        result_sheets = {}

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            headers = []
            rows = []
            numeric_cells = []

            all_rows = list(ws.iter_rows(values_only=False))

            # Find headers: first non-empty row
            header_row_index = None
            for idx, row in enumerate(all_rows):
                row_values = [cell.value for cell in row]
                if any(v is not None and str(v).strip() != "" for v in row_values):
                    headers = [
                        str(cell.value).strip() if cell.value is not None else ""
                        for cell in row
                    ]
                    header_row_index = idx
                    break

            # Extract data rows (all rows after header)
            if header_row_index is not None:
                for row in all_rows[header_row_index + 1:]:
                    row_values = [cell.value for cell in row]
                    rows.append(row_values)

                    # Extract numeric cells
                    for cell in row:
                        if isinstance(cell.value, (int, float)) and not isinstance(
                            cell.value, bool
                        ):
                            col_letter = get_column_letter(cell.column)
                            ref = f"{col_letter}{cell.row}"
                            numeric_cells.append(
                                {
                                    "ref": ref,
                                    "valeur": float(cell.value),
                                    "sheet": sheet_name,
                                    "row": cell.row,
                                    "col": cell.column,
                                }
                            )

                # Also check header row for numeric values
                for cell in all_rows[header_row_index]:
                    if isinstance(cell.value, (int, float)) and not isinstance(
                        cell.value, bool
                    ):
                        col_letter = get_column_letter(cell.column)
                        ref = f"{col_letter}{cell.row}"
                        numeric_cells.append(
                            {
                                "ref": ref,
                                "valeur": float(cell.value),
                                "sheet": sheet_name,
                                "row": cell.row,
                                "col": cell.column,
                            }
                        )

            result_sheets[sheet_name] = {
                "headers": headers,
                "rows": rows,
                "numeric_cells": numeric_cells,
            }

        wb.close()

        return {
            "sheets": result_sheets,
            "source": source,
            "total_sheets": len(result_sheets),
        }
