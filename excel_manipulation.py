import os
import re
import openpyxl
from openpyxl import load_workbook


class ExcelApi:
    def __init__(self, filename: str = None) -> None:

        try:

            if filename is None:
                raise ValueError("File name must be a non-empty string")

            if type(filename) is not str:
                raise TypeError("File name must be string")

            if not filename.endswith(".xlsx"):
                filename += ".xlsx"

            self.filename = filename

        except ValueError as e:
            print(f"Value error: {e}")

        except TypeError as e:
            print(f"Type error: {e}")

        except Exception as e:
            print(f"An unexcpected error occured: {e}")

    def excel_create(self, data: list = None) -> None:

        try:

            if data is None:
                raise ValueError("Data must be a non-empty list")

            if type(data) is not list:
                raise TypeError("Data must be list")

            wb = openpyxl.Workbook()

            sheet = wb.active

            for row_index, row_data in enumerate(data, start=1):
                for col_index, cell_val in enumerate(row_data, start=1):
                    sheet.cell(row=row_index, column=col_index, value=cell_val)

            wb.save(self.filename)

        except ValueError as e:
            print(f"Value error: {e}")

        except TypeError as e:
            print(f"Type error: {e}")

        except Exception as e:
            print(f"An unexcpected error occured: {e}")

        else:
            print(f"Excel file '{self.filename}' created successfully.")

    def excel_delete(self) -> None:

        try:

            if os.path.exists(self.filename):
                os.remove(self.filename)
                print(f"{self.filename} has been successfully deleted.")

            else:
                print(f"{self.filename} does not exist.")

        except OSError as e:
            print(f"Error: {e.strerror}")

        except Exception as e:
            print(f"An unexcepted error occured: {e}")

    def excel_updete(self,
                     sheetname: str = None,
                     cell: str = None,
                     value=0) -> None:

        def valide_cell(cell):
            pattern = re.compile(r"^[A-Z]+[1-9]\d*$")
            if not pattern.match(cell):
                raise ValueError(f"Invalid cell reference: '{cell}'")

        try:
            valide_cell(cell)

            if sheetname is None:
                raise ValueError("Sheet name parameter is None/Null")

            if cell is None:
                raise ValueError("Cell parameter is None/Null")

            wb = load_workbook(self.filename)

            if sheetname in wb.sheetnames:
                sheet = wb[sheetname]
            else:
                print(f"{sheetname} does not exist")

            sheet[cell] = value

            wb.save(self.filename)
            print(f"Updeted cell {cell} in {sheetname} with value '{value}'.")

        except FileNotFoundError:
            print(f"Error: The file '{self.filename}' does not exist.")

        except ValueError as e:
            print(f"Value error: {e}")

        except TypeError as e:
            print(f"Type error: Invalid parameter type -{e}")

        except Exception as e:
            print(f"An unexcepted error occured: {e}")

    def excel_open(self):

        try:
            if not os.path.exists(self.filename):
                print(f"File {self.filename} does not exist. Creating...")
                data = []
                self.excel_create(data)

            wb = load_workbook(self.filename)

            return wb
        except Exception as e:
            print(f"An unexpected error occured: {e}")

    def sheet_create(self, sheettitle: str = None) -> None:

        try:

            if sheettitle is None:
                raise ValueError("Sheet title parameter is None/Null")

            wb = load_workbook(self.filename)

            wb.create_sheet(title=sheettitle)

            wb.save(self.filename)

        except FileNotFoundError:
            print(f"Error: The file '{self.filename}' does not exist.")

        except TypeError as e:
            print(f"Error: Invalid parameter type -{e}")

        except Exception as e:
            print(f"An unexcepted error occured: {e}")

        else:
            print("New sheet is created successfully.")

    def sheet_name_change(self,
                          sheetname: str = None,
                          newname: str = None) -> None:

        try:

            if sheetname is None:
                raise ValueError("Sheet name parameter is None/Null")

            if newname is None:
                raise ValueError("Sheet name parameter is None/Null")

            wb = load_workbook(self.filename)

            if sheetname in wb.sheetnames:
                sheet = wb[sheetname]

                sheet.title = newname

                wb.save(self.filename)

                print("Sheet name changed")
            else:
                print(f"{sheetname}does not exist")

        except FileNotFoundError:
            print(f"Error: The file '{self.filename}' does not exist.")

        except TypeError as e:
            print(f"Error: Invalid parameter type -{e}")

        except Exception as e:
            print(f"An unexcepted error occured: {e}")

    def sheet_delete(self, sheetname: str = None) -> None:

        try:

            if sheetname is None:
                raise ValueError("Sheet name parameter is None/Null")

            wb = load_workbook(self.filename)

            if sheetname in wb.sheetnames:

                wb.remove(wb[sheetname])

                wb.save(self.filename)
                print(f"The sheet '{sheetname}' has been deleted.")
            else:
                print(f"The sheet '{sheetname}' does not exist.")

        except FileNotFoundError:
            print(f"Error: The file '{self.filename}' does not exist.")

        except TypeError as e:
            print(f"Error: Invalid parameter type -{e}")

        except Exception as e:
            print(f"An unexcepted error occured: {e}")
