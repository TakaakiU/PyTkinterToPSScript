import pandas
import os
import win32com.client

class ctrlExcel():
    def testexcel():
        result = 0
        
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            version = float(excel.Version)
            excel.Quit()

            # Error if Excel version is earlier than 2019
            if version < 16.0:
                result = -5001
        except Exception as err:
            result = -5002

        return result

    def readxlsx(input_path):
        """
        Read XLSX file
        """
        # Check if file exists
        if not os.path.isfile(input_path):
            return None
        # Read XLSX file
        try:
            xlsx_data = None
            xlsx_data = pandas.read_excel(input_path,
                                          header=None,
                                          sheet_name=0,
                                          skiprows=1)
        except OSError:
            print('Exception error: Failed to read XLSX')
        return xlsx_data
