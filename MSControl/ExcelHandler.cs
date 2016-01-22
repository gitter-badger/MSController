using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace MSController
{
    /// <summary>
    /// Handles Microsoft Excel, can read and write to cells and also check if a spreadsheet is open.
    /// </summary>
    public class ExcelHandler
    {
        Excel.Application excelApp;
        Excel.Workbooks workbooks;
        Excel.Workbook workbook;
        Excel.Sheets worksheets;
        Excel.Worksheet worksheet;
        Excel.Range range;
        object missing = System.Reflection.Missing.Value;

        /// <summary>
        /// Opens an excel spreadsheet for processing.
        /// </summary>
        /// <param name="filePath">The filepath string of the spreadsheet to be opened.</param>
        /// <param name="sheet">The worksheet to open.</param>
        public void open(string filePath, string sheet = "defualt")
        {
            excelApp = new Excel.Application();

            if (!File.Exists(filePath))
                throw new FileNotFoundException();

            excelApp.Visible = false;
            workbooks = excelApp.Workbooks;
            workbook = workbooks.Open(filePath, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
            worksheets = workbook.Worksheets;

            if (sheet.Equals("defualt"))
            {
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            }
            else
            {
                try
                {
                    worksheet = (Excel.Worksheet)worksheets.get_Item(sheet);
                }
                catch (Exception e)
                {
                    worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                }
            }            
        }

        /// <summary>
        /// Closes the excel spreadsheet.
        /// </summary>
        /// <param name="save">Boolean value of whether or not to save the file.</param>
        public void close(Boolean save = false)
        {
            if (save == true)
                workbook.Save();

            workbook.Close(0);
            excelApp.Quit();

            // Release the COM objects
            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheets);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            // The internet said do this and it works so it's here
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Creates an excel spreadsheet.
        /// </summary>
        /// <param name="filePath">The filepath string of the spreadsheet to be created.</param>
        public void create(string filePath)
        {
            // TODO: Create create() method
        }

        /// <summary>
        /// Checks whether a spreadsheet is open or not.
        /// </summary>
        /// <param name="filePath">String value of the column of the cell.</param>
        /// <returns>
        /// True if the file is open, false if not.
        /// </returns>
        public Boolean isOpen(string filePath)
        {
            // TODO: Create isOpen() method
            return false;
        }

        /// <summary>
        /// Gets the value from a specified cell in the open spreadsheet.
        /// </summary>
        /// <param name="column">String value of the column of the cell.</param>
        /// <param name="row">Int value of the row of the cell.</param>
        /// <returns>
        /// The value from the specified cell.
        /// </returns>
        public string getCell(string column, int row)
        {
            range = worksheet.Range[column + row];
            string cellValue = range.Value.ToString();

            return cellValue;
        }

        /// <summary>
        /// Gets the value from the last cell in a specified column in the open spreadsheet.
        /// </summary>
        /// <param name="column">String value of the column of the cell.</param>
        /// <returns>
        /// The value from the last cell in the specified column.
        /// </returns>
        public string getLastCellInColumn(string column)
        {
            int counter = 1;
            range = worksheet.Range[column + counter];
            string lastCell = "";

            while (range.Value != null)
            {
                lastCell = range.Value.ToString();
                counter++;
                range = worksheet.Range[column + counter];
            }

            return lastCell;
        }

        /// <summary>
        /// Gets the value from the last cell in a specified row in the open spreadsheet.
        /// </summary>
        /// <param name="row">Int value of the row of the cell.</param>
        /// <returns>
        /// The value from the last cell in the specified row.
        /// </returns>
        public string getLastCellInRow(int row)
        {
            // TODO: getLastColumnCell() method
            List<string> alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".Select(x => x.ToString()).ToList();

            for (int i = 0; i < 26; i++)
                for (int j = 0; j < 26; j++)
                    alphabet.Add(alphabet[i] + alphabet[j]);

            return "";
        }

        /// <summary>
        /// Writes a value to a specified cell in the open spreadsheet.
        /// </summary>
        /// <param name="column">String value of the column of the cell.</param>
        /// <param name="row">Int value of the row of the cell.</param>
        /// <param name="data">The value to write to the cell.</param>
        /// <param name="numberFormat">Whether the data should be formatted as a number (Prevents scientific notation being used).</param>
        public void writeCell(string column, int row, string data, Boolean numberFormat = false)
        {
            range = worksheet.Range[column + row.ToString()];
            range.Value = data;
            if (numberFormat)
                range.NumberFormat = "#";
        }

        /// <summary>
        /// Writes a value to the last cell in a specified column in the open spreadsheet.
        /// </summary>
        /// <param name="column">String value of the column of the cell.</param>
        /// <param name="data">The value to write to the cell.</param>
        /// <param name="numberFormat">Whether the data should be formatted as a number (Prevents scientific notation being used).</param>
        public void writeLastCellInColumn(string column, string data, Boolean numberFormat = false)
        {
            int counter = 1;
            range = worksheet.Range[column + counter];

            while (range.Value != null)
            {
                counter++;
                range = worksheet.Range[column + counter];
            }

            range.Value = data;
            if (numberFormat)
                range.NumberFormat = "#";
        }
    }

}
