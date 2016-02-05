using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace MSController
{
    /// <summary>
    /// Handles Microsoft Excel, can read and write to cells and also check if a spreadsheet is open.
    /// </summary>
    public class ExcelHandler
    {
        static Excel.Application excelApp;
        static Excel.Workbooks workbooks;
        static Excel.Workbook workbook;
        static Excel.Sheets worksheets;
        static Excel.Worksheet worksheet;
        static Excel.Range range;
        object missing = System.Reflection.Missing.Value;


        // Open, close, create, isOpen
        /// <summary>
        /// Opens an existing excel spreadsheet for processing.
        /// </summary>
        /// <param name="filePath">The filepath string of the spreadsheet to be opened.</param>
        /// <param name="sheet">The worksheet to open.</param>
        public void open(string filePath, string sheet = "defualt")
        {
            excelApp = new Excel.Application();

            // TODO: Create custom exception
            if (excelApp == null)
                throw new Exception();

            if (!File.Exists(filePath))
                throw new IOException();

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

            range = worksheet.Range["A" + 1];
        }

        /// <summary>
        /// Creates an excel spreadsheet.
        /// </summary>
        /// <param name="filePath">The filepath string of the spreadsheet to be created.</param>
        public void create(string filePath)
        {
            excelApp = new Excel.Application();

            // TODO: Create custom exception
            if (excelApp == null)
                throw new Exception();

            excelApp.Visible = false;
            workbooks = excelApp.Workbooks;
            workbook = workbooks.Add(missing);
            worksheets = workbook.Worksheets;
            worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            range = worksheet.Range["A" + 1];
            workbook.SaveAs(filePath);

            close();
        }

        /// <summary>
        /// Closes the excel spreadsheet.
        /// </summary>
        /// <param name="save">Boolean value of whether or not to save the file.</param>
        public void close(bool save = false)
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
        /// Checks whether a spreadsheet is open or not.
        /// </summary>
        /// <param name="filePath">String value of the column of the cell.</param>
        /// <returns>
        /// True if the file is open, false if not.
        /// </returns>
        public bool isOpen(string filePath)
        {
            // TODO: Create isOpen() method
            return false;
        }


        // Navigation
        /// <summary>
        /// Opens an excel spreadsheet, creates a new sheet then closes it.
        /// </summary>
        /// <param name="filePath">The filepath string of the spreadsheet to be opened.</param>
        /// <param name="sheet">The worksheet to create.</param>
        public void addSheet(string sheet)
        {
            worksheet = (Excel.Worksheet)worksheets.Add(worksheets[1], Type.Missing, Type.Missing, Type.Missing);
            worksheet.Name = sheet;
        }

        /// <summary>
        /// Switches from the current worksheet to the specified one.
        /// </summary>
        /// <param name="sheet">The worksheet to switch to. If it is not found it will instead switch to the default.</param>
        public void changeSheet(string sheet)
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


        // Read
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


        // Write
        /// <summary>
        /// Writes a value to a specified cell in the open spreadsheet.
        /// </summary>
        /// <param name="column">String value of the column of the cell.</param>
        /// <param name="row">Int value of the row of the cell.</param>
        /// <param name="data">The value to write to the cell.</param>
        /// <param name="numberFormat">Whether the data should be formatted as a number (Prevents scientific notation being used).</param>
        public void writeCell(string column, int row, string data, bool numberFormat = false)
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
        public void writeLastCellInColumn(string column, string data, bool numberFormat = false)
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


        // Delete
        /// <summary>
        /// Deletes the specified row from the spreadsheet.
        /// </summary>
        /// <param name="row">The row to delete.</param>
        public void deleteRow(int row)
        {
            // TODO
        }

        /// <summary>
        /// Deletes the specified column from the spreadsheet.
        /// </summary>
        /// <param name="column">The column to delete.</param>
        public void deleteColumn(string column)
        {
            // TODO
        }
    }

}
