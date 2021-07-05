using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace dsuexcellibrary
{
    public static class ExcelWriter
    {
        public static async Task<string> WriteHeaderAsync(string spreadsheetName, List<string> rowData)
        {
            string excelName = string.Format("{0}{1}", spreadsheetName, DateTime.Now.ToString("yyyyMMddHHmmss"));
            string excelPath = Path.Combine(Directory.GetCurrentDirectory(), "ExcelFile");
            if (!Directory.Exists(excelPath))
            {
                Directory.CreateDirectory(excelPath);
            }

            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            var fileName = string.Format("{0}.xlsx", excelName);
            fileName = Path.Combine(excelPath, fileName);
            if (File.Exists(fileName))
                File.Delete(fileName);
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = spreadsheetName };
            sheets.Append(sheet);

            Worksheet worksheet = new Worksheet();
            SheetData sheetData = new SheetData();

            #region Header
            if (rowData.Count > 0)
            {
                var row = new Row() { RowIndex = 1 };

                //List<Cell> cells = new List<Cell>(row.Elements<Cell>());
                //System.Console.WriteLine(cells.ToString());
                row.InsertAt<Cell>(new Cell() { CellReference = GetExcelColumnName(0), DataType = CellValues.String, CellValue = new CellValue("No") }, 0);

                var idx = 1;
                foreach (var item in rowData)
                {
                    var columnName = GetExcelColumnName(idx);
                    //System.Console.WriteLine("Column Name : " + columnName);
                    row.InsertAt<Cell>(new Cell() { CellReference = columnName, DataType = CellValues.String, CellValue = new CellValue(item) }, idx);
                    idx++;
                }
                sheetData.Append(row);

            }
            #endregion

            worksheet.Append(sheetData);
            worksheetPart.Worksheet = worksheet;

            workbookpart.Workbook.Save();

            // Close the document.
            spreadsheetDocument.Close();

            return fileName;
        }

        public static async Task<string> WriteBodyAsync(string spreadsheetName, List<List<string>> rowsData, int startAtRow)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(spreadsheetName, true))
            {
                WorksheetPart worksheetPart;
                Worksheet worksheet;
                SheetData sheetData;
                if (spreadsheetDocument.WorkbookPart.WorksheetParts.FirstOrDefault() != null)
                {
                    // System.Console.WriteLine("enter if");
                    worksheetPart = spreadsheetDocument.WorkbookPart.WorksheetParts.FirstOrDefault();
                    worksheet = worksheetPart.Worksheet;
                    sheetData = worksheet.Descendants<SheetData>().FirstOrDefault();
                    // System.Console.WriteLine(sheetData.InnerText);
                }
                else
                {
                    // System.Console.WriteLine("enter else");
                    worksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
                    worksheet = new Worksheet();
                    worksheetPart.Worksheet = worksheet;
                    sheetData = new SheetData();
                }
                uint rowIndex = (uint)startAtRow + 2;
                int numbering = startAtRow + 1;
                #region Body
                foreach (var rowData in rowsData)
                {
                    var row = new Row() { RowIndex = rowIndex };
                    var idx = 1;
                    row.InsertAt<Cell>(new Cell() { CellReference = GetExcelColumnName(0), DataType = CellValues.Number, CellValue = new CellValue(numbering) }, 0);
                    foreach (var item in rowData)
                    {
                        var columnName = GetExcelColumnName(idx);
                        // System.Console.WriteLine("Column Name : " + columnName);
                        row.InsertAt<Cell>(new Cell() { CellReference = columnName, DataType = CellValues.String, CellValue = new CellValue(item == null ? "null" : item.ToString()) }, idx);
                        idx++;
                    }
                    sheetData.Append(row);
                    rowIndex++;
                    numbering++;
                    // System.Console.WriteLine("Row Index : " + rowIndex);
                }
                #endregion

                //worksheet.Append(sheetData);
                // worksheetPart.Worksheet = worksheet;

                spreadsheetDocument.WorkbookPart.Workbook.Save();

                // Close the document.
                spreadsheetDocument.Close();
            }
            return string.Empty;
        }
        static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
    }
}