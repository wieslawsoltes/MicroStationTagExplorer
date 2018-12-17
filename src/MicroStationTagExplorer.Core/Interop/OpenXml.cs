using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace MicroStationTagExplorer.Core.Interop
{
    internal class OpenXml
    {
        public static void ExportValues(string path, object[,] values, int nRows, int nColumns, string sheetName, string tableName)
        {
            var spreadsheetDocument = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);

            var workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = sheetName
            };
            sheets.Append(sheet);

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            WriteValues(sheetData, values, nRows, nColumns);

            workbookpart.Workbook.Save();
            spreadsheetDocument.Close();
        }

        public static void WriteValues(SheetData sheetData, object[,] values, int nRows, int nColumns)
        {
            for (int r = 0; r < nRows; r++)
            {
                Row row = new Row();
                sheetData.Append(row);
                Cell previous = null;
                for (int c = 0; c < nColumns; c++)
                {
                    Cell cell = new Cell();
                    row.InsertAfter(cell, previous);
                    cell.CellValue = new CellValue(values[r, c].ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                    previous = cell;
                }
            }
        }
    }
}
