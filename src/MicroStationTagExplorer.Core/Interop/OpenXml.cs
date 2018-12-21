﻿using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace MicroStationTagExplorer.Core.Interop
{
    internal class OpenXml
    {
        private static string ToRange(uint columnNumber)
        {
            uint end = columnNumber;
            string name = string.Empty;
            while (end > 0)
            {
                uint modulo = (end - 1) % 26;
                name = Convert.ToChar(65 + modulo).ToString() + name;
                end = (uint)((end - modulo) / 26);
            }
            return name;
        }

        private static string ToRange(uint startRow, uint endRow, uint startColumn, uint endColumn)
        {
            return ToRange(startColumn) + startRow.ToString() + ":" + ToRange(endColumn) + endRow.ToString();
        }

        public static void ExportValues(string path, object[,] values, uint nRows, uint nColumns, string sheetName, string tableName)
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
            WriteTable(worksheetPart, 1U, values, nRows, nColumns);

            workbookpart.Workbook.Save();
            spreadsheetDocument.Close();
        }

        public static void WriteValues(SheetData sheetData, object[,] values, uint nRows, uint nColumns)
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

        private static void WriteTable(WorksheetPart worksheetPart, uint tableID, object[,] values, uint nRows, uint nColumns)
        {
            string range = ToRange(1, nRows, 1, nColumns);

            var autoFilter = new AutoFilter
            {
                Reference = range
            };

            var tableColumns = new TableColumns()
            {
                Count = nColumns
            };

            for (int c = 0; c < nColumns; c++)
            {
                var column = new TableColumn()
                {
                    Id = (UInt32Value)(c + 1U),
                    Name = values[0, c].ToString()
                };
                tableColumns.Append(column);
            }

            var tableStyleInfo = new TableStyleInfo()
            {
                Name = "TableStyleMedium2",
                ShowFirstColumn = true,
                ShowLastColumn = false,
                ShowRowStripes = true,
                ShowColumnStripes = true
            };

            var table = new Table()
            {
                Id = tableID,
                Name = $"Table{tableID}",
                DisplayName = $"Table{tableID}",
                Reference = range,
                TotalsRowShown = false
            };

            table.Append(autoFilter);
            table.Append(tableColumns);
            table.Append(tableStyleInfo);

            var tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>($"rId{tableID}");
            tableDefinitionPart.Table = table;
        }
    }
}
