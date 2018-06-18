using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace MicroStationTagExplorer
{
    internal class Excelnterop : IDisposable
    {
        private Excel.Application _application;
        private Excel.Workbook _workbook;
        private Excel.Worksheet _worksheet;

        public void CreateApplication()
        {
            _application = ComUtilities.CreateObject<Excel.Application>("Excel.Application");
            _application.Visible = true;
            _application.ErrorCheckingOptions.NumberAsText = false;
        }

        public void CreateWorkbook()
        {
            _workbook = _application.Workbooks.Add();
        }

        private void Format(Excel.Range range, int nColumns, string tableName)
        {
            _worksheet.Rows["2:2"].Select();
            _application.ActiveWindow.SplitColumn = 0;
            _application.ActiveWindow.SplitRow = 1;
            _application.ActiveWindow.FreezePanes = true;

            Excel.Range header = (Excel.Range)_worksheet.Range[_worksheet.Cells[1, 1], _worksheet.Cells[1, 1]];
            header.CurrentRegion.Select();
            _application.Selection.AutoFilter();

            Excel.Range columns = _worksheet.Columns[1].Resize(Type.Missing, nColumns);
            columns.Select();
            columns.EntireColumn.AutoFit();

            _worksheet.ListObjects.AddEx(Excel.XlListObjectSourceType.xlSrcRange, range, null, Excel.XlYesNoGuess.xlYes).Name = tableName;

            _worksheet.Range["A1"].Select();
        }

        public void ExportValues(object[,] values, int nRows, int nColumns, string tableName)
        {
            _worksheet = (Excel.Worksheet)_workbook.Worksheets.Add();

            Excel.Range range = (Excel.Range)_worksheet.Range[_worksheet.Cells[1, 1], _worksheet.Cells[nRows, nColumns]];

            range.NumberFormat = "@";
            range.Value = values;

            Format(range, nColumns, tableName);
        }

        public void Quit()
        {
            if (_application != null)
            {
                _application.Quit();
            }
        }

        public void Dispose()
        {
            ComUtilities.ReleaseComObject(_worksheet);
            _worksheet = null;

            ComUtilities.ReleaseComObject(_workbook);
            _workbook = null;

            ComUtilities.ReleaseComObject(_application);
            _application = null;
        }
    }
}
