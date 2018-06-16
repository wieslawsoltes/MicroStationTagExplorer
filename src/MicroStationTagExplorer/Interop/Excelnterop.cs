using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace MicroStationTagExplorer
{
    internal class Excelnterop : IDisposable
    {
        private Excel.Application _application;
        private Excel.Workbook _workbook;
        private Excel.Worksheet _worksheet;

        public Excelnterop()
        {
            Initialize();
        }

        private void Initialize()
        {
            _application = ComUtilities.CreateObject<Excel.Application>("Excel.Application");
            _application.Visible = true;
            _application.ErrorCheckingOptions.NumberAsText = false;
            _workbook = _application.Workbooks.Add();
        }

        public void ExportTags(object[,] values, int rows, int columns)
        {
            _worksheet = _workbook.Worksheets.Add();

            Excel.Range start = _worksheet.Cells[1, 1];
            Excel.Range end = _worksheet.Cells[rows, columns];
            Excel.Range range = _worksheet.Range[start, end];

            range.NumberFormat = "@";
            range.Value = values;

            _worksheet.Rows["2:2"].Select();
            _application.ActiveWindow.SplitColumn = 0;
            _application.ActiveWindow.SplitRow = 1;
            _application.ActiveWindow.FreezePanes = true;

            _worksheet.Range[_worksheet.Cells[1, 1], _worksheet.Cells[1, 1]].CurrentRegion.Select();
            _application.Selection.AutoFilter();

            _worksheet.Columns[1].Resize(Type.Missing, 6).Select();
            _worksheet.Columns[1].Resize(Type.Missing, 6).EntireColumn.AutoFit();

            _worksheet.ListObjects.AddEx(Excel.XlListObjectSourceType.xlSrcRange, range, null, Excel.XlYesNoGuess.xlYes).Name = "Tags";

            _worksheet.Range["A1"].Select();
        }

        public void Dispose()
        {
            ComUtilities.ReleaseComObject(_worksheet);
            _worksheet = null;

            ComUtilities.ReleaseComObject(_workbook);
            _workbook = null;

            //if (_application != null)
            //{
            //    _application.Quit();
            //}

            ComUtilities.ReleaseComObject(_application);
            _application = null;
        }
    }
}
