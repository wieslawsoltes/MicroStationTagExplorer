using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace MicroStationTagExplorer
{
    public static class Excelnterop
    {
        public static void ExportTags(object[,] values, int rows, int columns)
        {
            Excel.Application app = Utilities.CreateObject<Excel.Application>("Excel.Application");
            app.Visible = true;

            Excel.Workbook wb = app.Workbooks.Add();
            Excel.Worksheet ws = wb.Worksheets.Add();

            Excel.Range start = ws.Cells[1, 1];
            Excel.Range end = ws.Cells[rows, columns];
            Excel.Range range = ws.Range[start, end];

            range.Value = values;

            ws.Rows["2:2"].Select();
            app.ActiveWindow.SplitColumn = 0;
            app.ActiveWindow.SplitRow = 1;
            app.ActiveWindow.FreezePanes = true;

            ws.Range[ws.Cells[1, 1], ws.Cells[1, 1]].CurrentRegion.Select();
            app.Selection.AutoFilter();

            ws.Columns[1].Resize(Type.Missing, 6).Select();
            ws.Columns[1].Resize(Type.Missing, 6).EntireColumn.AutoFit();

            ws.ListObjects.AddEx(Excel.XlListObjectSourceType.xlSrcRange, range, null, Excel.XlYesNoGuess.xlYes).Name = "Tags";

            ws.Range["A1"].Select();
        }
    }
}
