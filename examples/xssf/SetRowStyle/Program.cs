using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace SetRowStyle
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet s1 = workbook.CreateSheet("Sheet1");

            ICellStyle rowstyle = workbook.CreateCellStyle();
            rowstyle.FillForegroundColor = IndexedColors.Red.Index;
            rowstyle.FillPattern = FillPattern.SolidForeground;

            ICellStyle c1Style = workbook.CreateCellStyle();
            c1Style.FillForegroundColor = IndexedColors.Yellow.Index;
            c1Style.FillPattern = FillPattern.SolidForeground;

            IRow r1 = s1.CreateRow(1);
            IRow r2= s1.CreateRow(2);
            r1.RowStyle = rowstyle;
            r2.RowStyle = rowstyle;

            ICell c1 = r2.CreateCell(2);
            c1.CellStyle = c1Style;
            c1.SetCellValue("Test");

            ICell c4 = r2.CreateCell(4);
            c4.CellStyle = c1Style;

            using(var fs=File.Create("test.xlsx"))
            {
                workbook.Write(fs);
            }
        }
    }
}
