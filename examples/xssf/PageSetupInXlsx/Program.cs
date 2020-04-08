using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.IO;

namespace NPOI.Examples.XSSF.PageSetupInXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet1 = wb.CreateSheet("new sheet");
            ISheet sheet2 = wb.CreateSheet("second sheet");

            // Set the columns to repeat from column 0 to 2 on the first sheet
            IRow row1 = sheet1.CreateRow(0);
            row1.CreateCell(0).SetCellValue(1);
            row1.CreateCell(1).SetCellValue(2);
            row1.CreateCell(2).SetCellValue(3);
            IRow row2 = sheet1.CreateRow(1);
            row2.CreateCell(1).SetCellValue(4);
            row2.CreateCell(2).SetCellValue(5);

            IRow row3 = sheet2.CreateRow(1);
            row3.CreateCell(0).SetCellValue(2.1);
            row3.CreateCell(4).SetCellValue(2.2);
            row3.CreateCell(5).SetCellValue(2.3);
            IRow row4 = sheet2.CreateRow(2);
            row4.CreateCell(4).SetCellValue(2.4);
            row4.CreateCell(5).SetCellValue(2.5);

            // Set the columns to repeat from column 0 to 2 on the first sheet
            sheet1.RepeatingColumns = new CellRangeAddress(-1, -1, 0, 2);
            // Set the the repeating rows and columns on the second sheet.
            sheet2.RepeatingColumns = new CellRangeAddress(-1, -1, 4, 5);
            sheet2.RepeatingRows = new CellRangeAddress(1, 2, -1, -1);

            //set the print area for the first sheet
            wb.SetPrintArea(0, 1, 2, 0, 3);

            FileStream sw = File.Create("test.xlsx");
            wb.Write(sw);
            sw.Close();
        }
    }
}