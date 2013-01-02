using System;
using System.Collections.Generic;
using System.Text;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using NPOI.HSSF.Util;

namespace NPOI.Examples.XSSF.BorderStylesInXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet A1");
            IRow row = sheet.CreateRow(1);
            // Create a cell and put a value in it.
            ICell cell = row.CreateCell(1);
            cell.SetCellValue(4);

            // Style the cell with borders all around.
            ICellStyle style = workbook.CreateCellStyle();
            style.BorderBottom = BorderStyle.THIN;
            style.BottomBorderColor = HSSFColor.BLACK.index;
            style.BorderLeft = BorderStyle.DASH_DOT_DOT;
            style.LeftBorderColor = HSSFColor.GREEN.index;
            style.BorderRight = BorderStyle.HAIR;
            style.RightBorderColor = HSSFColor.BLUE.index;
            style.BorderTop = BorderStyle.MEDIUM_DASHED;
            style.TopBorderColor = HSSFColor.ORANGE.index;

            style.BorderDiagonalLineStyle = BorderStyle.MEDIUM; //this property must be set before BorderDiagonal and BorderDiagonalColor
            style.BorderDiagonal = BorderDiagonal.FORWARD;
            style.BorderDiagonalColor = HSSFColor.GOLD.index;

            cell.CellStyle = style;
            // Create a cell and put a value in it.
            ICell cell2 = row.CreateCell(2);
            cell2.SetCellValue(5);
            ICellStyle style2 = workbook.CreateCellStyle();
            style2.BorderDiagonalLineStyle = BorderStyle.MEDIUM;
            style2.BorderDiagonal = BorderDiagonal.BACKWARD;
            style2.BorderDiagonalColor = HSSFColor.RED.index;
            cell2.CellStyle = style2;

            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
