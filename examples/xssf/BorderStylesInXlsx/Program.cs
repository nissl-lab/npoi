using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

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
            style.BorderBottom = BorderStyle.Thin;
            style.BottomBorderColor = IndexedColors.Black.Index;
            style.BorderLeft = BorderStyle.DashDotDot;
            style.LeftBorderColor = IndexedColors.Green.Index;
            style.BorderRight = BorderStyle.Hair;
            style.RightBorderColor = IndexedColors.Blue.Index;
            style.BorderTop = BorderStyle.MediumDashed;
            style.TopBorderColor = IndexedColors.Orange.Index;

            //create border diagonal
            style.BorderDiagonalLineStyle = BorderStyle.Medium; //this property must be set before BorderDiagonal and BorderDiagonalColor
            style.BorderDiagonal = BorderDiagonal.Forward;
            style.BorderDiagonalColor = IndexedColors.Gold.Index;

            cell.CellStyle = style;
            // Create a cell and put a value in it.
            ICell cell2 = row.CreateCell(2);
            cell2.SetCellValue(5);
            ICellStyle style2 = workbook.CreateCellStyle();
            style2.BorderDiagonalLineStyle = BorderStyle.Medium;
            style2.BorderDiagonal = BorderDiagonal.Backward;
            style2.BorderDiagonalColor = IndexedColors.Red.Index;
            cell2.CellStyle = style2;

            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
