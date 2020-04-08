using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace NPOI.Examples.XSSF.SetWidthAndHeightInXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");

            //set the width of columns
            sheet1.SetColumnWidth(0, 50 * 256);
            sheet1.SetColumnWidth(1, 100 * 256);
            sheet1.SetColumnWidth(2, 150 * 256);

            //set the width of height
            sheet1.CreateRow(0).Height = 100 * 20;
            sheet1.CreateRow(1).Height = 200 * 20;
            sheet1.CreateRow(2).Height = 300 * 20;

            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
