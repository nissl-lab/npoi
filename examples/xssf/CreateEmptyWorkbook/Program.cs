using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace NPOI.Examples.XSSF.CreateEmptyWorkbook
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook workbook = new XSSFWorkbook();
            workbook.CreateSheet("Sheet A1");
            workbook.CreateSheet("Sheet A2");
            workbook.CreateSheet("Sheet A3");

            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
