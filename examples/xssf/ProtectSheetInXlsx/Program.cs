using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace NPOI.Examples.XSSF.ProtectSheetInXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet1 = (XSSFSheet)workbook.CreateSheet("Sheet A1");

            sheet1.LockFormatRows(true);
            sheet1.LockFormatCells(true);
            sheet1.LockFormatColumns(true);
            sheet1.LockDeleteColumns(true);
            sheet1.LockDeleteRows(true);
            sheet1.LockInsertHyperlinks(true);
            sheet1.LockInsertColumns(true);
            sheet1.LockInsertRows(true);
            sheet1.ProtectSheet("password");

            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
