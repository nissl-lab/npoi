using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace BigGridTest
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet worksheet = workbook.CreateSheet("Sheet1");

            for (int rownum = 0; rownum < 10000; rownum++)
            {
                IRow row = worksheet.CreateRow(rownum);
                for (int celnum = 0; celnum < 20; celnum++)
                {
                    ICell Cell = row.CreateCell(celnum);
                    Cell.SetCellValue("Cell: Row-" + rownum + ";CellNo:" + celnum);
                }
            }

            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
