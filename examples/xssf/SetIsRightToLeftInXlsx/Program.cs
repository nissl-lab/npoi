using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace NPOI.Examples.XSSF.SetIsRightToLeftInXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook workbook = new XSSFWorkbook();

            ISheet sheet1 = workbook.CreateSheet("Sheet1");
            
            // Setting support for Right To Left
            sheet1.IsRightToLeft = true;

            sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");

            int x = 1;

            for(int i = 1; i <= 15; i++)
            {
                IRow row = sheet1.CreateRow(i);

                for(int j = 0; j < 15; j++)
                {
                    row.CreateCell(j).SetCellValue(x++);
                }
            }

            FileStream sw = File.Create("test.xlsx");

            workbook.Write(sw);

            sw.Close();
        }
    }
}
