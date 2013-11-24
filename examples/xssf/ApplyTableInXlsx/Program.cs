using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ApplyTableInXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet1 = (XSSFSheet)workbook.CreateSheet("Sheet1");
            sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
            int x = 1;
            for (int i = 1; i <= 15; i++)
            {
                IRow row = sheet1.CreateRow(i);
                for (int j = 0; j < 15; j++)
                {
                    row.CreateCell(j).SetCellValue(x++);
                }
            }
            XSSFTable table = sheet1.CreateTable();
            table.Name = "Tabella1";
            table.DisplayName = "Tabella1";

            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
