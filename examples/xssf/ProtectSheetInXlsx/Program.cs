using System;
using System.Collections.Generic;
using System.Text;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;

namespace NPOI.Examples.XSSF.ProtectSheetInXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet1=(XSSFSheet)workbook.CreateSheet("Sheet A1");

            sheet1.LockFormatRows();
            sheet1.LockFormatCells();
            sheet1.LockFormatColumns();
            sheet1.LockDeleteColumns();
            sheet1.LockDeleteRows();
            sheet1.LockInsertHyperlinks();
            sheet1.LockInsertColumns();
            sheet1.LockInsertRows();
            sheet1.ProtectSheet("password");
            

            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
