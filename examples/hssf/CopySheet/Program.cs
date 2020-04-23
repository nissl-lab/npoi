using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.IO;

namespace CopySheet
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World! Please Wait while processing...");
            //Excel worksheet combine example
            //Note: This example does not check for duplicate sheet names. Your test files should have different sheet names.

            IWorkbook book1 = new HSSFWorkbook(new FileStream("file1.xls", FileMode.Open));
            IWorkbook book2 = new HSSFWorkbook(new FileStream("file2.xls", FileMode.Open));
            IWorkbook product = new HSSFWorkbook();

            for (int i = 0; i < book1.NumberOfSheets; i++)
            {
                ISheet sheet1 = book1.GetSheetAt(i);
                sheet1.CopyTo(product, sheet1.SheetName, true, true);
            }

            for (int j = 0; j < book2.NumberOfSheets; j++)
            {
                ISheet sheet2 = book2.GetSheetAt(j);
                sheet2.CopyTo(product, sheet2.SheetName, true, true);
            }

            product.Write(new FileStream("test.xls", FileMode.Create, FileAccess.ReadWrite));
        }
    }
}
