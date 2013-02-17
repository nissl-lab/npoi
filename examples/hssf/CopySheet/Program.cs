using System;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using System.IO;

namespace CopySheet
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            //Excel worksheet combine example
            //You will be prompted to select two Excel files. test.xls will be created that combines the sheets
			//Note: This example does not check for duplicate sheet names. Your test files should have different sheet names.
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel document (*.xls)|*.xls";
            ofd.Title = "Select first Excel document";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                HSSFWorkbook book1 = new HSSFWorkbook(new FileStream(ofd.FileName, FileMode.Open));
                ofd.Title = "Select second Excel document";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    HSSFWorkbook book2 = new HSSFWorkbook(new FileStream(ofd.FileName, FileMode.Open));
                    HSSFWorkbook product = new HSSFWorkbook();

                    for (int i = 0; i < book1.NumberOfSheets; i++)
                    {
                        HSSFSheet sheet1 = book1.GetSheetAt(i) as HSSFSheet;
                        sheet1.CopyTo(product, sheet1.SheetName, true, true);
                    }
                    for (int j = 0; j < book2.NumberOfSheets; j++)
                    {
                        HSSFSheet sheet2 = book2.GetSheetAt(j) as HSSFSheet;
                        sheet2.CopyTo(product, sheet2.SheetName, true, true);
                    }
                    product.Write(new FileStream("test.xls", FileMode.Create, FileAccess.ReadWrite));
                }
                
            }
        }
        
    }
}
