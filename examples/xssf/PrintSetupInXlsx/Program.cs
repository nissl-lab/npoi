using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace PrintSetupInXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet1 = wb.CreateSheet("Sheet1");
            sheet1.SetMargin(MarginType.RightMargin, 0.5d);
            sheet1.SetMargin(MarginType.TopMargin, 0.6d);
            sheet1.SetMargin(MarginType.LeftMargin, 0.4d);
            sheet1.SetMargin(MarginType.BottomMargin, 0.3d);

            sheet1.PrintSetup.Copies = 3;
            sheet1.PrintSetup.NoColor = true;
            sheet1.PrintSetup.Landscape = true;
            sheet1.PrintSetup.PaperSize = (short)PaperSize.A4 + 1;

            sheet1.FitToPage = true;
            sheet1.PrintSetup.FitHeight = 2;
            sheet1.PrintSetup.FitWidth = 3;
            sheet1.IsPrintGridlines = true;

            sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
            int x = 1;
            for (int i = 1; i <= 15; i++)
            {
                IRow row = sheet1.CreateRow(i);

                for (int j = 0; j < 15; j++)
                    row.CreateCell(j).SetCellValue(x++);
            }

            ISheet sheet2 = wb.CreateSheet("Sheet2");
            sheet2.PrintSetup.Copies = 1;
            sheet2.PrintSetup.Landscape = false;
            sheet2.PrintSetup.Notes = true;
            //sheet2.PrintSetup.EndNote = true;
            //sheet2.PrintSetup.CellError = DisplayCellErrorType.ErrorAsNA;
            sheet2.PrintSetup.PaperSize = (short)PaperSize.A5 + 1;

            x = 100;
            for (int i = 1; i <= 15; i++)
            {
                IRow row = sheet2.CreateRow(i);

                for (int j = 0; j < 15; j++)
                    row.CreateCell(j).SetCellValue(x++);
            }

            FileStream sw = File.Create("test.xlsx");
            wb.Write(sw);
            sw.Close();
        }
    }
}
