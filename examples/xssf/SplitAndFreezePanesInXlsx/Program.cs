using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace NPOI.Examples.XSSF.SplitAndFreezePanesInXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("new sheet");
            ISheet sheet2 = workbook.CreateSheet("second sheet");
            ISheet sheet3 = workbook.CreateSheet("third sheet");
            ISheet sheet4 = workbook.CreateSheet("fourth sheet");

            // Freeze just one row
            sheet1.CreateFreezePane(0, 1, 0, 1);

            // Freeze just one column
            sheet2.CreateFreezePane(1, 0, 1, 0);

            // Freeze the columns and rows (forget about scrolling position of the lower right quadrant).
            sheet3.CreateFreezePane(2, 2);

            // Create a split with the lower left side being the active quadrant
            sheet4.CreateSplitPane(2000, 2000, 0, 0, PanePosition.LowerLeft);

            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
