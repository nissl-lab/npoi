using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace WritePerformanceTest
{
    /*
class Program
{
    static void Main(string[] args)
    {
        Stopwatch sw = new Stopwatch();
        sw.Start();

        IWorkbook workbook = new XSSFWorkbook();
        ICell cell;
        ISheet sheet = workbook.CreateSheet("StressTest");
        int i = 0;
        int rowLimit = 100000;

        DateTime originalTime = DateTime.Now;

        System.Console.WriteLine("Start time: " + originalTime);

        for (i = 0; i <= rowLimit; i++)
        {
            cell = sheet.CreateRow(i).CreateCell(0);
            //sheet.SetActiveCell(i, 0);
            cell.SetCellValue("ZOMG PLEASE SURVIVE THIS STRESS TEST");

            if(i % 10000 == 0)
            {
                System.Console.WriteLine("[" + (DateTime.Now - originalTime) + "]" + " " + i + " rows written");
            }
        }
        
        FileStream sw1 = File.Create("test.xlsx");
        workbook.Write(sw1);
        sw1.Close();

        sw.Stop();
        Console.WriteLine("Time: "+sw.ElapsedMilliseconds+" ms");
        //prompt user so we do not close the window with our data :)
        System.Console.Read();
    }
}*/

    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");
            sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
            int x = 1;

            Console.WriteLine("Start at " + DateTime.Now.ToString());
            for (int i = 1; i <= 70000; i++)
            {
                IRow row = sheet1.CreateRow(i);
                for (int j = 0; j < 15; j++)
                {
                    row.CreateCell(j).SetCellValue(x++);
                }
            }
            Console.WriteLine("End at " + DateTime.Now.ToString());

            FileStream sw = File.Create("test.xls");
            workbook.Write(sw);
            sw.Close();

            Console.Read();
        }
    }
}
