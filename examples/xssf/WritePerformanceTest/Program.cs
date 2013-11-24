using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace WritePerformanceTest
{
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
}
}
