using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.SS.UserModel;

namespace CopyRowsAndCellsInXls
{
    class Program
    {
        static void Main(string[] args)
        {
            InitializeWorkbook();

            ISheet s = hssfworkbook.GetSheetAt(0);
            
            ICell cell = s.GetRow(4).GetCell(1);
            cell.CopyCellTo(3); //copy B5 to D5

            IRow c = s.GetRow(3);
            c.CopyCell(0, 1);   //copy A4 to B4

            s.CopyRow(0,1);     //copy row A to row B, original row B will be moved to row C automatically
            WriteToFile();
        }
        static HSSFWorkbook hssfworkbook;

        static void WriteToFile()
        {
            //Write the stream data of workbook to the root directory
            FileStream file = new FileStream(@"test.xls", FileMode.Create);
            hssfworkbook.Write(file);
            file.Close();
        }

        static void InitializeWorkbook()
        {
            using (var fs = File.OpenRead(@"Data\test.xls"))
            { 

                hssfworkbook = new HSSFWorkbook(fs);
            }
        }
    }
}
