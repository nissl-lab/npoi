using System;
using System.Text;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;


namespace ShrinkToFitColumnInXls
{
    class Program
    {
        static void Main(string[] args)
        {
            InitializeWorkbook();

            Sheet sheet = hssfworkbook.CreateSheet("Sheet1");
            Row row = sheet.CreateRow(0);
            //create cell value
            Cell cell1 = row.CreateCell(0);
            cell1.SetCellValue("This is a test");
            //apply ShrinkToFit to cellstyle
            CellStyle cellstyle1 = hssfworkbook.CreateCellStyle();
            cellstyle1.ShrinkToFit = true;
            cell1.CellStyle = cellstyle1;
            //create cell value
            row.CreateCell(1).SetCellValue("Hello World");
            row.GetCell(1).CellStyle = cellstyle1;

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
            hssfworkbook = new HSSFWorkbook();

            //create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            //create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;
        }
    }
}
