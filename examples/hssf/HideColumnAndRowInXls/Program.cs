using System;
using System.Text;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;


namespace HideColumnAndRowInXls
{
    class Program
    {
        static void Main(string[] args)
        {
            InitializeWorkbook();

            Sheet s = hssfworkbook.CreateSheet("Sheet1");
            Row r1 = s.CreateRow(0);
            Row r2 = s.CreateRow(1);
            Row r3 = s.CreateRow(2);
            Row r4 = s.CreateRow(3);
            Row r5 = s.CreateRow(4);

            //hide Row 2
            r2.ZeroHeight = true;

            //hide column C
            s.SetColumnHidden(2, true);

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
