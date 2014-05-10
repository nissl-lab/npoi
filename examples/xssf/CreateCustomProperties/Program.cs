using NPOI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CreateCustomProperties
{
    class Program
    {
        static void Main(string[] args)
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");

            POIXMLProperties props = workbook.GetProperties();
            props.CoreProperties.Creator = "NPOI 2.0.5";
            props.CoreProperties.Created = DateTime.Now;
            if (!props.CustomProperties.Contains("NPOI Team"))
                props.CustomProperties.AddProperty("NPOI Team", "Hello World!");

            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
