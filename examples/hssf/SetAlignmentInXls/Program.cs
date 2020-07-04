/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

/* ================================================================
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * NPOI HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/

using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;

namespace SetAlignmentInXls
{
    class Program
    {
        static void Main(string[] args)
        {
            InitializeWorkbook();

            ISheet sheet1 = hssfworkbook.CreateSheet("Sheet1");

            //set the column width respectively
            sheet1.SetColumnWidth(0, 3000);
            sheet1.SetColumnWidth(1, 3000);
            sheet1.SetColumnWidth(2, 3000);

            for (int i = 1; i <= 10; i++)
            {
                //create the row
                IRow row = sheet1.CreateRow(i);
                //set height of the row
                row.HeightInPoints = 100;

                //create the first cell
                row.CreateCell(0).SetCellValue("Left");
                ICellStyle styleLeft = hssfworkbook.CreateCellStyle();
                styleLeft.Alignment = HorizontalAlignment.Left;
                styleLeft.VerticalAlignment = VerticalAlignment.Top;
                row.GetCell(0).CellStyle = styleLeft;
                //set indention for the text in the cell
                styleLeft.Indention = 3;

                //create the second cell
                row.CreateCell(1).SetCellValue("Center Hello World Hello WorldHello WorldHello WorldHello WorldHello World");
                ICellStyle styleMiddle = hssfworkbook.CreateCellStyle();
                styleMiddle.Alignment = HorizontalAlignment.Center;
                styleMiddle.VerticalAlignment = VerticalAlignment.Center;
                row.GetCell(1).CellStyle = styleMiddle;
                //wrap the text in the cell
                styleMiddle.WrapText = true;


                //create the third cell
                row.CreateCell(2).SetCellValue("Right");
                ICellStyle styleRight = hssfworkbook.CreateCellStyle();
                styleRight.Alignment = HorizontalAlignment.Justify;
                styleRight.VerticalAlignment = VerticalAlignment.Bottom;
                row.GetCell(2).CellStyle = styleRight;

            }

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

            ////create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            ////create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;
        }
    }
}
