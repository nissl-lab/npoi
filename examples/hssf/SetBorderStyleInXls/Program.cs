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
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System.IO;

/*
 This sample is copied from poi.hssf.usermodel.examples. Original name is Borders.java
 */
namespace SetBorderStyleInXls
{
    class Program
    {
        static HSSFWorkbook hssfworkbook;

        static void Main(string[] args)
        {
            InitializeWorkbook();

            ISheet sheet = hssfworkbook.CreateSheet("new sheet");

            // Create a row and put some cells in it. Rows are 0 based.
            IRow row = sheet.CreateRow(1);

            // Create a cell and put a value in it.
            ICell cell = row.CreateCell(1);
            cell.SetCellValue(4);

            // Style the cell with borders all around.
            ICellStyle style = hssfworkbook.CreateCellStyle();
            style.BorderBottom = BorderStyle.Thin;
            style.BottomBorderColor = HSSFColor.Black.Index;
            style.BorderLeft = BorderStyle.DashDotDot;
            style.LeftBorderColor = HSSFColor.Green.Index;
            style.BorderRight = BorderStyle.Hair;
            style.RightBorderColor = HSSFColor.Blue.Index;
            style.BorderTop = BorderStyle.MediumDashed;
            style.TopBorderColor = HSSFColor.Orange.Index;

            style.BorderDiagonal = BorderDiagonal.Forward;
            style.BorderDiagonalColor = HSSFColor.Gold.Index;
            style.BorderDiagonalLineStyle = BorderStyle.Medium;

            cell.CellStyle = style;
            // Create a cell and put a value in it.
            ICell cell2 = row.CreateCell(2);
            cell2.SetCellValue(5);
            ICellStyle style2 = hssfworkbook.CreateCellStyle();
            style2.BorderDiagonal = BorderDiagonal.Backward;
            style2.BorderDiagonalColor = HSSFColor.Red.Index;
            style2.BorderDiagonalLineStyle = BorderStyle.Medium;
            cell2.CellStyle = style2;

            WriteToFile();
        }

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

            //Create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            //Create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;
        }
    }
}
