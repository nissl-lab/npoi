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

using System;
using System.Text;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;

/* This sample is migrated from poi\hssf\usermodel\examples\RepeatingRowsAndColumns.java */

//This sample shows you how to set repeat rows or columns in print page
namespace RepeatingRowsAndColumns
{
    class Program
    {
        static void Main(string[] args)
        {
            InitializeWorkbook();

            ISheet sheet1 = hssfworkbook.CreateSheet("first sheet");
            hssfworkbook.CreateSheet("second sheet");
            hssfworkbook.CreateSheet("third sheet");

            IFont boldFont = hssfworkbook.CreateFont();
            boldFont.FontHeightInPoints = 22;
            boldFont.Boldweight = (short)FontBoldWeight.BOLD;

           ICellStyle boldStyle = hssfworkbook.CreateCellStyle();
            boldStyle.SetFont(boldFont);

            IRow row = sheet1.CreateRow(1);
            ICell cell = row.CreateCell(0);
            cell.SetCellValue("This quick brown fox");
            cell.CellStyle = (boldStyle);

            // Set the columns to repeat from column 0 to 2 on the first sheet
            hssfworkbook.SetRepeatingRowsAndColumns(0, 0, 2, -1, -1);
            // Set the rows to repeat from row 0 to 2 on the second sheet.
            hssfworkbook.SetRepeatingRowsAndColumns(1, -1, -1, 0, 2);
            // Set the the repeating rows and columns on the third sheet.
            hssfworkbook.SetRepeatingRowsAndColumns(2, 4, 5, 1, 2);

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
