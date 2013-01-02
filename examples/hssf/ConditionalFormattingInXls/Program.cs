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


using System;
using System.Text;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.HSSF.Record;
using NPOI.SS.Util;
using NPOI.SS.UserModel;


/* ================================================================
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * NPOI HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/
namespace ConditionalFormattingInXls
{
    class Program
    {
        static void Main(string[] args)
        {
            InitializeWorkbook();

            HSSFSheet sheet1 = (HSSFSheet)hssfworkbook.CreateSheet("Sheet1");

            ISheetConditionalFormatting hscf = sheet1.SheetConditionalFormatting;


            // Define a Conditional Formatting rule, which triggers formatting
            // when cell's value is bigger than 55 and smaller than 500
            // applies patternFormatting defined below.
            IConditionalFormattingRule rule = hscf.CreateConditionalFormattingRule(
                ComparisonOperator.BETWEEN,
                "55", // 1st formula 
                "500"     // 2nd formula 
            );

            // Create pattern with red background
            IPatternFormatting patternFmt = rule.CreatePatternFormatting();
            patternFmt.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.RED.index;

            //// Define a region containing first column
            CellRangeAddress[] regions = {
                new CellRangeAddress(0, 65535,0,1)
            };
            // Apply Conditional Formatting rule defined above to the regions  
            hscf.AddConditionalFormatting(regions, rule);

            //fill cell with numeric values
            sheet1.CreateRow(0).CreateCell(0).SetCellValue(50);
            sheet1.CreateRow(0).CreateCell(1).SetCellValue(101);
            sheet1.CreateRow(1).CreateCell(1).SetCellValue(25);
            sheet1.CreateRow(1).CreateCell(0).SetCellValue(150);

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

