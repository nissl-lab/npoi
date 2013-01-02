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


namespace NumberFormatInXls
{
    class Program
    {
        static void Main(string[] args)
        {
            InitializeWorkbook();

            ISheet sheet = hssfworkbook.CreateSheet("new sheet");
            //increase the width of Column A
            sheet.SetColumnWidth(0, 5000);
            //create the format instance
            IDataFormat format = hssfworkbook.CreateDataFormat();

            // Create a row and put some cells in it. Rows are 0 based.
            ICell cell = sheet.CreateRow(0).CreateCell(0);
            //set value for the cell
            cell.SetCellValue(1.2);
            //number format with 2 digits after the decimal point - "1.20"
           ICellStyle cellStyle = hssfworkbook.CreateCellStyle();
            cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");
            cell.CellStyle = cellStyle;

            //RMB currency format with comma    -   "¥20,000"
            ICell cell2 = sheet.CreateRow(1).CreateCell(0);
            cell2.SetCellValue(20000);
           ICellStyle cellStyle2 = hssfworkbook.CreateCellStyle();
            cellStyle2.DataFormat = format.GetFormat("¥#,##0");
            cell2.CellStyle = cellStyle2;
            
            //scentific number format   -   "3.15E+00"
            ICell cell3 = sheet.CreateRow(2).CreateCell(0);
            cell3.SetCellValue(3.151234);
           ICellStyle cellStyle3 = hssfworkbook.CreateCellStyle();
            cellStyle3.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00E+00");
            cell3.CellStyle = cellStyle3;

            //percent format, 2 digits after the decimal point    -  "99.33%"
            ICell cell4 = sheet.CreateRow(3).CreateCell(0);
            cell4.SetCellValue(0.99333);
           ICellStyle cellStyle4 = hssfworkbook.CreateCellStyle();
            cellStyle4.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00%");
            cell4.CellStyle = cellStyle4;
                     
            //phone number format - "021-65881234"
            ICell cell5 = sheet.CreateRow(4).CreateCell(0);
            cell5.SetCellValue( 02165881234);
           ICellStyle cellStyle5 = hssfworkbook.CreateCellStyle();
            cellStyle5.DataFormat = format.GetFormat("000-00000000");
            cell5.CellStyle = cellStyle5;

            //Chinese capitalized character number - 壹贰叁 元
            ICell cell6 = sheet.CreateRow(5).CreateCell(0);
            cell6.SetCellValue(123);
           ICellStyle cellStyle6 = hssfworkbook.CreateCellStyle();
            cellStyle6.DataFormat = format.GetFormat("[DbNum2][$-804]0 元");
            cell6.CellStyle = cellStyle6;

            //Chinese date string
            ICell cell7 = sheet.CreateRow(6).CreateCell(0);
            cell7.SetCellValue(new DateTime(2004,5,6));
           ICellStyle cellStyle7 = hssfworkbook.CreateCellStyle();
            cellStyle7.DataFormat = format.GetFormat("yyyy年m月d日");
            cell7.CellStyle = cellStyle7;


            //Chinese date string
            ICell cell8 = sheet.CreateRow(7).CreateCell(0);
            cell8.SetCellValue(new DateTime(2005, 11, 6));
           ICellStyle cellStyle8 = hssfworkbook.CreateCellStyle();
            cellStyle8.DataFormat = format.GetFormat("yyyy年m月d日");
            cell8.CellStyle = cellStyle8;

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
