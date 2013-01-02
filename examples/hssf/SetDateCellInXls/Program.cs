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


/*
 This sample is copied from poi.hssf.usermodel.examples. Original name is CreateDateCells.java
 */
namespace SetDateCellInXls
{
    class Program
    {
        static void Main(string[] args)
        {
            InitializeWorkbook();

            ISheet sheet = hssfworkbook.CreateSheet("new sheet");
            // Create a row and put some cells in it. Rows are 0 based.
            IRow row = sheet.CreateRow(0);

            // Create a cell and put a date value in it.  The first cell is not styled as a date.
            ICell cell = row.CreateCell(0);
            cell.SetCellValue(DateTime.Now);

            // we style the second cell as a date (and time).  It is important to Create a new cell style from the workbook
            // otherwise you can end up modifying the built in style and effecting not only this cell but other cells.
           ICellStyle cellStyle = hssfworkbook.CreateCellStyle();
            
            // Perhaps this may only works for Chinese date, I don't have english office on hand
            cellStyle.DataFormat = hssfworkbook.CreateDataFormat().GetFormat("[$-409]h:mm:ss AM/PM;@");
            cell.CellStyle=cellStyle;

            //set chinese date format
            ICell cell2 = row.CreateCell(1);
            cell2.SetCellValue(new DateTime(2008, 5, 5));
           ICellStyle cellStyle2 = hssfworkbook.CreateCellStyle();
            IDataFormat format = hssfworkbook.CreateDataFormat();
            cellStyle2.DataFormat = format.GetFormat("yyyy年m月d日");
            cell2.CellStyle = cellStyle2;

            ICell cell3 = row.CreateCell(2);
            cell3.CellFormula = "DateValue(\"2005-11-11 11:11:11\")";
           ICellStyle cellStyle3 = hssfworkbook.CreateCellStyle(); 
            cellStyle3.DataFormat = HSSFDataFormat.GetBuiltinFormat("m/d/yy h:mm");
            cell3.CellStyle = cellStyle3;

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
