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
using System.Collections.Generic;
using NPOI.SS.Util;
using NPOI.SS.UserModel;


namespace SetActiveCellRangeInXls
{
    class Program
    {
        static void Main(string[] args)
        {
            InitializeWorkbook();
            //use HSSFCell.SetAsActiveCell() to select B6 as the active column
            ISheet sheet1 = hssfworkbook.CreateSheet("ISheet A");
            CreateCellArray(sheet1);
            sheet1.GetRow(5).GetCell(1).SetAsActiveCell();
            //set TopRow and LeftCol to make B6 the first cell in the visible area
            sheet1.TopRow = 5;
            sheet1.LeftCol = 1;

            //use ISheet.SetActiveCell(), the sheet can be empty
            ISheet sheet2 = hssfworkbook.CreateSheet("ISheet B");
            sheet2.SetActiveCell(1, 5);

            //use ISheet.SetActiveCellRange to select a cell range
            ISheet sheet3 = hssfworkbook.CreateSheet("ISheet C");
            CreateCellArray(sheet3);
            sheet3.SetActiveCellRange(2, 20, 1, 50);
            //set the ISheet C as the active sheet
            hssfworkbook.SetActiveSheet(2);

            //use ISheet.SetActiveCellRange to select multiple cell ranges
            ISheet sheet4 = hssfworkbook.CreateSheet("ISheet D");
            CreateCellArray(sheet4);
            List<CellRangeAddress8Bit> cellranges = new List<CellRangeAddress8Bit>();
            cellranges.Add(new CellRangeAddress8Bit(1,5,10,100));
            cellranges.Add(new CellRangeAddress8Bit(6,7,8,9));
            sheet4.SetActiveCellRange(cellranges,1,6,9);

            WriteToFile();
        }

        static void CreateCellArray(ISheet sheet)
        {
            for (int i = 0; i < 300; i++)
            {
                IRow row=sheet.CreateRow(i);
                for (int j = 0; j < 150; j++)
                {
                    ICell cell = row.CreateCell(j);
                    cell.SetCellValue(i*j);
                }
            }
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
