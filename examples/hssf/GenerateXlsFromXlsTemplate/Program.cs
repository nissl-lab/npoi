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


namespace GenerateXlsFromXlsTemplate
{
    class Program
    {
        static void Main(string[] args)
        {
            InitializeWorkbook();

            ISheet sheet1 = hssfworkbook.GetSheet("Sheet1");
            //create cell on rows, since rows do already exist,it's not necessary to create rows again.
            sheet1.GetRow(1).GetCell(1).SetCellValue(200200);
            sheet1.GetRow(2).GetCell(1).SetCellValue(300);
            sheet1.GetRow(3).GetCell(1).SetCellValue(500050);
            sheet1.GetRow(4).GetCell(1).SetCellValue(8000);
            sheet1.GetRow(5).GetCell(1).SetCellValue(110);
            sheet1.GetRow(6).GetCell(1).SetCellValue(100);
            sheet1.GetRow(7).GetCell(1).SetCellValue(200);
            sheet1.GetRow(8).GetCell(1).SetCellValue(210);
            sheet1.GetRow(9).GetCell(1).SetCellValue(2300);
            sheet1.GetRow(10).GetCell(1).SetCellValue(240);
            sheet1.GetRow(11).GetCell(1).SetCellValue(180123);
            sheet1.GetRow(12).GetCell(1).SetCellValue(150);

            //Force excel to recalculate all the formula while open
            sheet1.ForceFormulaRecalculation = true;

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
            //read the template via FileStream, it is suggested to use FileAccess.Read to prevent file lock.
            //book1.xls is an Excel-2007-generated file, so some new unknown BIFF records are added. 
            FileStream file = new FileStream(@"template/book1.xls", FileMode.Open,FileAccess.Read);

            hssfworkbook = new HSSFWorkbook(file);

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
