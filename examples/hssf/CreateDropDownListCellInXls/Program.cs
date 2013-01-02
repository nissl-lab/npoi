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
using NPOI.SS.UserModel;
using NPOI.SS.Util;


/* ================================================================
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * NPOI HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/
namespace CreateDropDownListCellInXls
{
    class Program
    {
        static void Main(string[] args)
        {
            InitializeWorkbook();

            ISheet sheet1 = hssfworkbook.CreateSheet("Sheet1");
            ISheet sheet2 = hssfworkbook.CreateSheet("Sheet2");
            //create three items in Sheet2
            IRow row0 = sheet2.CreateRow(0);
            ICell cell0 = row0.CreateCell(4);
            cell0.SetCellValue("Product1");

            row0 = sheet2.CreateRow(1);
            cell0 = row0.CreateCell(4);
            cell0.SetCellValue("Product2");

            row0 = sheet2.CreateRow(2);
            cell0 = row0.CreateCell(4);
            cell0.SetCellValue("Product3");


            CellRangeAddressList rangeList = new CellRangeAddressList();

            //add the data validation to the first column (1-100 rows) 
            rangeList.AddCellRangeAddress(new CellRangeAddress(1, 100, 0, 0));
            DVConstraint dvconstraint = DVConstraint.CreateFormulaListConstraint("Sheet2!$E1:$E3");
            HSSFDataValidation dataValidation = new
                    HSSFDataValidation(rangeList, dvconstraint);
            //add the data validation to sheet1
            ((HSSFSheet)sheet1).AddValidationData(dataValidation); 

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
