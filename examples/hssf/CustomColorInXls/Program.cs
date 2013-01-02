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
using System.Collections.Generic;
using System.Text;

using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;

namespace CustomColorInXls
{
    class Program
    {
        static void Main(string[] args)
        {
            InitializeWorkbook();


            HSSFPalette palette = workbook.GetCustomPalette();
            palette.SetColorAtIndex(HSSFColor.PINK.index, (byte)255, (byte)234, (byte)222);
            //HSSFColor myColor = palette.AddColor((byte)253, (byte)0, (byte)0);

            ISheet sheet1 = workbook.CreateSheet("Sheet1");
            ICellStyle style1 = workbook.CreateCellStyle();
            style1.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.PINK.index;
            style1.FillPattern = FillPatternType.SOLID_FOREGROUND;
            sheet1.CreateRow(0).CreateCell(0).CellStyle = style1;

            WriteToFile();
        }

        static HSSFWorkbook workbook;

        static void WriteToFile()
        {
            //Write the stream data of workbook to the root directory
            FileStream file = new FileStream(@"test.xls", FileMode.Create);
            workbook.Write(file);
            file.Close();
        }

        static void InitializeWorkbook()
        {
            workbook = new HSSFWorkbook();

            ////create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            workbook.DocumentSummaryInformation = dsi;

            ////create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            workbook.SummaryInformation = si;
        }
    }
}
