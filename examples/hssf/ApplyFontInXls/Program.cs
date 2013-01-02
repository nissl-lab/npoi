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
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;

namespace ApplyFontInXls
{
    class Program
    {
        static void Main(string[] args)
        {
            InitializeWorkbook();

            ISheet sheet1=hssfworkbook.CreateSheet("Sheet1");

            //font style1: underlined, italic, red color, fontsize=20
            IFont font1 = hssfworkbook.CreateFont();
            font1.Color = HSSFColor.RED.index;
            font1.IsItalic = true;
            font1.Underline = (byte)FontUnderlineType.DOUBLE;
            font1.FontHeightInPoints = 20;

            //bind font with style 1
           ICellStyle style1 = hssfworkbook.CreateCellStyle();
            style1.SetFont(font1);

            //font style2: strikeout line, green color, fontsize=15, fontname='宋体'
            IFont font2 = hssfworkbook.CreateFont();
            font2.Color = HSSFColor.OLIVE_GREEN.index;
            font2.IsStrikeout=true;
            font2.FontHeightInPoints = 15;
            font2.FontName = "宋体";

            //bind font with style 2
           ICellStyle style2 = hssfworkbook.CreateCellStyle();
            style2.SetFont(font2);
            
            //apply font styles
            ICell cell1 = HSSFCellUtil.CreateCell(sheet1.CreateRow(1), 1, "Hello World!");
            cell1.CellStyle = style1;
            ICell cell2 = HSSFCellUtil.CreateCell(sheet1.CreateRow(3), 1, "早上好！");
            cell2.CellStyle = style2;

            //cell with rich text 
            ICell cell3 = sheet1.CreateRow(5).CreateCell(1);
            HSSFRichTextString richtext = new HSSFRichTextString("Microsoft OfficeTM");

            //apply font to "Microsoft Office"
            IFont font4 = hssfworkbook.CreateFont();
            font4.FontHeightInPoints = 12;
            richtext.ApplyFont(0, 16, font4);
            //apply font to "TM"
            IFont font3=hssfworkbook.CreateFont();
            font3.TypeOffset = (short)FontSuperScript.SUPER;
            font3.IsItalic = true;
            font3.Color = HSSFColor.BLUE.index;
            font3.FontHeightInPoints=8;
            richtext.ApplyFont(16, 18,font3);
            
            cell3.SetCellValue(richtext);

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
