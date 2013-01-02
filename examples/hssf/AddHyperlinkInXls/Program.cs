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

using NPOI.HSSF.Util;
using NPOI.SS.UserModel;


/* This sample is copied from poi/hssf/usermodel/examples/Hyperlink.java */
namespace AddHyperlinkInXls
{
    class Program
    {
        static void Main(string[] args)
        {
            InitializeWorkbook();

            ////cell style for hyperlinks
            ////by default hyperlinks are blue and underlined
           ICellStyle hlink_style = hssfworkbook.CreateCellStyle();
            IFont hlink_font = hssfworkbook.CreateFont();
            hlink_font.Underline = (byte)FontUnderlineType.SINGLE;
            hlink_font.Color = HSSFColor.BLUE.index;
            hlink_style.SetFont(hlink_font);

            ICell cell;
            ISheet sheet = hssfworkbook.CreateSheet("Hyperlinks");

            //URL
            cell = sheet.CreateRow(0).CreateCell(0);
            cell.SetCellValue("URL Link");
            HSSFHyperlink link = new HSSFHyperlink(HyperlinkType.URL);
            link.Address = ("http://poi.apache.org/");
            cell.Hyperlink = (link);
            cell.CellStyle = (hlink_style);

            //link to a file in the current directory
            cell = sheet.CreateRow(1).CreateCell(0);
            cell.SetCellValue("File Link");
            link = new HSSFHyperlink(HyperlinkType.FILE);
            link.Address = ("link1.xls");
            cell.Hyperlink = (link);
            cell.CellStyle = (hlink_style);

            //e-mail link
            cell = sheet.CreateRow(2).CreateCell(0);
            cell.SetCellValue("Email Link");
            link = new HSSFHyperlink(HyperlinkType.EMAIL);
            //note, if subject contains white spaces, make sure they are url-encoded
            link.Address = ("mailto:poi@apache.org?subject=Hyperlinks");
            cell.Hyperlink = (link);
            cell.CellStyle = (hlink_style);

            //link to a place in this workbook

            //Create a target sheet and cell
            ISheet sheet2 = hssfworkbook.CreateSheet("Target ISheet");
            sheet2.CreateRow(0).CreateCell(0).SetCellValue("Target ICell");

            cell = sheet.CreateRow(3).CreateCell(0);
            cell.SetCellValue("Worksheet Link");
            link = new HSSFHyperlink(HyperlinkType.DOCUMENT);
            link.Address = ("'Target ISheet'!A1");
            cell.Hyperlink = (link);
            cell.CellStyle = (hlink_style);

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
