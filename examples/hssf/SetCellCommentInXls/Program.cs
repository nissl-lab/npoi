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
using NPOI.HSSF.Util;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;


/*
 This sample is copied from poi.hssf.usermodel.examples. Original name is CellComments.java
 */
namespace SetCellCommentInXls
{
    class Program
    {
        static void Main(string[] args)
        {
            InitializeWorkbook();

            ISheet sheet = hssfworkbook.CreateSheet("ICell comments in POI HSSF");

            // Create the drawing patriarch. This is the top level container for all shapes including cell comments.
            IDrawing patr = (HSSFPatriarch)sheet.CreateDrawingPatriarch();

            //Create a cell in row 3
            ICell cell1 = sheet.CreateRow(3).CreateCell(1);
            cell1.SetCellValue(new HSSFRichTextString("Hello, World"));

            //anchor defines size and position of the comment in worksheet
            IComment comment1 = patr.CreateCellComment(new HSSFClientAnchor(0, 0, 0, 0, 4, 2, 6, 5));

             // set text in the comment
            comment1.String = (new HSSFRichTextString("We can set comments in POI"));

            //set comment author.
            //you can see it in the status bar when moving mouse over the commented cell
            comment1.Author = ("Apache Software Foundation");

            // The first way to assign comment to a cell is via HSSFCell.SetCellComment method
            cell1.CellComment = (comment1);

            //Create another cell in row 6
            ICell cell2 = sheet.CreateRow(6).CreateCell(1);
            cell2.SetCellValue(36.6);


            HSSFComment comment2 = (HSSFComment)patr.CreateCellComment(new HSSFClientAnchor(0, 0, 0, 0, 4, 8, 6, 11));
            //modify background color of the comment
            comment2.SetFillColor(204, 236, 255);

            HSSFRichTextString str = new HSSFRichTextString("Normal body temperature");

            //apply custom font to the text in the comment
            IFont font = hssfworkbook.CreateFont();
            font.FontName = ("Arial");
            font.FontHeightInPoints =10;
            font.Boldweight = (short)FontBoldWeight.BOLD;
            font.Color = HSSFColor.RED.index;
            str.ApplyFont(font);

            comment2.String = str;
            comment2.Visible = true; //by default comments are hidden. This one is always visible.

            comment2.Author = "Bill Gates";

            /**
             * The second way to assign comment to a cell is to implicitly specify its row and column.
             * Note, it is possible to set row and column of a non-existing cell.
             * It works, the commnet is visible.
             */
            comment2.Row = 6;
            comment2.Column = 1;

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
