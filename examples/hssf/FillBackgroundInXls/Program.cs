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


namespace FillBackgroundInXls
{
    class Program
    {
        static void Main(string[] args)
        {
            InitializeWorkbook();

            ISheet sheet1 = hssfworkbook.CreateSheet("Sheet1");
            
            //fill background
           ICellStyle style1 = hssfworkbook.CreateCellStyle();
            style1.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.BLUE.index;
            style1.FillPattern = FillPatternType.BIG_SPOTS;
            style1.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.PINK.index;
            sheet1.CreateRow(0).CreateCell(0).CellStyle = style1;

            //fill background
           ICellStyle style2 = hssfworkbook.CreateCellStyle();
            style2.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            style2.FillPattern = FillPatternType.ALT_BARS;
            style2.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.ROSE.index;
            sheet1.CreateRow(1).CreateCell(0).CellStyle = style2;

            //fill background
           ICellStyle style3 = hssfworkbook.CreateCellStyle();
            style3.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LIME.index;
            style3.FillPattern = FillPatternType.LESS_DOTS;
            style3.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.LIGHT_GREEN.index;
            sheet1.CreateRow(2).CreateCell(0).CellStyle = style3;

            //fill background
           ICellStyle style4 = hssfworkbook.CreateCellStyle();
            style4.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            style4.FillPattern = FillPatternType.LEAST_DOTS;
            style4.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.ROSE.index;
            sheet1.CreateRow(3).CreateCell(0).CellStyle = style4;

            //fill background
           ICellStyle style5 = hssfworkbook.CreateCellStyle();
            style5.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LIGHT_BLUE.index;
            style5.FillPattern = FillPatternType.BRICKS;
            style5.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.PLUM.index;
            sheet1.CreateRow(4).CreateCell(0).CellStyle = style5;

            //fill background
           ICellStyle style6 = hssfworkbook.CreateCellStyle();
            style6.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.SEA_GREEN.index;
            style6.FillPattern = FillPatternType.FINE_DOTS;
            style6.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.WHITE.index;
            sheet1.CreateRow(5).CreateCell(0).CellStyle = style6;

            //fill background
           ICellStyle style7 = hssfworkbook.CreateCellStyle();
            style7.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.ORANGE.index;
            style7.FillPattern = FillPatternType.DIAMONDS;
            style7.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.ORCHID.index;
            sheet1.CreateRow(6).CreateCell(0).CellStyle = style7;

            //fill background
           ICellStyle style8 = hssfworkbook.CreateCellStyle();
            style8.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.WHITE.index;
            style8.FillPattern = FillPatternType.SQUARES;
            style8.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.RED.index;
            sheet1.CreateRow(7).CreateCell(0).CellStyle = style8;

            //fill background
           ICellStyle style9 = hssfworkbook.CreateCellStyle();
            style9.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.ROYAL_BLUE.index;
            style9.FillPattern = FillPatternType.SPARSE_DOTS;
            style9.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            sheet1.CreateRow(8).CreateCell(0).CellStyle = style9;

            //fill background
           ICellStyle style10 = hssfworkbook.CreateCellStyle();
            style10.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.ROYAL_BLUE.index;
            style10.FillPattern = FillPatternType.THICK_BACKWARD_DIAG;
            style10.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            sheet1.CreateRow(9).CreateCell(0).CellStyle = style10;

            //fill background
           ICellStyle style11 = hssfworkbook.CreateCellStyle();
            style11.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.ROYAL_BLUE.index;
            style11.FillPattern = FillPatternType.THICK_FORWARD_DIAG;
            style11.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            sheet1.CreateRow(10).CreateCell(0).CellStyle = style11;

            //fill background
           ICellStyle style12 = hssfworkbook.CreateCellStyle();
            style12.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.ROYAL_BLUE.index;
            style12.FillPattern = FillPatternType.THICK_HORZ_BANDS;
            style12.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            sheet1.CreateRow(11).CreateCell(0).CellStyle = style12;


            //fill background
           ICellStyle style13 = hssfworkbook.CreateCellStyle();
            style13.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.ROYAL_BLUE.index;
            style13.FillPattern = FillPatternType.THICK_VERT_BANDS;
            style13.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            sheet1.CreateRow(12).CreateCell(0).CellStyle = style13;

            //fill background
           ICellStyle style14 = hssfworkbook.CreateCellStyle();
            style14.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.ROYAL_BLUE.index;
            style14.FillPattern = FillPatternType.THIN_BACKWARD_DIAG;
            style14.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            sheet1.CreateRow(13).CreateCell(0).CellStyle = style14;

            //fill background
           ICellStyle style15 = hssfworkbook.CreateCellStyle();
            style15.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.ROYAL_BLUE.index;
            style15.FillPattern = FillPatternType.THIN_FORWARD_DIAG;
            style15.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            sheet1.CreateRow(14).CreateCell(0).CellStyle = style15;

            //fill background
           ICellStyle style16 = hssfworkbook.CreateCellStyle();
            style16.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.ROYAL_BLUE.index;
            style16.FillPattern = FillPatternType.THIN_HORZ_BANDS;
            style16.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            sheet1.CreateRow(15).CreateCell(0).CellStyle = style16;

            //fill background
           ICellStyle style17 = hssfworkbook.CreateCellStyle();
            style17.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.ROYAL_BLUE.index;
            style17.FillPattern = FillPatternType.THIN_VERT_BANDS;
            style17.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            sheet1.CreateRow(16).CreateCell(0).CellStyle = style17;


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
