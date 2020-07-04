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

using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;

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
            style1.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Blue.Index;
            style1.FillPattern = FillPattern.BigSpots;
            style1.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Pink.Index;
            sheet1.CreateRow(0).CreateCell(0).CellStyle = style1;

            //fill background
            ICellStyle style2 = hssfworkbook.CreateCellStyle();
            style2.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            style2.FillPattern = FillPattern.AltBars;
            style2.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Rose.Index;
            sheet1.CreateRow(1).CreateCell(0).CellStyle = style2;

            //fill background
            ICellStyle style3 = hssfworkbook.CreateCellStyle();
            style3.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Lime.Index;
            style3.FillPattern = FillPattern.LessDots;
            style3.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.LightGreen.Index;
            sheet1.CreateRow(2).CreateCell(0).CellStyle = style3;

            //fill background
            ICellStyle style4 = hssfworkbook.CreateCellStyle();
            style4.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            style4.FillPattern = FillPattern.LeastDots;
            style4.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Rose.Index;
            sheet1.CreateRow(3).CreateCell(0).CellStyle = style4;

            //fill background
            ICellStyle style5 = hssfworkbook.CreateCellStyle();
            style5.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightBlue.Index;
            style5.FillPattern = FillPattern.Bricks;
            style5.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Plum.Index;
            sheet1.CreateRow(4).CreateCell(0).CellStyle = style5;

            //fill background
            ICellStyle style6 = hssfworkbook.CreateCellStyle();
            style6.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.SeaGreen.Index;
            style6.FillPattern = FillPattern.FineDots;
            style6.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.White.Index;
            sheet1.CreateRow(5).CreateCell(0).CellStyle = style6;

            //fill background
            ICellStyle style7 = hssfworkbook.CreateCellStyle();
            style7.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Orange.Index;
            style7.FillPattern = FillPattern.Diamonds;
            style7.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Orchid.Index;
            sheet1.CreateRow(6).CreateCell(0).CellStyle = style7;

            //fill background
            ICellStyle style8 = hssfworkbook.CreateCellStyle();
            style8.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.White.Index;
            style8.FillPattern = FillPattern.Squares;
            style8.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Red.Index;
            sheet1.CreateRow(7).CreateCell(0).CellStyle = style8;

            //fill background
            ICellStyle style9 = hssfworkbook.CreateCellStyle();
            style9.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style9.FillPattern = FillPattern.SparseDots;
            style9.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(8).CreateCell(0).CellStyle = style9;

            //fill background
            ICellStyle style10 = hssfworkbook.CreateCellStyle();
            style10.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style10.FillPattern = FillPattern.ThickBackwardDiagonals;
            style10.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(9).CreateCell(0).CellStyle = style10;

            //fill background
            ICellStyle style11 = hssfworkbook.CreateCellStyle();
            style11.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style11.FillPattern = FillPattern.ThickForwardDiagonals;
            style11.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(10).CreateCell(0).CellStyle = style11;

            //fill background
            ICellStyle style12 = hssfworkbook.CreateCellStyle();
            style12.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style12.FillPattern = FillPattern.ThickHorizontalBands;
            style12.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(11).CreateCell(0).CellStyle = style12;


            //fill background
            ICellStyle style13 = hssfworkbook.CreateCellStyle();
            style13.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style13.FillPattern = FillPattern.ThickVerticalBands;
            style13.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(12).CreateCell(0).CellStyle = style13;

            //fill background
            ICellStyle style14 = hssfworkbook.CreateCellStyle();
            style14.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style14.FillPattern = FillPattern.ThinBackwardDiagonals;
            style14.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(13).CreateCell(0).CellStyle = style14;

            //fill background
            ICellStyle style15 = hssfworkbook.CreateCellStyle();
            style15.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style15.FillPattern = FillPattern.ThinForwardDiagonals;
            style15.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(14).CreateCell(0).CellStyle = style15;

            //fill background
            ICellStyle style16 = hssfworkbook.CreateCellStyle();
            style16.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style16.FillPattern = FillPattern.ThinHorizontalBands;
            style16.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(15).CreateCell(0).CellStyle = style16;

            //fill background
            ICellStyle style17 = hssfworkbook.CreateCellStyle();
            style17.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style17.FillPattern = FillPattern.ThinVerticalBands;
            style17.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
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
