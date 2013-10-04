using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;

namespace NPOI.Examples.XSSF.FillBackgroundInXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook workbook = new XSSFWorkbook();

            ISheet sheet1 = workbook.CreateSheet("Sheet1");

            //fill background
            ICellStyle style1 = workbook.CreateCellStyle();
            style1.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Blue.Index;
            style1.FillPattern = FillPattern.BigSpots;
            style1.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Pink.Index;
            sheet1.CreateRow(0).CreateCell(0).CellStyle = style1;

            //fill background
            ICellStyle style2 = workbook.CreateCellStyle();
            style2.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            style2.FillPattern = FillPattern.AltBars;
            style2.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Rose.Index;
            sheet1.CreateRow(1).CreateCell(0).CellStyle = style2;

            //fill background
            ICellStyle style3 = workbook.CreateCellStyle();
            style3.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Lime.Index;
            style3.FillPattern = FillPattern.LessDots;
            style3.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.LightGreen.Index;
            sheet1.CreateRow(2).CreateCell(0).CellStyle = style3;

            //fill background
            ICellStyle style4 = workbook.CreateCellStyle();
            style4.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            style4.FillPattern = FillPattern.LeastDots;
            style4.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Rose.Index;
            sheet1.CreateRow(3).CreateCell(0).CellStyle = style4;

            //fill background
            ICellStyle style5 = workbook.CreateCellStyle();
            style5.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightBlue.Index;
            style5.FillPattern = FillPattern.Bricks;
            style5.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Plum.Index;
            sheet1.CreateRow(4).CreateCell(0).CellStyle = style5;

            //fill background
            ICellStyle style6 = workbook.CreateCellStyle();
            style6.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.SeaGreen.Index;
            style6.FillPattern = FillPattern.FineDots;
            style6.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.White.Index;
            sheet1.CreateRow(5).CreateCell(0).CellStyle = style6;

            //fill background
            ICellStyle style7 = workbook.CreateCellStyle();
            style7.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Orange.Index;
            style7.FillPattern = FillPattern.Diamonds;
            style7.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Orchid.Index;
            sheet1.CreateRow(6).CreateCell(0).CellStyle = style7;

            //fill background
            ICellStyle style8 = workbook.CreateCellStyle();
            style8.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.White.Index;
            style8.FillPattern = FillPattern.Squares;
            style8.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Red.Index;
            sheet1.CreateRow(7).CreateCell(0).CellStyle = style8;

            //fill background
            ICellStyle style9 = workbook.CreateCellStyle();
            style9.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style9.FillPattern = FillPattern.SparseDots;
            style9.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(8).CreateCell(0).CellStyle = style9;

            //fill background
            ICellStyle style10 = workbook.CreateCellStyle();
            style10.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style10.FillPattern = FillPattern.ThinBackwardDiagonals;
            style10.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(9).CreateCell(0).CellStyle = style10;

            //fill background
            ICellStyle style11 = workbook.CreateCellStyle();
            style11.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style11.FillPattern = FillPattern.ThickForwardDiagonals;
            style11.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(10).CreateCell(0).CellStyle = style11;

            //fill background
            ICellStyle style12 = workbook.CreateCellStyle();
            style12.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style12.FillPattern = FillPattern.ThickHorizontalBands;
            style12.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(11).CreateCell(0).CellStyle = style12;


            //fill background
            ICellStyle style13 = workbook.CreateCellStyle();
            style13.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style13.FillPattern = FillPattern.ThickVerticalBands;
            style13.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(12).CreateCell(0).CellStyle = style13;

            //fill background
            ICellStyle style14 = workbook.CreateCellStyle();
            style14.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style14.FillPattern = FillPattern.ThickBackwardDiagonals;
            style14.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(13).CreateCell(0).CellStyle = style14;

            //fill background
            ICellStyle style15 = workbook.CreateCellStyle();
            style15.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style15.FillPattern = FillPattern.ThinForwardDiagonals;
            style15.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(14).CreateCell(0).CellStyle = style15;

            //fill background
            ICellStyle style16 = workbook.CreateCellStyle();
            style16.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style16.FillPattern = FillPattern.ThinHorizontalBands;
            style16.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(15).CreateCell(0).CellStyle = style16;

            //fill background
            ICellStyle style17 = workbook.CreateCellStyle();
            style17.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style17.FillPattern = FillPattern.ThinVerticalBands;
            style17.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(16).CreateCell(0).CellStyle = style17;

            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
