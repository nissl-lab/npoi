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
            style1.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.BLUE.index;
            style1.FillPattern = FillPatternType.BIG_SPOTS;
            style1.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.PINK.index;
            sheet1.CreateRow(0).CreateCell(0).CellStyle = style1;

            //fill background
            ICellStyle style2 = workbook.CreateCellStyle();
            style2.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            style2.FillPattern = FillPatternType.ALT_BARS;
            style2.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.ROSE.index;
            sheet1.CreateRow(1).CreateCell(0).CellStyle = style2;

            //fill background
            ICellStyle style3 = workbook.CreateCellStyle();
            style3.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LIME.index;
            style3.FillPattern = FillPatternType.LESS_DOTS;
            style3.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.LIGHT_GREEN.index;
            sheet1.CreateRow(2).CreateCell(0).CellStyle = style3;

            //fill background
            ICellStyle style4 = workbook.CreateCellStyle();
            style4.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            style4.FillPattern = FillPatternType.LEAST_DOTS;
            style4.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.ROSE.index;
            sheet1.CreateRow(3).CreateCell(0).CellStyle = style4;

            //fill background
            ICellStyle style5 = workbook.CreateCellStyle();
            style5.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LIGHT_BLUE.index;
            style5.FillPattern = FillPatternType.BRICKS;
            style5.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.PLUM.index;
            sheet1.CreateRow(4).CreateCell(0).CellStyle = style5;

            //fill background
            ICellStyle style6 = workbook.CreateCellStyle();
            style6.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.SEA_GREEN.index;
            style6.FillPattern = FillPatternType.FINE_DOTS;
            style6.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.WHITE.index;
            sheet1.CreateRow(5).CreateCell(0).CellStyle = style6;

            //fill background
            ICellStyle style7 = workbook.CreateCellStyle();
            style7.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.ORANGE.index;
            style7.FillPattern = FillPatternType.DIAMONDS;
            style7.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.ORCHID.index;
            sheet1.CreateRow(6).CreateCell(0).CellStyle = style7;

            //fill background
            ICellStyle style8 = workbook.CreateCellStyle();
            style8.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.WHITE.index;
            style8.FillPattern = FillPatternType.SQUARES;
            style8.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.RED.index;
            sheet1.CreateRow(7).CreateCell(0).CellStyle = style8;

            //fill background
            ICellStyle style9 = workbook.CreateCellStyle();
            style9.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.ROYAL_BLUE.index;
            style9.FillPattern = FillPatternType.SPARSE_DOTS;
            style9.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            sheet1.CreateRow(8).CreateCell(0).CellStyle = style9;

            //fill background
            ICellStyle style10 = workbook.CreateCellStyle();
            style10.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.ROYAL_BLUE.index;
            style10.FillPattern = FillPatternType.THICK_BACKWARD_DIAG;
            style10.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            sheet1.CreateRow(9).CreateCell(0).CellStyle = style10;

            //fill background
            ICellStyle style11 = workbook.CreateCellStyle();
            style11.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.ROYAL_BLUE.index;
            style11.FillPattern = FillPatternType.THICK_FORWARD_DIAG;
            style11.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            sheet1.CreateRow(10).CreateCell(0).CellStyle = style11;

            //fill background
            ICellStyle style12 = workbook.CreateCellStyle();
            style12.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.ROYAL_BLUE.index;
            style12.FillPattern = FillPatternType.THICK_HORZ_BANDS;
            style12.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            sheet1.CreateRow(11).CreateCell(0).CellStyle = style12;


            //fill background
            ICellStyle style13 = workbook.CreateCellStyle();
            style13.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.ROYAL_BLUE.index;
            style13.FillPattern = FillPatternType.THICK_VERT_BANDS;
            style13.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            sheet1.CreateRow(12).CreateCell(0).CellStyle = style13;

            //fill background
            ICellStyle style14 = workbook.CreateCellStyle();
            style14.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.ROYAL_BLUE.index;
            style14.FillPattern = FillPatternType.THIN_BACKWARD_DIAG;
            style14.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            sheet1.CreateRow(13).CreateCell(0).CellStyle = style14;

            //fill background
            ICellStyle style15 = workbook.CreateCellStyle();
            style15.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.ROYAL_BLUE.index;
            style15.FillPattern = FillPatternType.THIN_FORWARD_DIAG;
            style15.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            sheet1.CreateRow(14).CreateCell(0).CellStyle = style15;

            //fill background
            ICellStyle style16 = workbook.CreateCellStyle();
            style16.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.ROYAL_BLUE.index;
            style16.FillPattern = FillPatternType.THIN_HORZ_BANDS;
            style16.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            sheet1.CreateRow(15).CreateCell(0).CellStyle = style16;

            //fill background
            ICellStyle style17 = workbook.CreateCellStyle();
            style17.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.ROYAL_BLUE.index;
            style17.FillPattern = FillPatternType.THIN_VERT_BANDS;
            style17.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.YELLOW.index;
            sheet1.CreateRow(16).CreateCell(0).CellStyle = style17;

            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
