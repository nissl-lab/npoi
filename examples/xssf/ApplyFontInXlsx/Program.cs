using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace NPOI.Examples.XSSF.ApplyFontInXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");

            //font style1: underlined, italic, red color, fontsize=20
            IFont font1 = workbook.CreateFont();
            font1.Color = IndexedColors.Red.Index;
            font1.IsItalic = true;
            font1.Underline = FontUnderlineType.Double;
            font1.FontHeightInPoints = 20;

            //bind font with style 1
            ICellStyle style1 = workbook.CreateCellStyle();
            style1.SetFont(font1);

            //font style2: strikeout line, green color, fontsize=15, fontname='宋体'
            IFont font2 = workbook.CreateFont();
            font2.Color = IndexedColors.OliveGreen.Index;
            font2.IsStrikeout = true;
            font2.FontHeightInPoints = 15;
            font2.FontName = "宋体";

            //bind font with style 2
            ICellStyle style2 = workbook.CreateCellStyle();
            style2.SetFont(font2);

            //apply font styles
            ICell cell1 = sheet1.CreateRow(1).CreateCell(1);
            cell1.SetCellValue("Hello World!");
            cell1.CellStyle = style1;
            ICell cell2 = sheet1.CreateRow(3).CreateCell(1);
            cell2.SetCellValue("早上好！");
            cell2.CellStyle = style2;

            ////cell with rich text 
            ICell cell3 = sheet1.CreateRow(5).CreateCell(1);
            XSSFRichTextString richtext = new XSSFRichTextString("Microsoft OfficeTM");

            //apply font to "Microsoft Office"
            IFont font4 = workbook.CreateFont();
            font4.FontHeightInPoints = 12;
            richtext.ApplyFont(0, 16, font4);
            //apply font to "TM"
            IFont font3 = workbook.CreateFont();
            font3.TypeOffset = FontSuperScript.Super;
            font3.IsItalic = true;
            font3.Color = IndexedColors.Blue.Index;
            font3.FontHeightInPoints = 8;
            richtext.ApplyFont(16, 18, font3);

            cell3.SetCellValue(richtext);

            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
