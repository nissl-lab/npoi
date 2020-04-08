using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace NPOI.Examples.XSSF.CreateCommentInXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("some comments");

            // Create the drawing patriarch. This is the top level container for all shapes including cell comments.
            IDrawing patr = sheet.CreateDrawingPatriarch();

            //Create a cell in row 3
            ICell cell1 = sheet.CreateRow(3).CreateCell(1);
            cell1.SetCellValue(new XSSFRichTextString("Hello, World"));

            //anchor defines size and position of the comment in worksheet
            IComment comment1 = patr.CreateCellComment(new XSSFClientAnchor(0, 0, 0, 0, 4, 2, 6, 5));

            // set text in the comment
            comment1.String = new XSSFRichTextString("We can set comments in POI");

            //set comment author.
            //you can see it in the status bar when moving mouse over the commented cell
            comment1.Author = "Apache Software Foundation";

            // The first way to assign comment to a cell is via HSSFCell.SetCellComment method
            cell1.CellComment = comment1;

            //Create another cell in row 6
            ICell cell2 = sheet.CreateRow(6).CreateCell(1);
            cell2.SetCellValue(36.6);


            IComment comment2 = patr.CreateCellComment(new XSSFClientAnchor(0, 0, 0, 0, 4, 8, 6, 11));
            //modify background color of the comment
            //comment2.SetFillColor(204, 236, 255);

            XSSFRichTextString str = new XSSFRichTextString("Normal body temperature");

            //apply custom font to the text in the comment
            IFont font = workbook.CreateFont();
            font.FontName = "Arial";
            font.FontHeightInPoints = 10;
            font.IsBold = true;
            font.Color = HSSFColor.Red.Index;
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

            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
