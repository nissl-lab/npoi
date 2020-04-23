using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace NPOI.Examples.XSSF.AddHyperlinkInXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ////cell style for hyperlinks
            ////by default hyperlinks are blue and underlined
            ICellStyle hlink_style = workbook.CreateCellStyle();
            IFont hlink_font = workbook.CreateFont();
            hlink_font.Underline = FontUnderlineType.Single;
            hlink_font.Color = HSSFColor.Blue.Index;
            hlink_style.SetFont(hlink_font);

            ICell cell;
            ISheet sheet = workbook.CreateSheet("Hyperlinks");

            //URL
            cell = sheet.CreateRow(0).CreateCell(0);
            cell.SetCellValue("URL Link");
            XSSFHyperlink link = new XSSFHyperlink(HyperlinkType.Url);
            link.Address = ("http://poi.apache.org/");
            cell.Hyperlink = (link);
            cell.CellStyle = (hlink_style);

            //link to a file in the current directory
            cell = sheet.CreateRow(1).CreateCell(0);
            cell.SetCellValue("File Link");
            link = new XSSFHyperlink(HyperlinkType.File);
            link.Address = ("link1.xls");
            cell.Hyperlink = (link);
            cell.CellStyle = (hlink_style);

            //e-mail link
            cell = sheet.CreateRow(2).CreateCell(0);
            cell.SetCellValue("Email Link");
            link = new XSSFHyperlink(HyperlinkType.Email);
            //note, if subject contains white spaces, make sure they are url-encoded
            link.Address = ("mailto:poi@apache.org?subject=Hyperlinks");
            cell.Hyperlink = (link);
            cell.CellStyle = (hlink_style);

            //link to a place in this workbook

            //Create a target sheet and cell
            ISheet sheet2 = workbook.CreateSheet("Target ISheet");
            sheet2.CreateRow(0).CreateCell(0).SetCellValue("Target ICell");

            cell = sheet.CreateRow(3).CreateCell(0);
            cell.SetCellValue("Worksheet Link");
            link = new XSSFHyperlink(HyperlinkType.Document);
            link.Address = ("'Target ISheet'!A1");
            cell.Hyperlink = (link);
            cell.CellStyle = (hlink_style);

            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
