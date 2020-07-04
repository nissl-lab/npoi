using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.IO;

namespace ConditionalFormats
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook wb = new XSSFWorkbook();

            SameCell(wb.CreateSheet("Same Cell"));
            MultiCell(wb.CreateSheet("MultiCell"));
            Errors(wb.CreateSheet("Errors"));
            HideDupplicates(wb.CreateSheet("Hide Dups"));
            FormatDuplicates(wb.CreateSheet("Duplicates"));
            InList(wb.CreateSheet("In List"));
            Expiry(wb.CreateSheet("Expiry"));
            ShadeAlt(wb.CreateSheet("Shade Alt"));
            ShadeBands(wb.CreateSheet("Shade Bands"));

            using (var fs = File.Create("test.xlsx"))
            {
                wb.Write(fs);
            }
        }
        /**
 * Highlight cells based on their values
 */
        static void SameCell(ISheet sheet)
        {
            sheet.CreateRow(0).CreateCell(0).SetCellValue(84);
            sheet.CreateRow(1).CreateCell(0).SetCellValue(74);
            sheet.CreateRow(2).CreateCell(0).SetCellValue(50);
            sheet.CreateRow(3).CreateCell(0).SetCellValue(51);
            sheet.CreateRow(4).CreateCell(0).SetCellValue(49);
            sheet.CreateRow(5).CreateCell(0).SetCellValue(41);

            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            // Condition 1: Cell Value Is   greater than  70   (Blue Fill)
            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule(ComparisonOperator.GreaterThan, "70");
            IPatternFormatting fill1 = rule1.CreatePatternFormatting();
            fill1.FillBackgroundColor = (IndexedColors.Blue.Index);
            fill1.FillPattern = FillPattern.SolidForeground;

            // Condition 2: Cell Value Is  less than      50   (Green Fill)
            IConditionalFormattingRule rule2 = sheetCF.CreateConditionalFormattingRule(ComparisonOperator.LessThan, "50");
            IPatternFormatting fill2 = rule2.CreatePatternFormatting();
            fill2.FillBackgroundColor = (IndexedColors.Green.Index);
            fill2.FillPattern = FillPattern.SolidForeground;

            CellRangeAddress[] regions = {
                CellRangeAddress.ValueOf("A1:A6")
        };

            sheetCF.AddConditionalFormatting(regions, rule1, rule2);

            sheet.GetRow(0).CreateCell(2).SetCellValue("<== Condition 1: Cell Value is greater than 70 (Blue Fill)");
            sheet.GetRow(4).CreateCell(2).SetCellValue("<== Condition 2: Cell Value is less than 50 (Green Fill)");
        }

        /**
         * Highlight multiple cells based on a formula
         */
        static void MultiCell(ISheet sheet)
        {
            // header row
            IRow row0 = sheet.CreateRow(0);
            row0.CreateCell(0).SetCellValue("Units");
            row0.CreateCell(1).SetCellValue("Cost");
            row0.CreateCell(2).SetCellValue("Total");

            IRow row1 = sheet.CreateRow(1);
            row1.CreateCell(0).SetCellValue(71);
            row1.CreateCell(1).SetCellValue(29);
            row1.CreateCell(2).SetCellValue(2059);

            IRow row2 = sheet.CreateRow(2);
            row2.CreateCell(0).SetCellValue(85);
            row2.CreateCell(1).SetCellValue(29);
            row2.CreateCell(2).SetCellValue(2059);

            IRow row3 = sheet.CreateRow(3);
            row3.CreateCell(0).SetCellValue(71);
            row3.CreateCell(1).SetCellValue(29);
            row3.CreateCell(2).SetCellValue(2059);

            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            // Condition 1: Formula Is   =$B2>75   (Blue Fill)
            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule("$A2>75");
            IPatternFormatting fill1 = rule1.CreatePatternFormatting();
            fill1.FillBackgroundColor = IndexedColors.Blue.Index;
            fill1.FillPattern = FillPattern.SolidForeground;

            CellRangeAddress[] regions = {
                CellRangeAddress.ValueOf("A2:C4")
        };

            sheetCF.AddConditionalFormatting(regions, rule1);

            sheet.GetRow(2).CreateCell(4).SetCellValue("<== Condition 1: Formula is =$B2>75   (Blue Fill)");
        }

        /**
         *  Use Excel conditional formatting to check for errors,
         *  and change the font colour to match the cell colour.
         *  In this example, if formula result is  #DIV/0! then it will have white font colour.
         */
        static void Errors(ISheet sheet)
        {
            sheet.CreateRow(0).CreateCell(0).SetCellValue(84);
            sheet.CreateRow(1).CreateCell(0).SetCellValue(0);
            sheet.CreateRow(2).CreateCell(0).SetCellFormula("ROUND(A1/A2,0)");
            sheet.CreateRow(3).CreateCell(0).SetCellValue(0);
            sheet.CreateRow(4).CreateCell(0).SetCellFormula("ROUND(A6/A4,0)");
            sheet.CreateRow(5).CreateCell(0).SetCellValue(41);

            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            // Condition 1: Formula Is   =ISERROR(C2)   (White Font)
            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule("ISERROR(A1)");
            IFontFormatting font = rule1.CreateFontFormatting();
            font.FontColorIndex = (IndexedColors.White.Index);

            CellRangeAddress[] regions = {
                CellRangeAddress.ValueOf("A1:A6")
        };

            sheetCF.AddConditionalFormatting(regions, rule1);

            sheet.GetRow(2).CreateCell(1).SetCellValue("<== The error in this cell is hidden. Condition: Formula is   =ISERROR(C2)   (White Font)");
            sheet.GetRow(4).CreateCell(1).SetCellValue("<== The error in this cell is hidden. Condition: Formula is   =ISERROR(C2)   (White Font)");
        }

        /**
         * Use Excel conditional formatting to hide the duplicate values,
         * and make the list easier to read. In this example, when the table is sorted by Region,
         * the second (and subsequent) occurences of each region name will have white font colour.
         */
        static void HideDupplicates(ISheet sheet)
        {
            sheet.CreateRow(0).CreateCell(0).SetCellValue("City");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("Boston");
            sheet.CreateRow(2).CreateCell(0).SetCellValue("Boston");
            sheet.CreateRow(3).CreateCell(0).SetCellValue("Chicago");
            sheet.CreateRow(4).CreateCell(0).SetCellValue("Chicago");
            sheet.CreateRow(5).CreateCell(0).SetCellValue("New York");

            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            // Condition 1: Formula Is   =A2=A1   (White Font)
            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule("A2=A1");
            IFontFormatting font = rule1.CreateFontFormatting();
            font.FontColorIndex = IndexedColors.White.Index;

            CellRangeAddress[] regions = {
                CellRangeAddress.ValueOf("A2:A6")
            };

            sheetCF.AddConditionalFormatting(regions, rule1);

            sheet.GetRow(1).CreateCell(1).SetCellValue("<== the second (and subsequent) " +
                    "occurences of each region name will have white font colour.  " +
                    "Condition: Formula Is   =A2=A1   (White Font)");
        }

        /**
         * Use Excel conditional formatting to highlight duplicate entries in a column.
         */
        static void FormatDuplicates(ISheet sheet)
        {
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
            sheet.CreateRow(1).CreateCell(0).SetCellValue(4);
            sheet.CreateRow(2).CreateCell(0).SetCellValue(3);
            sheet.CreateRow(3).CreateCell(0).SetCellValue(6);
            sheet.CreateRow(4).CreateCell(0).SetCellValue(3);
            sheet.CreateRow(5).CreateCell(0).SetCellValue(5);
            sheet.CreateRow(6).CreateCell(0).SetCellValue(8);
            sheet.CreateRow(7).CreateCell(0).SetCellValue(0);
            sheet.CreateRow(8).CreateCell(0).SetCellValue(2);
            sheet.CreateRow(9).CreateCell(0).SetCellValue(8);
            sheet.CreateRow(10).CreateCell(0).SetCellValue(6);

            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            // Condition 1: Formula Is   =A2=A1   (White Font)
            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule("COUNTIF($A$2:$A$11,A2)>1");
            IFontFormatting font = rule1.CreateFontFormatting();
            font.SetFontStyle(false, true);
            font.FontColorIndex = (IndexedColors.Blue.Index);

            CellRangeAddress[] regions = {
                CellRangeAddress.ValueOf("A2:A11")
        };

            sheetCF.AddConditionalFormatting(regions, rule1);

            sheet.GetRow(2).CreateCell(1).SetCellValue("<== Duplicates numbers in the column are highlighted.  " +
                    "Condition: Formula Is =COUNTIF($A$2:$A$11,A2)>1   (Blue Font)");
        }

        /**
         * Use Excel conditional formatting to highlight items that are in a list on the worksheet.
         */
        static void InList(ISheet sheet)
        {
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Codes");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("AA");
            sheet.CreateRow(2).CreateCell(0).SetCellValue("BB");
            sheet.CreateRow(3).CreateCell(0).SetCellValue("GG");
            sheet.CreateRow(4).CreateCell(0).SetCellValue("AA");
            sheet.CreateRow(5).CreateCell(0).SetCellValue("FF");
            sheet.CreateRow(6).CreateCell(0).SetCellValue("XX");
            sheet.CreateRow(7).CreateCell(0).SetCellValue("CC");

            sheet.GetRow(0).CreateCell(2).SetCellValue("Valid");
            sheet.GetRow(1).CreateCell(2).SetCellValue("AA");
            sheet.GetRow(2).CreateCell(2).SetCellValue("BB");
            sheet.GetRow(3).CreateCell(2).SetCellValue("CC");

            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            // Condition 1: Formula Is   =A2=A1   (White Font)
            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule("COUNTIF($C$2:$C$4,A2)");
            IPatternFormatting fill1 = rule1.CreatePatternFormatting();
            fill1.FillBackgroundColor = (IndexedColors.LightBlue.Index);
            fill1.FillPattern = FillPattern.SolidForeground;

            CellRangeAddress[] regions = {
                CellRangeAddress.ValueOf("A2:A8")
        };

            sheetCF.AddConditionalFormatting(regions, rule1);

            sheet.GetRow(2).CreateCell(3).SetCellValue("<== Use Excel conditional formatting to highlight items that are in a list on the worksheet");
        }

        /**
         *  Use Excel conditional formatting to highlight payments that are due in the next thirty days.
         *  In this example, Due dates are entered in cells A2:A4.
         */
        static void Expiry(ISheet sheet)
        {
            ICellStyle style = sheet.Workbook.CreateCellStyle();
            style.DataFormat = (short)BuiltinFormats.GetBuiltinFormat("d-mmm");

            sheet.CreateRow(0).CreateCell(0).SetCellValue("Date");
            sheet.CreateRow(1).CreateCell(0).SetCellFormula("TODAY()+29");
            sheet.CreateRow(2).CreateCell(0).SetCellFormula("A2+1");
            sheet.CreateRow(3).CreateCell(0).SetCellFormula("A3+1");

            for (int rownum = 1; rownum <= 3; rownum++) sheet.GetRow(rownum).GetCell(0).CellStyle = style;

            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            // Condition 1: Formula Is   =A2=A1   (White Font)
            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule("AND(A2-TODAY()>=0,A2-TODAY()<=30)");
            IFontFormatting font = rule1.CreateFontFormatting();
            font.SetFontStyle(false, true);
            font.FontColorIndex = IndexedColors.Blue.Index;

            CellRangeAddress[] regions = {
                CellRangeAddress.ValueOf("A2:A4")
        };

            sheetCF.AddConditionalFormatting(regions, rule1);

            sheet.GetRow(0).CreateCell(1).SetCellValue("Dates within the next 30 days are highlighted");
        }

        /**
         * Use Excel conditional formatting to shade alternating rows on the worksheet
         */
        static void ShadeAlt(ISheet sheet)
        {
            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            // Condition 1: Formula Is   =A2=A1   (White Font)
            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule("MOD(ROW(),2)");
            IPatternFormatting fill1 = rule1.CreatePatternFormatting();
            fill1.FillBackgroundColor = (IndexedColors.LightGreen.Index);
            fill1.FillPattern = FillPattern.SolidForeground;

            CellRangeAddress[] regions = {
                CellRangeAddress.ValueOf("A1:Z100")
        };

            sheetCF.AddConditionalFormatting(regions, rule1);

            sheet.CreateRow(0).CreateCell(1).SetCellValue("Shade Alternating Rows");
            sheet.CreateRow(1).CreateCell(1).SetCellValue("Condition: Formula Is  =MOD(ROW(),2)   (Light Green Fill)");
        }

        /**
         * You can use Excel conditional formatting to shade bands of rows on the worksheet. 
         * In this example, 3 rows are shaded light grey, and 3 are left with no shading.
         * In the MOD function, the total number of rows in the set of banded rows (6) is entered.
         */
        static void ShadeBands(ISheet sheet)
        {
            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule("MOD(ROW(),6)<3");
            IPatternFormatting fill1 = rule1.CreatePatternFormatting();
            fill1.FillBackgroundColor = (IndexedColors.Grey25Percent.Index);
            fill1.FillPattern = FillPattern.SolidForeground;

            CellRangeAddress[] regions = {
                CellRangeAddress.ValueOf("A1:Z100")
        };

            sheetCF.AddConditionalFormatting(regions, rule1);

            sheet.CreateRow(0).CreateCell(1).SetCellValue("Shade Bands of Rows");
            sheet.CreateRow(1).CreateCell(1).SetCellValue("Condition: Formula is  =MOD(ROW(),6)<2   (Light Grey Fill)");
        }
    }
}
