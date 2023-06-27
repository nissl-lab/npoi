/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

using NPOI.SS.UserModel;
using System;
using System.Text;
using NPOI.Util;
using NPOI.HSSF.UserModel;
using System.IO;
namespace TestCases.SS.Util
{


    /**
     * Creates a spreadsheet that Checks Excel's comparison of various IEEE double values.
     * The class {@link NumberComparisonExamples} Contains specific comparison examples
     * (along with their expected results) that Get encoded into rows of the spreadsheet.
     * Each example is Checked with a formula (in column I) that displays either "OK" or
     * "ERROR" depending on whether actual results match those expected.
     *
     * @author Josh Micich
     */
    public class NumberComparingSpreadsheetGenerator
    {

        private class SheetWriter
        {

            private ISheet _sheet;
            private int _rowIndex;

            public SheetWriter(IWorkbook wb)
            {
                ISheet sheet = wb.CreateSheet("Sheet1");

                WriteHeaderRow(wb, sheet);
                _sheet = sheet;
                _rowIndex = 1;
            }

            public void AddTestRow(double a, double b, int expResult)
            {
                WriteDataRow(_sheet, _rowIndex++, a, b, expResult);
            }
        }

        private static void WriteHeaderCell(IRow row, int i, String text, ICellStyle style)
        {
            ICell cell = row.CreateCell(i);
            cell.SetCellValue(new HSSFRichTextString(text));
            cell.CellStyle = (style);
        }
        static void WriteHeaderRow(IWorkbook wb, ISheet sheet)
        {
            sheet.SetColumnWidth(0, 6000);
            sheet.SetColumnWidth(1, 6000);
            sheet.SetColumnWidth(2, 3600);
            sheet.SetColumnWidth(3, 3600);
            sheet.SetColumnWidth(4, 2400);
            sheet.SetColumnWidth(5, 2400);
            sheet.SetColumnWidth(6, 2400);
            sheet.SetColumnWidth(7, 2400);
            sheet.SetColumnWidth(8, 2400);
            IRow row = sheet.CreateRow(0);
            ICellStyle style = wb.CreateCellStyle();
            IFont font = wb.CreateFont();
            font.IsBold = true;
            style.SetFont(font);
            WriteHeaderCell(row, 0, "Raw Long Bits A", style);
            WriteHeaderCell(row, 1, "Raw Long Bits B", style);
            WriteHeaderCell(row, 2, "Value A", style);
            WriteHeaderCell(row, 3, "Value B", style);
            WriteHeaderCell(row, 4, "Exp Cmp", style);
            WriteHeaderCell(row, 5, "LT", style);
            WriteHeaderCell(row, 6, "EQ", style);
            WriteHeaderCell(row, 7, "GT", style);
            WriteHeaderCell(row, 8, "Check", style);
        }
        /**
         * Fills a spreadsheet row with one comparison example. The two numeric values are written to
         * columns C and D. Columns (F, G and H) respectively Get formulas ("v0<v1", "v0=v1", "v0>v1"),
         * which will be Evaluated by Excel. Column D Gets the expected comparison result. Column I
         * Gets a formula to check that Excel's comparison results match that predicted in column D.
         *
         * @param v0 the first value to be Compared
         * @param v1 the second value to be Compared
         * @param expRes expected comparison result (-1, 0, or +1)
         */
        static void WriteDataRow(ISheet sheet, int rowIx, double v0, double v1, int expRes)
        {
            IRow row = sheet.CreateRow(rowIx);

            int rowNum = rowIx + 1;


            row.CreateCell(0).SetCellValue(FormatDoubleAsHex(v0));
            row.CreateCell(1).SetCellValue(FormatDoubleAsHex(v1));
            row.CreateCell(2).SetCellValue(v0);
            row.CreateCell(3).SetCellValue(v1);
            row.CreateCell(4).SetCellValue(expRes < 0 ? "LT" : expRes > 0 ? "GT" : "EQ");
            row.CreateCell(5).CellFormula = ("C" + rowNum + "<" + "D" + rowNum);
            row.CreateCell(6).CellFormula = ("C" + rowNum + "=" + "D" + rowNum);
            row.CreateCell(7).CellFormula = ("C" + rowNum + ">" + "D" + rowNum);
            // TODO - bug elsewhere in POI - something wrong with encoding of NOT() function
            String frm = "if(or(" +
                "and(E#='LT', F#      , G#=FALSE, H#=FALSE)," +
                "and(E#='EQ', F#=FALSE, G#      , H#=FALSE)," +
                "and(E#='GT', F#=FALSE, G#=FALSE, H#      )" +
                "), 'OK', 'error')";
            row.CreateCell(8).CellFormula = (frm.Replace("#", rowNum.ToString()).Replace('\'', '"'));
        }

        private static String FormatDoubleAsHex(double d)
        {
            long l = BitConverter.DoubleToInt64Bits(d);
            StringBuilder sb = new StringBuilder(20);
            sb.Append(HexDump.LongToHex(l)).Append('L');
            return sb.ToString();
        }

        //public static void Main(String[] args)
        //{

        //    IWorkbook wb = new HSSFWorkbook();
        //    SheetWriter sw = new SheetWriter(wb);
        //    ComparisonExample[] ces = NumberComparisonExamples.GetComparisonExamples();
        //    for (int i = 0; i < ces.Length; i++)
        //    {
        //        ComparisonExample ce = ces[i];
        //        sw.AddTestRow(ce.GetA(), ce.GetB(), ce.GetExpectedResult());
        //    }


        //    FileInfo outputFile = new FileInfo("ExcelNumberCompare.xls");

        //    FileStream os = File.OpenWrite(outputFile.FullName);
        //    wb.Write(os);
        //    os.Close();

        //    Console.WriteLine("Finished writing '" + outputFile.FullName + "'");
        //}
    }
}
