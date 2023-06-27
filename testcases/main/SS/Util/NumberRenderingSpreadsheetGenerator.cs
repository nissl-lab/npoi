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

using NPOI.HSSF.UserModel;
using System.Collections.Generic;
using System;
using NPOI.Util;
using NPOI.SS.UserModel;
using System.Text;
using System.IO;

namespace TestCases.SS.Util
{

    /**
     * Creates a spreadsheet that demonstrates Excel's rendering of various IEEE double values.
     * 
     * @author Josh Micich
     */
    public class NumberRenderingSpreadsheetGenerator
    {

        private class SheetWriter
        {

            private ISheet _sheet;
            private int _rowIndex;
            private List<long> _ReplacementNaNs;

            public SheetWriter(HSSFWorkbook wb)
            {
                ISheet sheet = wb.CreateSheet("Sheet1");

                WriteHeaderRow(wb, sheet);
                _sheet = sheet;
                _rowIndex = 1;
                _ReplacementNaNs = new List<long>();
            }

            public void AddTestRow(long rawBits, String expectedExcelRendering)
            {
                WriteDataRow(_sheet, _rowIndex++, rawBits, expectedExcelRendering);
                if (Double.IsNaN(BitConverter.Int64BitsToDouble(rawBits)))
                {
                    _ReplacementNaNs.Add(rawBits);
                }
            }

            public long[] GetReplacementNaNs()
            {
                int nRepls = _ReplacementNaNs.Count;
                long[] result = new long[nRepls];
                for (int i = 0; i < nRepls; i++)
                {
                    result[i] = _ReplacementNaNs[i];
                }
                return result;
            }

        }
        /** 0x7ff8000000000000 encoded in little endian order */
        private static byte[] JAVA_NAN_BYTES = HexRead.ReadFromString("00 00 00 00 00 00 F8 7F");

        private static void WriteHeaderCell(IRow row, int i, String text, ICellStyle style)
        {
            ICell cell = row.CreateCell(i);
            cell.SetCellValue(new HSSFRichTextString(text));
            cell.CellStyle = (style);
        }
        static void WriteHeaderRow(IWorkbook wb, ISheet sheet)
        {
            sheet.SetColumnWidth(0, 3000);
            sheet.SetColumnWidth(1, 6000);
            sheet.SetColumnWidth(2, 6000);
            sheet.SetColumnWidth(3, 6000);
            sheet.SetColumnWidth(4, 6000);
            sheet.SetColumnWidth(5, 1600);
            sheet.SetColumnWidth(6, 20000);
            IRow row = sheet.CreateRow(0);
            ICellStyle style = wb.CreateCellStyle();
            IFont font = wb.CreateFont();
            WriteHeaderCell(row, 0, "Value", style);
            font.IsBold = true;
            style.SetFont(font);
            WriteHeaderCell(row, 1, "Raw Long Bits", style);
            WriteHeaderCell(row, 2, "JDK Double Rendering", style);
            WriteHeaderCell(row, 3, "Actual Rendering", style);
            WriteHeaderCell(row, 4, "Expected Rendering", style);
            WriteHeaderCell(row, 5, "Match", style);
            WriteHeaderCell(row, 6, "Java Metadata", style);
        }
        static void WriteDataRow(ISheet sheet, int rowIx, long rawLongBits, String expectedExcelRendering)
        {
            double d = BitConverter.Int64BitsToDouble(rawLongBits);
            IRow row = sheet.CreateRow(rowIx);

            int rowNum = rowIx + 1;
            String cel0ref = "A" + rowNum;
            String rawBitsText = FormatLongAsHex(rawLongBits);
            String jmExpr = "'ec(" + rawBitsText + ", ''\" & C" + rowNum + " & \"'', ''\" & D" + rowNum + " & \"''),'";

            // The 'Match' column will contain 'OK' if the metadata (from NumberToTextConversionExamples)
            // matches Excel's rendering.
            String matchExpr = "if(D" + rowNum + "=E" + rowNum + ", \"OK\", \"ERROR\")";

            row.CreateCell(0).SetCellValue(d);
            row.CreateCell(1).SetCellValue(new HSSFRichTextString(rawBitsText));
            row.CreateCell(2).SetCellValue(new HSSFRichTextString(d.ToString()));
            row.CreateCell(3).CellFormula = ("\"\" & " + cel0ref);
            row.CreateCell(4).SetCellValue(new HSSFRichTextString(expectedExcelRendering));
            row.CreateCell(5).CellFormula = (matchExpr);
            row.CreateCell(6).CellFormula = (jmExpr.Replace("'", "\""));

            //if (false)
            //{
            //    // for observing arithmetic near numeric range boundaries
            //    row.CreateCell(7).CellFormula=(cel0ref + " * 1.0001");
            //    row.CreateCell(8).CellFormula=(cel0ref + " / 1.0001");
            //}
        }

        private static String FormatLongAsHex(long l)
        {
            StringBuilder sb = new StringBuilder(20);
            sb.Append(HexDump.LongToHex(l)).Append('L');
            return sb.ToString();
        }

        //public static void Main(String[] args)
        //{
        //    WriteJavaDoc();

        //    HSSFWorkbook wb = new HSSFWorkbook();
        //    SheetWriter sw = new SheetWriter(wb);

        //    NumberToTextConversionExamples.ExampleConversion[] exampleValues = NumberToTextConversionExamples.GetExampleConversions();
        //    for (int i = 0; i < exampleValues.Length; i++)
        //    {
        //        TestCases.SS.Util.NumberToTextConversionExamples.ExampleConversion example = exampleValues[i];
        //        sw.AddTestRow(example.RawDoubleBits, example.ExcelRendering);
        //    }

        //    MemoryStream baos = new MemoryStream();
        //    wb.Write(baos);
        //    byte[] fileContent = baos.ToArray();
        //    ReplaceNaNs(fileContent, sw.GetReplacementNaNs());


        //    FileInfo outputFile = new FileInfo("ExcelNumberRendering.xls");

        //    FileStream os = File.OpenWrite(outputFile.FullName);
        //    os.Write(fileContent, 0, fileContent.Length);
        //    os.Close();
        //    Console.WriteLine("Finished writing '" + outputFile.FullName + "'");
        //}

        public static void WriteJavaDoc()
        {

            NumberToTextConversionExamples.ExampleConversion[] exampleConversions = NumberToTextConversionExamples.GetExampleConversions();
            for (int i = 0; i < exampleConversions.Length; i++)
            {
                NumberToTextConversionExamples.ExampleConversion ec = exampleConversions[i];
                String line = " * <tr><td>"
                    + FormatLongAsHex(ec.RawDoubleBits)
                    + "</td><td>" + ec.DoubleValue.ToString()
                    + "</td><td>" + ec.ExcelRendering + "</td></tr>";

                Console.WriteLine(line);
            }
        }



        private static void ReplaceNaNs(byte[] fileContent, long[] ReplacementNaNs)
        {
            int countFound = 0;
            for (int i = 0; i < fileContent.Length; i++)
            {
                if (IsNaNBytes(fileContent, i))
                {
                    WriteLong(fileContent, i, ReplacementNaNs[countFound]);
                    countFound++;
                }
            }
            if (countFound < ReplacementNaNs.Length)
            {
                throw new Exception("wrong repl count");
            }

        }

        private static void WriteLong(byte[] bb, int i, long val)
        {
            String oldVal = InterpretLong(bb, i);
            bb[i + 7] = (byte)(val >> 56);
            bb[i + 6] = (byte)(val >> 48);
            bb[i + 5] = (byte)(val >> 40);
            bb[i + 4] = (byte)(val >> 32);
            bb[i + 3] = (byte)(val >> 24);
            bb[i + 2] = (byte)(val >> 16);
            bb[i + 1] = (byte)(val >> 8);
            bb[i + 0] = (byte)(val >> 0);
            //if (false)
            //{
            //    String newVal = interpretLong(bb, i);
            //    Console.WriteLine("Changed offset " + i + " from " + oldVal + " to " + newVal);
            //}

        }

        private static String InterpretLong(byte[] fileContent, int offset)
        {
            Stream is1 = new MemoryStream(fileContent, offset, 8);
            long l;
            l = LittleEndian.ReadLong(is1);
            return "0x" + StringUtil.ToHexString(l).ToUpper();
        }

        private static bool IsNaNBytes(byte[] fileContent, int offset)
        {
            if (offset + JAVA_NAN_BYTES.Length > fileContent.Length)
            {
                return false;
            }
            // excel NaN bits: 0xFFFF0420003C0000L
            // java NaN bits :0x7ff8000000000000L
            return AreArraySectionsEqual(fileContent, offset, JAVA_NAN_BYTES);
        }
        private static bool AreArraySectionsEqual(byte[] bb, int off, byte[] section)
        {
            for (int i = section.Length - 1; i >= 0; i--)
            {
                if (bb[off + i] != section[i])
                {
                    return false;
                }
            }
            return true;
        }
    }
}
