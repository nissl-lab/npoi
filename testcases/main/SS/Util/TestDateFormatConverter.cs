/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

using System;
using System.Globalization;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NUnit.Framework;

namespace TestCases.SS.Util
{
    [TestFixture]
    public class TestDateFormatConverter
    {

        private void OutputLocaleDataFormats(DateTime date, bool dates, bool times, int style, String styleName)
        {

            IWorkbook workbook = new HSSFWorkbook();
            try
            {
                String sheetName;
                if (dates)
                {
                    if (times)
                    {
                        sheetName = "DateTimes";
                    }
                    else
                    {
                        sheetName = "Dates";
                    }
                }
                else
                {
                    sheetName = "Times";
                }
                ISheet sheet = workbook.CreateSheet(sheetName);
                IRow header = sheet.CreateRow(0);
                header.CreateCell(0).SetCellValue("locale");
                header.CreateCell(1).SetCellValue("DisplayName");
                header.CreateCell(2).SetCellValue("Excel " + styleName);
                header.CreateCell(3).SetCellValue("java.text.DateFormat");
                header.CreateCell(4).SetCellValue("Equals");
                header.CreateCell(5).SetCellValue("Java pattern");
                header.CreateCell(6).SetCellValue("Excel pattern");

                int rowNum = 1;
                foreach (CultureInfo locale in CultureInfo.GetCultures(CultureTypes.AllCultures))
                {
                    if (string.IsNullOrEmpty(locale.ToString()))
                        continue;
                    IRow row = sheet.CreateRow(rowNum++);

                    row.CreateCell(0).SetCellValue(locale.ToString());
                    row.CreateCell(1).SetCellValue(locale.DisplayName);

                    string csharpDateFormatPattern;
                    if (dates)
                    {
                        if (times)
                        {
                            csharpDateFormatPattern = DateFormat.GetDateTimePattern(style, style, locale);
                        }
                        else
                        {
                            csharpDateFormatPattern = DateFormat.GetDatePattern(style, locale);
                        }
                    }
                    else
                    {
                        csharpDateFormatPattern = DateFormat.GetTimePattern(style, locale);
                    }

                    //Excel Date Value
                    ICell cell = row.CreateCell(2);

                    cell.SetCellValue(date);
                    ICellStyle cellStyle = row.Sheet.Workbook.CreateCellStyle();

                    //String csharpDateFormatPattern = locale.DateTimeFormat.LongDatePattern;
                    String excelFormatPattern = DateFormatConverter.Convert(locale, csharpDateFormatPattern);

                    IDataFormat poiFormat = row.Sheet.Workbook.CreateDataFormat();
                    cellStyle.DataFormat = (poiFormat.GetFormat(excelFormatPattern));
                    cell.CellStyle = (cellStyle);

                    //C# Date value
                    row.CreateCell(3).SetCellValue(date.ToString(csharpDateFormatPattern, locale.DateTimeFormat));



                    // the formula returns TRUE is the formatted date in column C equals to the string in column D
                    row.CreateCell(4).SetCellFormula("TEXT(C" + rowNum + ",G" + rowNum + ")=D" + rowNum);
                    //C# pattern
                    row.CreateCell(5).SetCellValue(csharpDateFormatPattern);
                    //excel pattern
                    row.CreateCell(6).SetCellValue(excelFormatPattern);
                }

                //FileInfo outputFile = TempFile.CreateTempFile("Locale" + sheetName + styleName, ".xlsx");
                string filename = "Locale" + sheetName + styleName + ".xls";
                FileStream outputStream = new FileStream(filename, FileMode.Create);
                try
                {
                    workbook.Write(outputStream);
                }
                finally
                {
                    outputStream.Close();
                }
                System.Console.WriteLine("Open " + filename + " in Excel");
            }
            finally
            {
                workbook.Close();
            }
            

        }
        [Test]
        public void TestJavaDateFormatsInExcel()
        {

            DateTime date = DateTime.Now;
            
            OutputLocaleDataFormats(date, true, false, DateFormat.DEFAULT, "Default");
            OutputLocaleDataFormats(date, true, false, DateFormat.SHORT, "Short");
            OutputLocaleDataFormats(date, true, false, DateFormat.MEDIUM, "Medium");
            OutputLocaleDataFormats(date, true, false, DateFormat.LONG, "Long");
            OutputLocaleDataFormats(date, true, false, DateFormat.FULL, "Full");

            OutputLocaleDataFormats(date, true, true, DateFormat.DEFAULT, "Default");
            OutputLocaleDataFormats(date, true, true, DateFormat.SHORT, "Short");
            OutputLocaleDataFormats(date, true, true, DateFormat.MEDIUM, "Medium");
            OutputLocaleDataFormats(date, true, true, DateFormat.LONG, "Long");
            OutputLocaleDataFormats(date, true, true, DateFormat.FULL, "Full");

            OutputLocaleDataFormats(date, false, true, DateFormat.DEFAULT, "Default");
            OutputLocaleDataFormats(date, false, true, DateFormat.SHORT, "Short");
            OutputLocaleDataFormats(date, false, true, DateFormat.MEDIUM, "Medium");
            OutputLocaleDataFormats(date, false, true, DateFormat.LONG, "Long");
            OutputLocaleDataFormats(date, false, true, DateFormat.FULL, "Full");
        }

    }

}