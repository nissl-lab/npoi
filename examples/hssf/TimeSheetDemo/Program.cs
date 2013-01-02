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

using System;
using System.Collections.Generic;
using System.Text;

using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace TimeSheetDemo
{
    /// <summary>
    /// converted from http://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
    /// </summary>
    class Program
    {
        static String[] titles = {
            "Person",	"ID", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun",
            "Total\nHrs", "Overtime\nHrs", "Regular\nHrs"
        };
        static Object[,] sample_data = {
            {"Yegor Kozlov", "YK", 5.0, 8.0, 10.0, 5.0, 5.0, 7.0, 6.0},
            {"Gisella Bronzetti", "GB", 4.0, 3.0, 1.0, 3.5, null, null, 4.0}
        };


        static void Main(string[] args)
        {
            InitializeWorkbook();

            Dictionary<String, ICellStyle> styles = CreateStyles(hssfworkbook);

            ISheet sheet = hssfworkbook.CreateSheet("Timesheet");
            IPrintSetup printSetup = sheet.PrintSetup;
            printSetup.Landscape = true;
            sheet.FitToPage=(true);
            sheet.HorizontallyCenter=(true);

            //title row
            IRow titleRow = sheet.CreateRow(0);
            titleRow.HeightInPoints=(45);
            ICell titleCell = titleRow.CreateCell(0);
            titleCell.SetCellValue("Weekly Timesheet");
            titleCell.CellStyle= (styles["title"]);
            sheet.AddMergedRegion(CellRangeAddress.ValueOf("$A$1:$L$1"));

            //header row
            IRow headerRow = sheet.CreateRow(1);
            headerRow.HeightInPoints = (40);
            ICell headerCell;
            for (int i = 0; i < titles.Length; i++)
            {
                headerCell = headerRow.CreateCell(i);
                headerCell.SetCellValue(titles[i]);
                headerCell.CellStyle = (styles["header"]);
            }


            int rownum = 2;
            for (int i = 0; i < 10; i++)
            {
                IRow row = sheet.CreateRow(rownum++);
                for (int j = 0; j < titles.Length; j++)
                {
                    ICell cell = row.CreateCell(j);
                    if (j == 9)
                    {
                        //the 10th cell contains sum over week days, e.g. SUM(C3:I3)
                        String reference = "C" + rownum + ":I" + rownum;
                        cell.CellFormula = ("SUM(" + reference + ")");
                        cell.CellStyle = (styles["formula"]);
                    }
                    else if (j == 11)
                    {
                        cell.CellFormula = ("J" + rownum + "-K" + rownum);
                        cell.CellStyle = (styles["formula"]);
                    }
                    else
                    {
                        cell.CellStyle = (styles["cell"]);
                    }
                }
            }

            //row with totals below
            IRow sumRow = sheet.CreateRow(rownum++);
            sumRow.HeightInPoints = (35);
            ICell cell1 = sumRow.CreateCell(0);
            cell1.CellStyle = (styles["formula"]);

            ICell cell2 = sumRow.CreateCell(1);
            cell2.SetCellValue("Total Hrs:");
            cell2.CellStyle=(styles["formula"]);

            for (int j = 2; j < 12; j++) {
                ICell cell = sumRow.CreateCell(j);
                String reference = (char)('A' + j) + "3:" + (char)('A' + j) + "12";
                cell.CellFormula = ("SUM(" + reference + ")");
                if(j >= 9)
                    cell.CellStyle = (styles["formula_2"]);
                else
                    cell.CellStyle = (styles["formula"]);
            }

            rownum++;
            sumRow = sheet.CreateRow(rownum++);
            sumRow.HeightInPoints = 25;
            ICell cell3 = sumRow.CreateCell(0);
            cell3.SetCellValue("Total Regular Hours");
            cell3.CellStyle = styles["formula"];
            cell3 = sumRow.CreateCell(1);
            cell3.CellFormula = ("L13");
            cell3.CellStyle=styles["formula_2"];
            sumRow = sheet.CreateRow(rownum++);
            sumRow.HeightInPoints = (25);
            cell3 = sumRow.CreateCell(0);
            cell3.SetCellValue("Total Overtime Hours");
            cell3.CellStyle = styles["formula"];
            cell3 = sumRow.CreateCell(1);
            cell3.CellFormula = ("K13");
            cell3.CellStyle = styles["formula_2"];

                    //set sample data
            for (int i = 0; i < sample_data.GetLength(0); i++)
            {
                IRow row = sheet.GetRow(2 + i);
                for (int j = 0; j < sample_data.GetLength(1); j++)
                {
                    if (sample_data[i,j] == null)
                        continue;

                    if (sample_data[i,j] is String)
                    {
                        row.GetCell(j).SetCellValue((String)sample_data[i,j]);
                    }
                    else
                    {
                        row.GetCell(j).SetCellValue((Double)sample_data[i,j]);
                    }
                }
            }

                    //finally set column widths, the width is measured in units of 1/256th of a character width
        sheet.SetColumnWidth(0, 30*256); //30 characters wide
        for (int i = 2; i < 9; i++) {
            sheet.SetColumnWidth(i, 6*256);  //6 characters wide
        }
        sheet.SetColumnWidth(10, 10*256); //10 characters wide


            WriteToFile();
        }

        /**
   * Create a library of cell styles
   */
        private static Dictionary<String, ICellStyle> CreateStyles(IWorkbook wb)
        {
            Dictionary<String, ICellStyle> styles = new Dictionary<String, ICellStyle>();
            ICellStyle style;
            IFont titleFont = wb.CreateFont();
            titleFont.FontHeightInPoints = ((short)18);
            titleFont.Boldweight = (short)FontBoldWeight.BOLD;
            style = wb.CreateCellStyle();
            style.Alignment = HorizontalAlignment.CENTER;
            style.VerticalAlignment = VerticalAlignment.CENTER;
            style.SetFont(titleFont);
            styles.Add("title", style);

            IFont monthFont = wb.CreateFont();
            monthFont.FontHeightInPoints = ((short)11);
            monthFont.Color = (IndexedColors.WHITE.Index);
            style = wb.CreateCellStyle();
            style.Alignment = HorizontalAlignment.CENTER;
            style.VerticalAlignment = VerticalAlignment.CENTER;
            style.FillForegroundColor = (IndexedColors.GREY_50_PERCENT.Index);
            style.FillPattern = FillPatternType.SOLID_FOREGROUND;
            style.SetFont(monthFont);
            style.WrapText = (true);
            styles.Add("header", style);

            style = wb.CreateCellStyle();
            style.Alignment = HorizontalAlignment.CENTER;
            style.WrapText = (true);
            style.BorderRight = BorderStyle.THIN;
            style.RightBorderColor = (IndexedColors.BLACK.Index);
            style.BorderLeft = BorderStyle.THIN;
            style.LeftBorderColor = (IndexedColors.BLACK.Index);
            style.BorderTop = BorderStyle.THIN;
            style.TopBorderColor = (IndexedColors.BLACK.Index);
            style.BorderBottom = BorderStyle.THIN;
            style.BottomBorderColor = (IndexedColors.BLACK.Index);
            styles.Add("cell", style);

            style = wb.CreateCellStyle();
            style.Alignment = HorizontalAlignment.CENTER;
            style.VerticalAlignment = VerticalAlignment.CENTER;
            style.FillForegroundColor = (IndexedColors.GREY_25_PERCENT.Index);
            style.FillPattern = FillPatternType.SOLID_FOREGROUND;
            style.DataFormat = (wb.CreateDataFormat().GetFormat("0.00"));
            styles.Add("formula", style);

            style = wb.CreateCellStyle();
            style.Alignment = HorizontalAlignment.CENTER;
            style.VerticalAlignment = VerticalAlignment.CENTER;
            style.FillForegroundColor = (IndexedColors.GREY_40_PERCENT.Index);
            style.FillPattern = FillPatternType.SOLID_FOREGROUND;
            style.DataFormat = wb.CreateDataFormat().GetFormat("0.00");
            styles.Add("formula_2", style);

            return styles;
        }


        static HSSFWorkbook hssfworkbook;

        static void WriteToFile()
        {
            // Write the output to a file
            String filename = "timesheet.xls";
            //Write the stream data of workbook to the root directory
            FileStream file = new FileStream(filename, FileMode.Create);
            hssfworkbook.Write(file);
            file.Close();
        }

        static void InitializeWorkbook()
        {
            hssfworkbook = new HSSFWorkbook();

            ////create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            ////create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;
        }
    }
}
