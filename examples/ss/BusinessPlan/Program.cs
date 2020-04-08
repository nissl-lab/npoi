using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace BusinessPlan
{
    class Program
    {
        private static SimpleDateFormat fmt = new SimpleDateFormat("dd-MMM");

        private static String[] titles = {
            "ID", "Project Name", "Owner", "Days", "Start", "End"};

        //sample data to fill the sheet.
        private static String[][] data = new string[18][];
        /*{
            {"1.0", "Marketing Research Tactical Plan", "J. Dow", "70", "9-Jul", null,
                "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"},
            null,
            {"1.1", "Scope Definition Phase", "J. Dow", "10", "9-Jul", null,
                "x", "x", null, null,  null, null, null, null, null, null, null},
            {"1.1.1", "Define research objectives", "J. Dow", "3", "9-Jul", null,
                    "x", null, null, null,  null, null, null, null, null, null, null},
            {"1.1.2", "Define research requirements", "S. Jones", "7", "10-Jul", null,
                "x", "x", null, null,  null, null, null, null, null, null, null},
            {"1.1.3", "Determine in-house resource or hire vendor", "J. Dow", "2", "15-Jul", null,
                "x", "x", null, null,  null, null, null, null, null, null, null},
            null,
            {"1.2", "Vendor Selection Phase", "J. Dow", "19", "19-Jul", null,
                null, "x", "x", "x",  "x", null, null, null, null, null, null},
            {"1.2.1", "Define vendor selection criteria", "J. Dow", "3", "19-Jul", null,
                null, "x", null, null,  null, null, null, null, null, null, null},
            {"1.2.2", "Develop vendor selection questionnaire", "S. Jones, T. Wates", "2", "22-Jul", null,
                null, "x", "x", null,  null, null, null, null, null, null, null},
            {"1.2.3", "Develop Statement of Work", "S. Jones", "4", "26-Jul", null,
                null, null, "x", "x",  null, null, null, null, null, null, null},
            {"1.2.4", "Evaluate proposal", "J. Dow, S. Jones", "4", "2-Aug", null,
                null, null, null, "x",  "x", null, null, null, null, null, null},
            {"1.2.5", "Select vendor", "J. Dow", "1", "6-Aug", null,
                null, null, null, null,  "x", null, null, null, null, null, null},
            null,
            {"1.3", "Research Phase", "G. Lee", "47", "9-Aug", null,
                null, null, null, null,  "x", "x", "x", "x", "x", "x", "x"},
            {"1.3.1", "Develop market research information needs questionnaire", "G. Lee", "2", "9-Aug", null,
                null, null, null, null,  "x", null, null, null, null, null, null},
            {"1.3.2", "Interview marketing group for market research needs", "G. Lee", "2", "11-Aug", null,
                null, null, null, null,  "x", "x", null, null, null, null, null},
            {"1.3.3", "Document information needs", "G. Lee, S. Jones", "1", "13-Aug", null,
                null, null, null, null,  null, "x", null, null, null, null, null},
    };*/
        static void Main(string[] args)
        {
            data[0] = new string[] {"1.0", "Marketing Research Tactical Plan", "J. Dow", "70", "9-Jul", null,
                "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"};
            data[1] = new string[] { null };
            data[2] = new string[] {"1.1", "Scope Definition Phase", "J. Dow", "10", "9-Jul", null,
                "x", "x", null, null,  null, null, null, null, null, null, null};
            data[3] = new string[] {"1.1.1", "Define research objectives", "J. Dow", "3", "9-Jul", null,
                    "x", null, null, null,  null, null, null, null, null, null, null};
            data[4] = new string[] {"1.1.2", "Define research requirements", "S. Jones", "7", "10-Jul", null,
                "x", "x", null, null,  null, null, null, null, null, null, null};
            data[5] = new string[] {"1.1.3", "Determine in-house resource or hire vendor", "J. Dow", "2", "15-Jul", null,
                "x", "x", null, null,  null, null, null, null, null, null, null};
            data[6] = new string[] { null };
            data[7] = new string[] {"1.2", "Vendor Selection Phase", "J. Dow", "19", "19-Jul", null,
                null, "x", "x", "x",  "x", null, null, null, null, null, null};
            data[8] = new string[] {"1.2.1", "Define vendor selection criteria", "J. Dow", "3", "19-Jul", null,
                null, "x", null, null,  null, null, null, null, null, null, null};
            data[9] = new string[] {"1.2.2", "Develop vendor selection questionnaire", "S. Jones, T. Wates", "2", "22-Jul", null,
                null, "x", "x", null,  null, null, null, null, null, null, null};
            data[10] = new string[] {"1.2.3", "Develop Statement of Work", "S. Jones", "4", "26-Jul", null,
                null, null, "x", "x",  null, null, null, null, null, null, null};
            data[11] = new string[] {"1.2.4", "Evaluate proposal", "J. Dow, S. Jones", "4", "2-Aug", null,
                null, null, null, "x",  "x", null, null, null, null, null, null};
            data[12] = new string[] {"1.2.5", "Select vendor", "J. Dow", "1", "6-Aug", null,
                null, null, null, null,  "x", null, null, null, null, null, null};
            data[13] = new string[] { null };
            data[14] = new string[] {"1.3", "Research Phase", "G. Lee", "47", "9-Aug", null,
                null, null, null, null,  "x", "x", "x", "x", "x", "x", "x"};
            data[15] = new string[] {"1.3.1", "Develop market research information needs questionnaire", "G. Lee", "2", "9-Aug", null,
                null, null, null, null,  "x", null, null, null, null, null, null};
            data[16] = new string[] {"1.3.2", "Interview marketing group for market research needs", "G. Lee", "2", "11-Aug", null,
                null, null, null, null,  "x", "x", null, null, null, null, null};
            data[17] = new string[] {"1.3.3", "Document information needs", "G. Lee, S. Jones", "1", "13-Aug", null,
                null, null, null, null,  null, "x", null, null, null, null, null};

            IWorkbook wb;

            if (args.Length > 0 && args[0].Equals("-xls"))
                wb = new HSSFWorkbook();
            else
                wb = new XSSFWorkbook();

            Dictionary<String, ICellStyle> styles = createStyles(wb);

            ISheet sheet = wb.CreateSheet("Business Plan");

            //turn off gridlines
            sheet.DisplayGridlines = (false);
            sheet.IsPrintGridlines = (false);
            sheet.FitToPage = (true);
            sheet.HorizontallyCenter = (true);
            IPrintSetup printSetup = sheet.PrintSetup;
            printSetup.Landscape = (true);

            //the following three statements are required only for HSSF
            sheet.Autobreaks = (true);
            printSetup.FitHeight = ((short)1);
            printSetup.FitWidth = ((short)1);

            //the header row: centered text in 48pt font
            IRow headerRow = sheet.CreateRow(0);
            headerRow.HeightInPoints = (12.75f);
            for (int i = 0; i < titles.Length; i++)
            {
                ICell cell = headerRow.CreateCell(i);
                cell.SetCellValue(titles[i]);
                cell.CellStyle = (styles["header"]);
            }
            //columns for 11 weeks starting from 9-Jul
            DateTime dt = new DateTime(DateTime.Now.Year, 6, 9);
            for (int i = 0; i < 11; i++)
            {
                ICell cell = headerRow.CreateCell(titles.Length + i);
                cell.SetCellValue(dt);
                cell.CellStyle = (styles[("header_date")]);
                //calendar.roll(Calendar.WEEK_OF_YEAR, true);
                dt.AddDays(7);
            }
            //freeze the first row
            sheet.CreateFreezePane(0, 1);

            IRow row;
            //ICell cell;
            int rownum = 1;
            for (int i = 0; i < data.Length; i++, rownum++)
            {
                row = sheet.CreateRow(rownum);
                if (data[i] == null) continue;

                for (int j = 0; j < data[i].Length; j++)
                {
                    ICell cell = row.CreateCell(j);
                    String styleName;
                    bool isHeader = i == 0 || data[i - 1] == null;
                    switch (j)
                    {
                        case 0:
                            if (isHeader)
                            {
                                styleName = "cell_b";
                                cell.SetCellValue(Double.Parse(data[i][j]));
                            }
                            else
                            {
                                styleName = "cell_normal";
                                cell.SetCellValue(data[i][j]);
                            }
                            break;
                        case 1:
                            if (isHeader)
                            {
                                styleName = i == 0 ? "cell_h" : "cell_bb";
                            }
                            else
                            {
                                styleName = "cell_indented";
                            }
                            cell.SetCellValue(data[i][j]);
                            break;
                        case 2:
                            styleName = isHeader ? "cell_b" : "cell_normal";
                            cell.SetCellValue(data[i][j]);
                            break;
                        case 3:
                            styleName = isHeader ? "cell_b_centered" : "cell_normal_centered";
                            cell.SetCellValue(int.Parse(data[i][j]));
                            break;
                        case 4:
                            {
                                //calendar.setTime(fmt.parse(data[i][j]));
                                //calendar.set(Calendar.YEAR, year);

                                DateTime dt2 = DateTime.Parse(DateTime.Now.Year.ToString() + "-" + data[i][j]);

                                cell.SetCellValue(dt2);
                                styleName = isHeader ? "cell_b_date" : "cell_normal_date";
                                break;
                            }
                        case 5:
                            {
                                int r = rownum + 1;
                                String fmla = "IF(AND(D" + r + ",E" + r + "),E" + r + "+D" + r + ",\"\")";
                                cell.SetCellFormula(fmla);
                                styleName = isHeader ? "cell_bg" : "cell_g";
                                break;
                            }
                        default:
                            styleName = data[i][j] != null ? "cell_blue" : "cell_normal";
                            break;
                    }

                    cell.CellStyle = (styles[(styleName)]);
                }
            }

            //group rows for each phase, row numbers are 0-based
            sheet.GroupRow(4, 6);
            sheet.GroupRow(9, 13);
            sheet.GroupRow(16, 18);

            //set column widths, the width is measured in units of 1/256th of a character width
            sheet.SetColumnWidth(0, 256 * 6);
            sheet.SetColumnWidth(1, 256 * 33);
            sheet.SetColumnWidth(2, 256 * 20);
            sheet.SetZoom(75);


            // Write the output to a file
            String file = "businessplan.xls";
            if (wb is XSSFWorkbook) file += "x";
            using (FileStream out1 = new FileStream(file, FileMode.Create))
            {
                wb.Write(out1);
                out1.Close();
            }

        }

        /**
     * create a library of cell styles
     */
        private static Dictionary<String, ICellStyle> createStyles(IWorkbook wb)
        {
            Dictionary<String, ICellStyle> styles = new Dictionary<String, ICellStyle>();
            IDataFormat df = wb.CreateDataFormat();

            ICellStyle style;
            IFont headerFont = wb.CreateFont();
            headerFont.IsBold = true;
            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.FillForegroundColor = (IndexedColors.LightCornflowerBlue.Index);
            style.FillPattern = FillPattern.SolidForeground;
            style.SetFont(headerFont);
            styles.Add("header", style);

            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.FillForegroundColor = (IndexedColors.LightCornflowerBlue.Index);
            style.FillPattern = FillPattern.SolidForeground;
            style.SetFont(headerFont);
            style.DataFormat = (df.GetFormat("d-mmm"));
            styles.Add("header_date", style);

            IFont font1 = wb.CreateFont();
            font1.IsBold = true;
            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.SetFont(font1);
            styles.Add("cell_b", style);

            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.SetFont(font1);
            styles.Add("cell_b_centered", style);

            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.SetFont(font1);
            style.DataFormat = (df.GetFormat("d-mmm"));
            styles.Add("cell_b_date", style);

            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.SetFont(font1);
            style.FillForegroundColor = (IndexedColors.Grey25Percent.Index);
            style.FillPattern = FillPattern.SolidForeground;
            style.DataFormat = (df.GetFormat("d-mmm"));
            styles.Add("cell_g", style);

            IFont font2 = wb.CreateFont();
            font2.Color = (IndexedColors.Blue.Index);
            font2.IsBold = true;
            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.SetFont(font2);
            styles.Add("cell_bb", style);

            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.SetFont(font1);
            style.FillForegroundColor = (IndexedColors.Grey25Percent.Index);
            style.FillPattern = FillPattern.SolidForeground;
            style.DataFormat = (df.GetFormat("d-mmm"));
            styles.Add("cell_bg", style);

            IFont font3 = wb.CreateFont();
            font3.FontHeightInPoints = ((short)14);
            font3.Color = (IndexedColors.DarkBlue.Index);
            font3.IsBold = true;
            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.SetFont(font3);
            style.WrapText = (true);
            styles.Add("cell_h", style);

            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.WrapText = (true);
            styles.Add("cell_normal", style);

            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.WrapText = (true);
            styles.Add("cell_normal_centered", style);

            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.WrapText = (true);
            style.DataFormat = (df.GetFormat("d-mmm"));
            styles.Add("cell_normal_date", style);

            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.Indention = ((short)1);
            style.WrapText = (true);
            styles.Add("cell_indented", style);

            style = CreateBorderedStyle(wb);
            style.FillForegroundColor = (IndexedColors.Blue.Index);
            style.FillPattern = FillPattern.SolidForeground;
            styles.Add("cell_blue", style);

            return styles;
        }

        private static ICellStyle CreateBorderedStyle(IWorkbook wb)
        {
            ICellStyle style = wb.CreateCellStyle();
            style.BorderRight = BorderStyle.Thin;
            style.RightBorderColor = (IndexedColors.Black.Index);
            style.BorderBottom = BorderStyle.Thin;
            style.BottomBorderColor = (IndexedColors.Black.Index);
            style.BorderLeft = BorderStyle.Thin;
            style.LeftBorderColor = (IndexedColors.Black.Index);
            style.BorderTop = BorderStyle.Thin;
            style.TopBorderColor = (IndexedColors.Black.Index);
            return style;
        }
    }
}
