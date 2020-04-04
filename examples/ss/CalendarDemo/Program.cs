using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;
using System.IO;

namespace CalendarDemo
{
    class Program
    {
        private static String[] days = {
            "Sunday", "Monday", "Tuesday",
            "Wednesday", "Thursday", "Friday", "Saturday"};

        private static String[] months = {
            "January", "February", "March","April", "May", "June","July", "August",
            "September","October", "November", "December"};
        static void Main(string[] args)
        {
            DateTime dt = DateTime.Now;
            bool xlsx = false;
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i][0] == '-')
                {
                    xlsx = args[i].Equals("-xlsx");
                }
                else
                {
                    dt = new DateTime(dt.Year, int.Parse(args[i]), dt.Day);
                }
            }
            int year = dt.Year;

            IWorkbook wb = xlsx ? new XSSFWorkbook() as IWorkbook : new HSSFWorkbook() as IWorkbook;

            Dictionary<String, ICellStyle> styles = createStyles(wb);
            DateTime dtM;
            for (int month = 0; month < 12; month++)
            {
                dtM = new DateTime(dt.Year, month + 1, 1);
                //calendar.set(Calendar.MONTH, month);
                //calendar.set(Calendar.DAY_OF_MONTH, 1);
                //create a sheet for each month
                ISheet sheet = wb.CreateSheet(months[month]);

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
                headerRow.HeightInPoints = (80);
                ICell titleCell = headerRow.CreateCell(0);
                titleCell.SetCellValue(months[month] + " " + year);
                titleCell.CellStyle = (styles[("title")]);
                sheet.AddMergedRegion(CellRangeAddress.ValueOf("$A$1:$N$1"));

                //header with month titles
                IRow monthRow = sheet.CreateRow(1);
                for (int i = 0; i < days.Length; i++)
                {
                    //set column widths, the width is measured in units of 1/256th of a character width
                    sheet.SetColumnWidth(i * 2, 5 * 256); //the column is 5 characters wide
                    sheet.SetColumnWidth(i * 2 + 1, 13 * 256); //the column is 13 characters wide
                    sheet.AddMergedRegion(new CellRangeAddress(1, 1, i * 2, i * 2 + 1));
                    ICell monthCell = monthRow.CreateCell(i * 2);
                    monthCell.SetCellValue(days[i]);
                    monthCell.CellStyle = (styles["month"]);
                }

                int cnt = 1, day = 1;
                int rownum = 2;
                for (int j = 0; j < 6; j++)
                {
                    IRow row = sheet.CreateRow(rownum++);
                    row.HeightInPoints = (100);
                    for (int i = 0; i < days.Length; i++)
                    {
                        ICell dayCell_1 = row.CreateCell(i * 2);
                        ICell dayCell_2 = row.CreateCell(i * 2 + 1);

                        int day_of_week = (int)dtM.DayOfWeek;
                        if (cnt >= day_of_week && dtM.Month == (month+1))
                        {
                            dayCell_1.SetCellValue(day);
                            //calendar.set(Calendar.DAY_OF_MONTH, ++day);
                            dtM.AddDays(++day);

                            if (i == 0 || i == days.Length - 1)
                            {
                                dayCell_1.CellStyle = (styles["weekend_left"]);
                                dayCell_2.CellStyle = (styles["weekend_right"]);
                            }
                            else
                            {
                                dayCell_1.CellStyle = (styles["workday_left"]);
                                dayCell_2.CellStyle = (styles["workday_right"]);
                            }
                        }
                        else
                        {
                            dayCell_1.CellStyle = (styles["grey_left"]);
                            dayCell_2.CellStyle = (styles["grey_right"]);
                        }
                        cnt++;
                    }
                    if (dtM.Month > (month+1)) break;
                }
            }

            // Write the output to a file
            String file = "calendar.xls";
            if (wb is XSSFWorkbook) file += "x";
            FileStream out1 = new FileStream(file, FileMode.Create);
            wb.Write(out1);
            out1.Close();
        }

        /**
     * cell styles used for formatting calendar sheets
     */
        private static Dictionary<String, ICellStyle> createStyles(IWorkbook wb)
        {
            Dictionary<String, ICellStyle> styles = new Dictionary<String, ICellStyle>();

            short borderColor = IndexedColors.Grey50Percent.Index;

            ICellStyle style;
            IFont titleFont = wb.CreateFont();
            titleFont.FontHeightInPoints = ((short)48);
            titleFont.Color = (IndexedColors.DarkBlue.Index);
            style = wb.CreateCellStyle();
            style.Alignment = (HorizontalAlignment.Center);
            style.VerticalAlignment = VerticalAlignment.Center;
            style.SetFont(titleFont);
            styles.Add("title", style);

            IFont monthFont = wb.CreateFont();
            monthFont.FontHeightInPoints = ((short)12);
            monthFont.Color = (IndexedColors.White.Index);
            monthFont.Boldweight = (short)(FontBoldWeight.Bold);
            style = wb.CreateCellStyle();
            style.Alignment = (HorizontalAlignment.Center);
            style.VerticalAlignment = (VerticalAlignment.Center);
            style.FillForegroundColor = (IndexedColors.DarkBlue.Index);
            style.FillPattern = (FillPattern.SolidForeground);
            style.SetFont(monthFont);
            styles.Add("month", style);

            IFont dayFont = wb.CreateFont();
            dayFont.FontHeightInPoints = ((short)14);
            dayFont.Boldweight = (short)(FontBoldWeight.Bold);
            style = wb.CreateCellStyle();
            style.Alignment = (HorizontalAlignment.Left);
            style.VerticalAlignment = (VerticalAlignment.Top);
            style.FillForegroundColor = (IndexedColors.LightCornflowerBlue.Index);
            style.FillPattern = (FillPattern.SolidForeground);
            style.BorderLeft = (BorderStyle.Thin);
            style.LeftBorderColor = (borderColor);
            style.BorderBottom = (BorderStyle.Thin);
            style.BottomBorderColor = (borderColor);
            style.SetFont(dayFont);
            styles.Add("weekend_left", style);

            style = wb.CreateCellStyle();
            style.Alignment = (HorizontalAlignment.Center);
            style.VerticalAlignment = (VerticalAlignment.Top);
            style.FillForegroundColor = (IndexedColors.LightCornflowerBlue.Index);
            style.FillPattern = (FillPattern.SolidForeground);
            style.BorderRight = (BorderStyle.Thin);
            style.RightBorderColor = (borderColor);
            style.BorderBottom = (BorderStyle.Thin);
            style.BottomBorderColor = (borderColor);
            styles.Add("weekend_right", style);

            style = wb.CreateCellStyle();
            style.Alignment = (HorizontalAlignment.Left);
            style.VerticalAlignment = (VerticalAlignment.Top);
            style.BorderLeft = (BorderStyle.Thin);
            style.FillForegroundColor = (IndexedColors.White.Index);
            style.FillPattern = (FillPattern.SolidForeground);
            style.LeftBorderColor = (borderColor);
            style.BorderBottom = (BorderStyle.Thin);
            style.BottomBorderColor = (borderColor);
            style.SetFont(dayFont);
            styles.Add("workday_left", style);

            style = wb.CreateCellStyle();
            style.Alignment = (HorizontalAlignment.Center);
            style.VerticalAlignment = (VerticalAlignment.Top);
            style.FillForegroundColor = (IndexedColors.White.Index);
            style.FillPattern = (FillPattern.SolidForeground);
            style.BorderRight = (BorderStyle.Thin);
            style.RightBorderColor = (borderColor);
            style.BorderBottom = (BorderStyle.Thin);
            style.BottomBorderColor = (borderColor);
            styles.Add("workday_right", style);

            style = wb.CreateCellStyle();
            style.BorderLeft = (BorderStyle.Thin);
            style.FillForegroundColor = (IndexedColors.Grey25Percent.Index);
            style.FillPattern = (FillPattern.SolidForeground);
            style.BorderBottom = (BorderStyle.Thin);
            style.BottomBorderColor = (borderColor);
            styles.Add("grey_left", style);

            style = wb.CreateCellStyle();
            style.FillForegroundColor = (IndexedColors.Grey25Percent.Index);
            style.FillPattern = (FillPattern.SolidForeground);
            style.BorderRight = (BorderStyle.Thin);
            style.RightBorderColor = (borderColor);
            style.BorderBottom = (BorderStyle.Thin);
            style.BottomBorderColor = (borderColor);
            styles.Add("grey_right", style);

            return styles;
        }
    }
}
