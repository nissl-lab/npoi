using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace NPOI.Examples.XSSF.DataFormatsInXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook workbook = new XSSFWorkbook();

            ISheet sheet = workbook.CreateSheet("Sheet1");
            //increase the width of Column A
            sheet.SetColumnWidth(0, 5000);
            //create the format instance
            IDataFormat format = workbook.CreateDataFormat();

            // Create a row and put some cells in it. Rows are 0 based.
            ICell cell = sheet.CreateRow(0).CreateCell(0);
            //number format with 2 digits after the decimal point - "1.20"
            SetValueAndFormat(workbook, cell, 1.2, HSSFDataFormat.GetBuiltinFormat("0.00"));

            //RMB currency format with comma    -   "¥20,000"
            ICell cell2 = sheet.CreateRow(1).CreateCell(0);
            SetValueAndFormat(workbook, cell2, 20000, format.GetFormat("¥#,##0"));

            //scentific number format   -   "3.15E+00"
            ICell cell3 = sheet.CreateRow(2).CreateCell(0);
            SetValueAndFormat(workbook, cell3, 3.151234, format.GetFormat("0.00E+00"));

            //percent format, 2 digits after the decimal point    -  "99.33%"
            ICell cell4 = sheet.CreateRow(3).CreateCell(0);
            SetValueAndFormat(workbook, cell4, 0.99333, format.GetFormat("0.00%"));

            //phone number format - "021-65881234"
            ICell cell5 = sheet.CreateRow(4).CreateCell(0);
            SetValueAndFormat(workbook, cell5, 02165881234, format.GetFormat("000-00000000"));

            //Chinese capitalized character number - 壹贰叁 元
            ICell cell6 = sheet.CreateRow(5).CreateCell(0);
            SetValueAndFormat(workbook, cell6, 123, format.GetFormat("[DbNum2][$-804]0 元"));

            //Chinese date string
            ICell cell7 = sheet.CreateRow(6).CreateCell(0);
            SetValueAndFormat(workbook, cell7, new DateTime(2004, 5, 6), format.GetFormat("yyyy年m月d日"));
            cell7.SetCellValue(new DateTime(2004, 5, 6));

            //Chinese date string
            ICell cell8 = sheet.CreateRow(7).CreateCell(0);
            SetValueAndFormat(workbook, cell8, new DateTime(2005, 11, 6), format.GetFormat("yyyy年m月d日"));

            //formula value with datetime style 
            ICell cell9 = sheet.CreateRow(8).CreateCell(0);
            cell9.CellFormula = "DateValue(\"2005-11-11\")+TIMEVALUE(\"11:11:11\")";
            ICellStyle cellStyle9 = workbook.CreateCellStyle();
            cellStyle9.DataFormat = HSSFDataFormat.GetBuiltinFormat("m/d/yy h:mm");
            cell9.CellStyle = cellStyle9;

            //display current time
            ICell cell10 = sheet.CreateRow(9).CreateCell(0);
            SetValueAndFormat(workbook, cell10, DateTime.Now, format.GetFormat("[$-409]h:mm:ss AM/PM;@"));

            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
        static void SetValueAndFormat(IWorkbook workbook, ICell cell, int value, short formatId)
        {
            cell.SetCellValue(value);
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.DataFormat = formatId;
            cell.CellStyle = cellStyle;
        }

        static void SetValueAndFormat(IWorkbook workbook, ICell cell, double value, short formatId)
        {
            cell.SetCellValue(value);
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.DataFormat = formatId;
            cell.CellStyle = cellStyle;
        }
        static void SetValueAndFormat(IWorkbook workbook, ICell cell, DateTime value, short formatId)
        {
            //set value for the cell
            if (value != null)
                cell.SetCellValue(value);

            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.DataFormat = formatId;
            cell.CellStyle = cellStyle;
        }
    }
}
