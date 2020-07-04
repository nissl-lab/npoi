using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.IO;

namespace SetBordersOfRegion
{
    class Program
    {
        static HSSFWorkbook hssfworkbook;

        static void Main(string[] args)
        {
            InitializeWorkbook();

            ISheet sheet1 = hssfworkbook.CreateSheet("Sheet1");
            //create a common style
            ICellStyle blackBorder = hssfworkbook.CreateCellStyle();
            blackBorder.BorderBottom = BorderStyle.Thin;
            blackBorder.BorderLeft = BorderStyle.Thin;
            blackBorder.BorderRight = BorderStyle.Thin;
            blackBorder.BorderTop = BorderStyle.Thin;
            blackBorder.BottomBorderColor = HSSFColor.Black.Index;
            blackBorder.LeftBorderColor = HSSFColor.Black.Index;
            blackBorder.RightBorderColor = HSSFColor.Black.Index;
            blackBorder.TopBorderColor = HSSFColor.Black.Index;

            //create horizontal 1-9
            for (int i = 1; i <= 9; i++)
            {
                sheet1.CreateRow(0).CreateCell(i).SetCellValue(i);
            }
            //create vertical 1-9
            for (int i = 1; i <= 9; i++)
            {
                sheet1.CreateRow(i).CreateCell(0).SetCellValue(i);
            }
            //create the cell formula
            for (int iRow = 1; iRow <= 9; iRow++)
            {
                IRow row = sheet1.GetRow(iRow);
                for (int iCol = 1; iCol <= 9; iCol++)
                {
                    //the first cell of each row * the first cell of each column
                    string formula = GetCellPosition(iRow, 0) + "*" + GetCellPosition(0, iCol);
                    ICell cell = row.CreateCell(iCol);
                    cell.CellFormula = formula;
                    //set the cellstyle to the cell
                    cell.CellStyle = blackBorder;
                }
            }

            WriteToFile();
        }

        static string GetCellPosition(int row, int col)
        {
            col = Convert.ToInt32('A') + col;
            row = row + 1;
            return ((char)col) + row.ToString();
        }

        static void WriteToFile()
        {
            //Write the stream data of workbook to the root directory
            FileStream file = new FileStream(@"test.xls", FileMode.Create);
            hssfworkbook.Write(file);
            file.Close();
        }

        static void InitializeWorkbook()
        {
            hssfworkbook = new HSSFWorkbook();
        }
    }
}
