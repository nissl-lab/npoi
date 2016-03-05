using NPOI.SS.UserModel;
using NPOI.SS.UserModel.Charts;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ScatterChart
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet 1");
            int NUM_OF_ROWS = 3;
            int NUM_OF_COLUMNS = 10;

            // Create a row and put some cells in it. Rows are 0 based.
            IRow row;
            ICell cell;
            for (int rowIndex = 0; rowIndex < NUM_OF_ROWS; rowIndex++)
            {
                row = sheet.CreateRow((short)rowIndex);
                for (int colIndex = 0; colIndex < NUM_OF_COLUMNS; colIndex++)
                {
                    cell = row.CreateCell((short)colIndex);
                    cell.SetCellValue(colIndex * (rowIndex + 1));
                }
            }

            IDrawing drawing = sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = drawing.CreateAnchor(0, 0, 0, 0, 0, 5, 10, 15);

            IChart chart = drawing.CreateChart(anchor);
            IChartLegend legend = chart.GetOrCreateLegend();
            legend.Position = (LegendPosition.TopRight);

            IScatterChartData<double, double> data = chart.ChartDataFactory.CreateScatterChartData<double, double>();

            IValueAxis bottomAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Bottom);
            IValueAxis leftAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);
            leftAxis.Crosses = AxisCrosses.AutoZero;

            IChartDataSource<double> xs = DataSources.FromNumericCellRange(sheet, new CellRangeAddress(0, 0, 0, NUM_OF_COLUMNS - 1));
            IChartDataSource<double> ys1 = DataSources.FromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, NUM_OF_COLUMNS - 1));
            IChartDataSource<double> ys2 = DataSources.FromNumericCellRange(sheet, new CellRangeAddress(2, 2, 0, NUM_OF_COLUMNS - 1));


            data.AddSeries(xs, ys1);
            data.AddSeries(xs, ys2);
            chart.Plot(data, bottomAxis, leftAxis);

            // Write the output to a file
            FileStream sw = File.Create("test.xlsx");
            wb.Write(sw);
            sw.Close();
        }
    }
}
