using NPOI.SS.UserModel;
using NPOI.SS.UserModel.Charts;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace LineChart
{
    class Program
    {
        const int NUM_OF_ROWS = 3;
           const int NUM_OF_COLUMNS = 10;

          static void CreateChart(IDrawing drawing, ISheet sheet, IClientAnchor anchor, string serie1, string serie2)
        {
            IChart chart = drawing.CreateChart(anchor);
            IChartLegend legend = chart.GetOrCreateLegend();
            legend.Position = LegendPosition.TopRight;

            ILineChartData<double, double> data = chart.ChartDataFactory.CreateLineChartData<double, double>();

            // Use a category axis for the bottom axis.
            IChartAxis bottomAxis = chart.GetChartAxisFactory().CreateCategoryAxis(AxisPosition.Bottom);
            IValueAxis leftAxis = chart.GetChartAxisFactory().CreateValueAxis(AxisPosition.Left);
            leftAxis.SetCrosses(AxisCrosses.AutoZero);

            IChartDataSource<double> xs = DataSources.FromNumericCellRange(sheet, new CellRangeAddress(0, 0, 0, NUM_OF_COLUMNS - 1));
            IChartDataSource<double> ys1 = DataSources.FromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, NUM_OF_COLUMNS - 1));
            IChartDataSource<double> ys2 = DataSources.FromNumericCellRange(sheet, new CellRangeAddress(2, 2, 0, NUM_OF_COLUMNS - 1));

            var s1 = data.AddSeries(xs, ys1);
            s1.SetTitle(serie1);
            var s2 = data.AddSeries(xs, ys2);
            s2.SetTitle(serie2);

            chart.Plot(data, bottomAxis, leftAxis);
        }

        static void Main(string[] args)
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("linechart");


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
            IClientAnchor anchor1 = drawing.CreateAnchor(0, 0, 0, 0, 0, 5, 10, 15);
            CreateChart(drawing, sheet, anchor1, "title1","title2");
            IClientAnchor anchor2 = drawing.CreateAnchor(0, 0, 0, 0, 0, 20, 10, 35);
            CreateChart(drawing, sheet, anchor2, "s1", "s2");
            using (FileStream fs =File.Create("test.xlsx"))
            {
                wb.Write(fs);
            }
        }
    }
}
