using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.SS.UserModel.Charts;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.XSSF.UserModel.Charts
{
    /// <summary>
    /// Holds data for a XSSF Line Chart
    /// </summary>
    /// <typeparam name="Tx"></typeparam>
    /// <typeparam name="Ty"></typeparam>
    public class XSSFLineChartData<Tx, Ty> : ILineChartData<Tx,Ty>
    {
        /**
         * List of all data series.
         */
        private List<ILineChartSeries<Tx, Ty>> series;

        public XSSFLineChartData()
        {
            series = new List<ILineChartSeries<Tx, Ty>>();
        }

        public class Series:AbstractXSSFChartSeries, ILineChartSeries<Tx,Ty> 
        {
            private int id;
            private int order;
            private IChartDataSource<Tx> categories;
            private IChartDataSource<Ty> values;

            internal Series(int id, int order,
                            IChartDataSource<Tx> categories,
                            IChartDataSource<Ty> values) 
            {
                this.id = id;
                this.order = order;
                this.categories = categories;
                this.values = values;
            }

            public IChartDataSource<Tx> GetCategoryAxisData() {
                return categories;
            }

            public IChartDataSource<Ty> GetValues() {
                return values;
            }

            internal void AddToChart(CT_LineChart ctLineChart) {
                CT_LineSer ctLineSer = ctLineChart.AddNewSer();
                CT_Grouping ctGrouping = ctLineChart.AddNewGrouping();
                ctGrouping.val = ST_Grouping.standard;
                ctLineSer.AddNewIdx().val= (uint)id;
                ctLineSer.AddNewOrder().val = (uint)order;

                // No marker symbol on the chart line.
                ctLineSer.AddNewMarker().AddNewSymbol().val = ST_MarkerStyle.none;

                CT_AxDataSource catDS = ctLineSer.AddNewCat();
                XSSFChartUtil.BuildAxDataSource(catDS, categories);
                CT_NumDataSource valueDS = ctLineSer.AddNewVal();
                XSSFChartUtil.BuildNumDataSource(valueDS, values);

                if (IsTitleSet) {
                    ctLineSer.tx = GetCTSerTx();
                }
            }
        }
        public ILineChartSeries<Tx, Ty> AddSeries(IChartDataSource<Tx> categoryAxisData, IChartDataSource<Ty> values)
        {
            if (!values.IsNumeric)
            {
                throw new ArgumentException("Value data source must be numeric.");
            }
            int numOfSeries = series.Count;
            Series newSeries = new Series(numOfSeries, numOfSeries, categoryAxisData, values);
            series.Add(newSeries);
            return newSeries;
        }

        public List<ILineChartSeries<Tx, Ty>> GetSeries()
        {
            return series;
        }

        public void FillChart(SS.UserModel.IChart chart, params IChartAxis[] axis)
        {
            if (!(chart is XSSFChart))
            {
                throw new ArgumentException("Chart must be instance of XSSFChart");
            }

            XSSFChart xssfChart = (XSSFChart)chart;
            CT_PlotArea plotArea = xssfChart.GetCTChart().plotArea;
            CT_LineChart lineChart = plotArea.AddNewLineChart();
            lineChart.AddNewVaryColors().val = 0;

            foreach (Series s in series)
            {
                s.AddToChart(lineChart);
            }

            foreach (IChartAxis ax in axis)
            {
                lineChart.AddNewAxId().val = (uint)ax.Id;
            }
        }
    }
}
