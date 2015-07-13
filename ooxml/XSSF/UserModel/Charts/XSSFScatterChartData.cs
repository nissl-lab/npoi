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

using System.Collections.Generic;
using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.SS.Util;
using NPOI.SS.UserModel.Charts;
using NPOI.XSSF.UserModel.Charts;
using NPOI.SS.UserModel;
using System;
namespace NPOI.XSSF.UserModel.Charts
{

    /**
     * Represents DrawingML scatter chart.
     *
     * @author Roman Kashitsyn
     */
    public class XSSFScatterChartData<Tx, Ty> : IScatterChartData<Tx, Ty>
    {

        /**
         * List of all data series.
         */
        private List<IScatterChartSeries<Tx, Ty>> series;

        public XSSFScatterChartData()
        {
            series = new List<IScatterChartSeries<Tx, Ty>>();
        }
        /**
         * Package private ScatterChartSerie implementation.
         */
        internal class Series : AbstractXSSFChartSeries, IScatterChartSeries<Tx, Ty>
        {
            private int id;
            private int order;
            private bool useCache;
            //private DataMarker xMarker;
            //private DataMarker yMarker;
            //private XSSFNumberCache lastCaclulatedXCache;
            //private XSSFNumberCache lastCalculatedYCache;

            private IChartDataSource<Tx> xs;
            private IChartDataSource<Ty> ys;

            internal Series(int id, int order,
                IChartDataSource<Tx> xs, IChartDataSource<Ty> ys)
                : base()
            {

                this.id = id;
                this.order = order;
                //this.useCache = true;
                this.xs = xs;
                this.ys = ys;
            }

            /**
             * Returns data source used for X axis values.
             * @return data source used for X axis values
             */
            public IChartDataSource<Tx> GetXValues()
            {
                return xs;
            }

            /**
             * Returns data source used for Y axis values.
             * @return data source used for Y axis values
             */
            public IChartDataSource<Ty> GetYValues()
            {
                return ys;
            }

            /**
             * @param useCache if true, cached results will be Added on plot
             */
            public void SetUseCache(bool useCache)
            {
                this.useCache = useCache;
            }


            internal void AddToChart(CT_ScatterChart ctScatterChart)
            {
                CT_ScatterSer scatterSer = ctScatterChart.AddNewSer();
                scatterSer.AddNewIdx().val = (uint)this.id;
                scatterSer.AddNewOrder().val = (uint)this.order;

                CT_AxDataSource xVal = scatterSer.AddNewXVal();
                XSSFChartUtil.BuildAxDataSource<Tx>(xVal, xs);
                CT_NumDataSource yVal = scatterSer.AddNewYVal();
                XSSFChartUtil.BuildNumDataSource<Ty>(yVal, ys);

                if (IsTitleSet)
                {
                    scatterSer.tx = this.GetCTSerTx();
                }
            }
        }

        public IScatterChartSeries<Tx,Ty> AddSeries(IChartDataSource<Tx> xs, IChartDataSource<Ty> ys)
        {
            if (!ys.IsNumeric)
            {
                throw new ArgumentException("Y axis data source must be numeric.");
            }
            int numOfSeries = series.Count;
            Series newSerie = new Series(numOfSeries, numOfSeries, xs, ys);
            series.Add(newSerie);
            return newSerie;
        }

        public void FillChart(IChart chart, IChartAxis[] axis)
        {
            if (!(chart is XSSFChart))
            {
                throw new ArgumentException("Chart must be instance of XSSFChart");
            }

            XSSFChart xssfChart = (XSSFChart)chart;
            CT_PlotArea plotArea = xssfChart.GetCTChart().plotArea;
            CT_ScatterChart scatterChart = plotArea.AddNewScatterChart();
            AddStyle(scatterChart);

            foreach (Series s in series)
            {
                s.AddToChart(scatterChart);
            }

            foreach (IChartAxis ax in axis)
            {
                scatterChart.AddNewAxId().val = (uint)ax.Id;
            }
        }

        public List<IScatterChartSeries<Tx, Ty>> GetSeries()
        {
            return series;
        }

        private void AddStyle(CT_ScatterChart ctScatterChart)
        {
            CT_ScatterStyle scatterStyle = ctScatterChart.AddNewScatterStyle();
            scatterStyle.val = ST_ScatterStyle.lineMarker;
        }
    }

}