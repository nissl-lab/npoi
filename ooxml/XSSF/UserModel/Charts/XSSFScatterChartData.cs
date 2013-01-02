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
        private List<IScatterChartSerie<Tx, Ty>> series;

        public XSSFScatterChartData()
        {
            series = new List<IScatterChartSerie<Tx, Ty>>();
        }

        /**
         * Package private ScatterChartSerie implementation.
         */
        public class Serie : IScatterChartSerie<Tx, Ty>
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

            internal Serie(int id, int order, 
                IChartDataSource<Tx> xs, IChartDataSource<Ty> ys)
                : base()
            {

                this.id = id;
                this.order = order;
                //this.useCache = true;
                this.xs = xs;
                this.ys = ys;
            }

            /*public void SetXValues(DataMarker marker)
            {
                xMarker = marker;
            }

            public void SetYValues(DataMarker marker)
            {
                yMarker = marker;
            }*/
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

            /// <summary>
            /// Returns last calculated number cache for X axis.
            /// </summary>
            //internal XSSFNumberCache LastCaculatedXCache
            //{
            //    get
            //    {
            //        return lastCaclulatedXCache;
            //    }
            //}
            /// <summary>
            /// Returns last calculated number cache for Y axis.
            /// </summary>
            //internal XSSFNumberCache LastCalculatedYCache
            //{
            //    get
            //    {
            //        return lastCalculatedYCache;
            //    }
            //}

            internal void AddToChart(CT_ScatterChart ctScatterChart)
            {
                CT_ScatterSer scatterSer = ctScatterChart.AddNewSer();
                scatterSer.AddNewIdx().val= (uint)this.id;
                scatterSer.AddNewOrder().val = (uint)this.order;

                /* TODO: add some logic to automatically recognize cell
                 * types and choose appropriate data representation for
                 * X axis.
                 */
                /*CT_AxDataSource xVal = scatterSer.AddNewXVal();
                CT_NumRef xNumRef = xVal.AddNewNumRef();
                xNumRef.f = (xMarker.FormatAsString());

                CT_NumDataSource yVal = scatterSer.AddNewYVal();
                CT_NumRef yNumRef = yVal.AddNewNumRef();
                yNumRef.f = (yMarker.FormatAsString());

                if (useCache)
                {
                    // We can not store cache since markers are not immutable
                    XSSFNumberCache.BuildCache(xMarker, xNumRef);
                    lastCalculatedYCache = XSSFNumberCache.BuildCache(yMarker, yNumRef);
                }
                */
                CT_AxDataSource xVal = scatterSer.AddNewXVal();
                XSSFChartUtil.BuildAxDataSource<Tx>(xVal, xs);
                CT_NumDataSource yVal = scatterSer.AddNewYVal();
                XSSFChartUtil.BuildNumDataSource<Ty>(yVal, ys);
            }
        }

        /*public IScatterChartSerie AddSerie(DataMarker xMarker, DataMarker yMarker)
        {
            int numOfSeries = series.Count;
            Serie newSerie = new Serie(numOfSeries, numOfSeries);
            newSerie.SetXValues(xMarker);
            newSerie.SetYValues(yMarker);
            series.Add(newSerie);
            return newSerie;
        }*/
        public IScatterChartSerie<Tx,Ty> AddSerie(IChartDataSource<Tx> xs, IChartDataSource<Ty> ys)
        {
            if (!ys.IsNumeric)
            {
                throw new ArgumentException("Y axis data source must be numeric.");
            }
            int numOfSeries = series.Count;
            Serie newSerie = new Serie(numOfSeries, numOfSeries, xs, ys);
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

            foreach (Serie s in series)
            {
                s.AddToChart(scatterChart);
            }

            foreach (IChartAxis ax in axis)
            {
                scatterChart.AddNewAxId().val = (uint)ax.GetId();
            }
        }

        public List<IScatterChartSerie<Tx, Ty>> GetSeries()
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