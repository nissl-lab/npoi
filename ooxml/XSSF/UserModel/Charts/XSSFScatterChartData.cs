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
    public class XSSFScatterChartData : IScatterChartData
    {

        /**
         * List of all data series.
         */
        private List<IScatterChartSerie> series;

        public XSSFScatterChartData()
        {
            series = new List<IScatterChartSerie>();
        }

        /**
         * Package private ScatterChartSerie implementation.
         */
        public class Serie : IScatterChartSerie
        {
            private int id;
            private int order;
            private bool useCache;
            private DataMarker xMarker;
            private DataMarker yMarker;
            private XSSFNumberCache lastCaclulatedXCache;
            private XSSFNumberCache lastCalculatedYCache;

            internal Serie(int id, int order)
                : base()
            {

                this.id = id;
                this.order = order;
                this.useCache = true;
            }

            public void SetXValues(DataMarker marker)
            {
                xMarker = marker;
            }

            public void SetYValues(DataMarker marker)
            {
                yMarker = marker;
            }

            /**
             * @param useCache if true, cached results will be Added on plot
             */
            public void SetUseCache(bool useCache)
            {
                this.useCache = useCache;
            }

            /**
             * Returns last calculated number cache for X axis.
             * @return last calculated number cache for X axis.
             */
            XSSFNumberCache GetLastCaculatedXCache()
            {
                return lastCaclulatedXCache;
            }

            /**
             * Returns last calculated number cache for Y axis.
             * @return last calculated number cache for Y axis.
             */
            XSSFNumberCache GetLastCalculatedYCache()
            {
                return lastCalculatedYCache;
            }

            internal void AddToChart(CT_ScatterChart ctScatterChart)
            {
                CT_ScatterSer scatterSer = ctScatterChart.AddNewSer();
                scatterSer.AddNewIdx().val= (uint)this.id;
                scatterSer.AddNewOrder().val = (uint)this.order;

                /* TODO: add some logic to automatically recognize cell
                 * types and choose appropriate data representation for
                 * X axis.
                 */
                CT_AxDataSource xVal = scatterSer.AddNewXVal();
                CT_NumRef xNumRef = xVal.AddNewNumRef();
                xNumRef.f = (xMarker.FormatAsString());

                CT_NumDataSource yVal = scatterSer.AddNewYVal();
                CT_NumRef yNumRef = yVal.AddNewNumRef();
                yNumRef.f = (yMarker.FormatAsString());

                if (useCache)
                {
                    /* We can not store cache since markers are not immutable */
                    XSSFNumberCache.BuildCache(xMarker, xNumRef);
                    lastCalculatedYCache = XSSFNumberCache.BuildCache(yMarker, yNumRef);
                }
            }
        }

        public IScatterChartSerie AddSerie(DataMarker xMarker, DataMarker yMarker)
        {
            int numOfSeries = series.Count;
            Serie newSerie = new Serie(numOfSeries, numOfSeries);
            newSerie.SetXValues(xMarker);
            newSerie.SetYValues(yMarker);
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

        public List<IScatterChartSerie> GetSeries()
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