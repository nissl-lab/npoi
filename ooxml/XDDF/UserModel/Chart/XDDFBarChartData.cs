/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.XDDF.UserModel.Chart
{
    using NPOI.XDDF.UserModel;
    using NPOI.OpenXmlFormats.Dml.Chart;

    public class XDDFBarChartData<T, V> : XDDFChartData<T, V>
    {
        private CT_BarChart chart;

        public XDDFBarChartData(CT_BarChart chart, Dictionary<long, XDDFChartAxis> categories,
                Dictionary<long, XDDFValueAxis> values)
        {
            this.chart = chart;
            if(chart.barDir == null)
            {
                chart.AddNewBarDir().val = BarDirection.Bar.ToST_BarDir();
            }
            foreach(CT_BarSer series in chart.ser)
            {
                this.series.Add(new Series(series, series.cat, series.val));
            }
            DefineAxes(categories, values);
        }

        private void DefineAxes(Dictionary<long, XDDFChartAxis> categories, 
            Dictionary<long, XDDFValueAxis> values)
        {
            if(chart.SizeOfAxIdArray() == 0)
            {
                foreach(long id in categories.Keys)
                {
                    chart.AddNewAxId().val = (uint) id;
                }
                foreach(long id in values.Keys)
                {
                    chart.AddNewAxId().val = (uint) id;
                }
            }
            DefineAxes(chart.GetAxIdArray(), categories, values);
        }
        public override void SetVaryColors(bool varyColors)
        {
            if(chart.IsSetVaryColors())
            {
                chart.varyColors.val = varyColors ? 1 : 0;
            }
            else
            {
                chart.AddNewVaryColors().val = varyColors ? 1 : 0;
            }
        }

        public BarDirection GetBarDirection()
        {
            return BarDirectionExtensions.ValueOf(chart.barDir.val);
        }

        public void SetBarDirection(BarDirection direction)
        {
            chart.barDir.val = direction.ToST_BarDir();
        }

        public BarGrouping GetBarGrouping()
        {
            if(chart.IsSetGrouping())
            {
                return BarGroupingExtensions.ValueOf(chart.grouping.val);
            }
            else
            {
                return BarGrouping.Standard;
            }
        }

        public void SetBarGrouping(BarGrouping grouping)
        {
            if(chart.IsSetGrouping())
            {
                chart.grouping.val = grouping.ToST_BarGrouping();
            }
            else
            {
                chart.AddNewGrouping().val = grouping.ToST_BarGrouping();
            }
        }

        public int GetGapWidth()
        {
            if(chart.IsSetGapWidth())
            {
                return chart.gapWidth.val;
            }
            else
            {
                return 0;
            }
        }

        public void SetGapWidth(int width)
        {
            if(chart.IsSetGapWidth())
            {
                chart.gapWidth.val = (ushort) width;
            }
            else
            {
                chart.AddNewGapWidth().val = (ushort) width;
            }
        }
        public override XDDFChartData<T, V>.Series AddSeries(IXDDFDataSource<T> category,
                IXDDFNumericalDataSource<V> values)
        {
            int index = this.series.Count;
            CT_BarSer ctSer = this.chart.AddNewSer();
            ctSer.AddNewTx();
            ctSer.AddNewCat();
            ctSer.AddNewVal();
            ctSer.AddNewIdx().val = (uint) index;
            ctSer.AddNewOrder().val = (uint) index;
            Series added = new Series(ctSer, category, values);
            this.series.Add(added);
            return added;
        }

        public class Series : XDDFChartData<T, V>.Series
        {
            private CT_BarSer series;

            internal Series(CT_BarSer series, IXDDFDataSource<T> category,
                    IXDDFNumericalDataSource<V> values)
                    : base(category, values)
            {

                this.series = series;
            }

            internal Series(CT_BarSer series, CT_AxDataSource category, CT_NumDataSource values)
                    : base(XDDFDataSourcesFactory.FromDataSource(category) as IXDDFDataSource<T>,
                          XDDFDataSourcesFactory.FromDataSource(values) as IXDDFNumericalDataSource<V>)
            {

                this.series = series;
            }
            protected override CT_SerTx GetSeriesText()
            {
                if(series.IsSetTx())
                {
                    return series.tx;
                }
                else
                {
                    return series.AddNewTx();
                }
            }
            public override void SetShowLeaderLines(bool showLeaderLines)
            {
                if(!series.IsSetDLbls())
                {
                    series.AddNewDLbls();
                }
                if(series.dLbls.IsSetShowLeaderLines())
                {
                    series.dLbls.showLeaderLines.val = showLeaderLines ? 1 : 0;
                }
                else
                {
                    series.dLbls.AddNewShowLeaderLines().val = showLeaderLines ? 1 : 0;
                }
            }
            public override XDDFShapeProperties GetShapeProperties()
            {
                if(series.IsSetSpPr())
                {
                    return new XDDFShapeProperties(series.spPr);
                }
                else
                {
                    return null;
                }
            }
            public override void SetShapeProperties(XDDFShapeProperties properties)
            {
                if(properties == null)
                {
                    if(series.IsSetSpPr())
                    {
                        series.UnsetSpPr();
                    }
                }
                else
                {
                    if(series.IsSetSpPr())
                    {
                        series.spPr = properties.GetXmlObject();
                    }
                    else
                    {
                        series.AddNewSpPr().Set(properties.GetXmlObject());
                    }
                }
            }
            protected override CT_AxDataSource GetAxDS()
            {
                return series.cat;
            }
            protected override CT_NumDataSource GetNumDS()
            {
                return series.val;
            }
        }
    }
}
