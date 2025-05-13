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
    using NPOI.SS.Util;
    using NPOI.XDDF.UserModel;
    using NPOI.OpenXmlFormats.Dml.Chart;
    /// <summary>
    /// Base of all XDDF Chart Data
    /// </summary>
    public abstract class XDDFChartData<T, V>
    {
        protected List<Series> series;
        private XDDFCategoryAxis categoryAxis;
        private List<XDDFValueAxis> valueAxes;

        protected XDDFChartData()
        {
            this.series = [];
        }

        protected void DefineAxes(CT_UnsignedInt[] axes, Dictionary<long, XDDFChartAxis> categories,
                Dictionary<long, XDDFValueAxis> values)
        {
            List<XDDFValueAxis> list = new List<XDDFValueAxis>(axes.Length);
            foreach(CT_UnsignedInt axe in axes)
            {
                long axisId = axe.val;
                XDDFChartAxis category = categories[axisId];
                if(category == null)
                {
                    XDDFValueAxis axis = values[axisId];
                    if(axis != null)
                    {
                        list.Add(axis);
                    }
                }
                else if(category is XDDFCategoryAxis)
                {
                    this.categoryAxis = (XDDFCategoryAxis) category;
                }
            }
            this.valueAxes = list;
        }

        public XDDFCategoryAxis GetCategoryAxis()
        {
            return categoryAxis;
        }

        public List<XDDFValueAxis> GetValueAxes()
        {
            return valueAxes;
        }

        public List<Series> GetSeries()
        {
            return series;
        }

        public abstract void SetVaryColors(bool varyColors);

        public abstract Series AddSeries(IXDDFDataSource<T> category,
                IXDDFNumericalDataSource<V> values);

        public abstract class Series
        {
            protected abstract CT_SerTx GetSeriesText();

            public abstract void SetShowLeaderLines(bool showLeaderLines);
            public abstract XDDFShapeProperties GetShapeProperties();
            public abstract void SetShapeProperties(XDDFShapeProperties properties);

            protected IXDDFDataSource<T> categoryData;
            protected IXDDFNumericalDataSource<V> valuesData;

            protected abstract CT_AxDataSource GetAxDS();

            protected abstract CT_NumDataSource GetNumDS();

            protected Series(IXDDFDataSource<T> category, IXDDFNumericalDataSource<V> values)
            {
                ReplaceData(category, values);
            }

            public void ReplaceData(IXDDFDataSource<T> category, IXDDFNumericalDataSource<V> values)
            {
                if(category == null || values == null)
                {
                    throw new InvalidOperationException("Category and values must be defined before filling chart data.");
                }
                int numOfPoints = category.PointCount;
                if(numOfPoints != values.PointCount)
                {
                    throw new InvalidOperationException("Category and values must have the same point count.");
                }
                this.categoryData = category;
                this.valuesData = values;
            }

            public void SetTitle(string title, CellReference titleRef)
            {
                if(titleRef == null)
                {
                    GetSeriesText().v = title;
                }
                else
                {
                    CT_StrRef ref1;
                    if(GetSeriesText().IsSetStrRef())
                    {
                        ref1 = GetSeriesText().strRef;
                    }
                    else
                    {
                        ref1 = GetSeriesText().AddNewStrRef();
                    }
                    CT_StrData cache;
                    if(ref1.IsSetStrCache())
                    {
                        cache = ref1.strCache;
                    }
                    else
                    {
                        cache = ref1.AddNewStrCache();
                    }
                    if(cache.SizeOfPtArray() < 1)
                    {
                        cache.AddNewPtCount().val = 1;
                        cache.AddNewPt().idx = 0;
                    }
                    cache.GetPtArray(0).v = title;
                    ref1.f = titleRef.FormatAsString();
                }
            }

            public IXDDFDataSource<T> GetCategoryData()
            {
                return categoryData;
            }

            public IXDDFNumericalDataSource<V> GetValuesData()
            {
                return valuesData;
            }

            public void Plot()
            {
                int numOfPoints = categoryData.PointCount;
                if(categoryData.IsNumeric)
                {
                    CT_NumData cache1 = RetrieveNumCache(GetAxDS(), categoryData);
                    FillNumCache(cache1, numOfPoints, (IXDDFNumericalDataSource<V>) categoryData);
                }
                else
                {
                    CT_StrData cache2 = RetrieveStrCache(GetAxDS(), categoryData);
                    FillStringCache(cache2, numOfPoints, categoryData);
                }
                CT_NumData cache = RetrieveNumCache(GetNumDS(), valuesData);
                FillNumCache(cache, numOfPoints, valuesData);
            }

            private CT_NumData RetrieveNumCache(CT_AxDataSource axDataSource, 
                IXDDFDataSource<T> data)
            {
                CT_NumData numCache;
                if(data.IsReference)
                {
                    CT_NumRef numRef;
                    if(axDataSource.IsSetNumRef())
                    {
                        numRef = axDataSource.numRef;
                    }
                    else
                    {
                        numRef = axDataSource.AddNewNumRef();
                    }
                    if(numRef.IsSetNumCache())
                    {
                        numCache = numRef.numCache;
                    }
                    else
                    {
                        numCache = numRef.AddNewNumCache();
                    }
                    numRef.f = data.DataRangeReference;
                    if(axDataSource.IsSetNumLit())
                    {
                        axDataSource.UnsetNumLit();
                    }
                }
                else
                {
                    if(axDataSource.IsSetNumLit())
                    {
                        numCache = axDataSource.numLit;
                    }
                    else
                    {
                        numCache = axDataSource.AddNewNumLit();
                    }
                    if(axDataSource.IsSetNumRef())
                    {
                        axDataSource.UnsetNumRef();
                    }
                }
                return numCache;
            }

            private CT_StrData RetrieveStrCache(CT_AxDataSource axDataSource, IXDDFDataSource<T> data)
            {
                CT_StrData strCache;
                if(data.IsReference)
                {
                    CT_StrRef strRef;
                    if(axDataSource.IsSetStrRef())
                    {
                        strRef = axDataSource.strRef;
                    }
                    else
                    {
                        strRef = axDataSource.AddNewStrRef();
                    }
                    if(strRef.IsSetStrCache())
                    {
                        strCache = strRef.strCache;
                    }
                    else
                    {
                        strCache = strRef.AddNewStrCache();
                    }
                    strRef.f = data.DataRangeReference;
                    if(axDataSource.IsSetStrLit())
                    {
                        axDataSource.UnsetStrLit();
                    }
                }
                else
                {
                    if(axDataSource.IsSetStrLit())
                    {
                        strCache = axDataSource.strLit;
                    }
                    else
                    {
                        strCache = axDataSource.AddNewStrLit();
                    }
                    if(axDataSource.IsSetStrRef())
                    {
                        axDataSource.UnsetStrRef();
                    }
                }
                return strCache;
            }

            private CT_NumData RetrieveNumCache(CT_NumDataSource numDataSource,
                IXDDFNumericalDataSource<V> data)
            {
                CT_NumData numCache;
                if(data.IsReference)
                {
                    CT_NumRef numRef;
                    if(numDataSource.IsSetNumRef())
                    {
                        numRef = numDataSource.numRef;
                    }
                    else
                    {
                        numRef = numDataSource.AddNewNumRef();
                    }
                    if(numRef.IsSetNumCache())
                    {
                        numCache = numRef.numCache;
                    }
                    else
                    {
                        numCache = numRef.AddNewNumCache();
                    }
                    numRef.f = data.DataRangeReference;
                    if(numDataSource.IsSetNumLit())
                    {
                        numDataSource.UnsetNumLit();
                    }
                }
                else
                {
                    if(numDataSource.IsSetNumLit())
                    {
                        numCache = numDataSource.numLit;
                    }
                    else
                    {
                        numCache = numDataSource.AddNewNumLit();
                    }
                    if(numDataSource.IsSetNumRef())
                    {
                        numDataSource.UnsetNumRef();
                    }
                }
                return numCache;
            }

            private void FillStringCache(CT_StrData cache, int numOfPoints, IXDDFDataSource<T> data)
            {
                cache.SetPtArray(null); // unset old values
                if(cache.IsSetPtCount())
                {
                    cache.ptCount.val = (uint)numOfPoints;
                }
                else
                {
                    cache.AddNewPtCount().val = (uint)numOfPoints;
                }
                for(int i = 0; i < numOfPoints; ++i)
                {
                    string value = data.GetPointAt(i).ToString();
                    if(value != null)
                    {
                        CT_StrVal ctStrVal = cache.AddNewPt();
                        ctStrVal.idx = (uint)i;
                        ctStrVal.v = value;
                    }
                }
            }

            private void FillNumCache(CT_NumData cache, int numOfPoints, IXDDFNumericalDataSource<V> data)
            {
                string formatCode = data.FormatCode;
                if(formatCode == null)
                {
                    if(cache.IsSetFormatCode())
                    {
                        cache.UnsetFormatCode();
                    }
                }
                else
                {
                    cache.formatCode = formatCode;
                }
                cache.SetPtArray(null); // unset old values
                if(cache.IsSetPtCount())
                {
                    cache.ptCount.val = (uint)numOfPoints;
                }
                else
                {
                    cache.AddNewPtCount().val = (uint)numOfPoints;
                }
                for(int i = 0; i < numOfPoints; ++i)
                {
                    object value = data.GetPointAt(i);
                    if(value != null)
                    {
                        CT_NumVal ctNumVal = cache.AddNewPt();
                        ctNumVal.idx = (uint)i;
                        ctNumVal.v = value.ToString();
                    }
                }
            }
        }
    }
}