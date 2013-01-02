/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

using System;
using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.SS.UserModel.Charts;

namespace NPOI.XSSF.UserModel.Charts
{
    /**
     * Package private class with utility methods.
     *
     * @author Roman Kashitsyn
     */
    internal class XSSFChartUtil
    {

        private XSSFChartUtil() { }

        /**
         * Builds CTAxDataSource object content from POI ChartDataSource.
         * @param ctAxDataSource OOXML data source to build
         * @param dataSource POI data source to use
         */
        public static void BuildAxDataSource<T>(CT_AxDataSource ctAxDataSource, IChartDataSource<T> dataSource)
        {
            if (dataSource.IsNumeric)
            {
                if (dataSource.IsReference)
                {
                    BuildNumRef(ctAxDataSource.AddNewNumRef(), dataSource);
                }
                else
                {
                    BuildNumLit(ctAxDataSource.AddNewNumLit(), dataSource);
                }
            }
            else
            {
                if (dataSource.IsReference)
                {
                    BuildStrRef(ctAxDataSource.AddNewStrRef(), dataSource);
                }
                else
                {
                    BuildStrLit(ctAxDataSource.AddNewStrLit(), dataSource);
                }
            }
        }

        /**
         * Builds CTNumDataSource object content from POI ChartDataSource
         * @param ctNumDataSource OOXML data source to build
         * @param dataSource POI data source to use
         */
        public static void BuildNumDataSource<T>(CT_NumDataSource ctNumDataSource,
                                              IChartDataSource<T> dataSource)
        {
            if (dataSource.IsReference)
            {
                BuildNumRef(ctNumDataSource.AddNewNumRef(), dataSource);
            }
            else
            {
                BuildNumLit(ctNumDataSource.AddNewNumLit(), dataSource);
            }
        }

        private static void BuildNumRef<T>(CT_NumRef ctNumRef, IChartDataSource<T> dataSource)
        {
            ctNumRef.f = (dataSource.FormulaString);
            CT_NumData cache = ctNumRef.AddNewNumCache();
            FillNumCache(cache, dataSource);
        }

        private static void BuildNumLit<T>(CT_NumData ctNumData, IChartDataSource<T> dataSource)
        {
            FillNumCache(ctNumData, dataSource);
        }

        private static void BuildStrRef<T>(CT_StrRef ctStrRef, IChartDataSource<T> dataSource)
        {
            ctStrRef.f = (dataSource.FormulaString);
            CT_StrData cache = ctStrRef.AddNewStrCache();
            FillStringCache(cache, dataSource);
        }

        private static void BuildStrLit<T>(CT_StrData ctStrData, IChartDataSource<T> dataSource)
        {
            FillStringCache(ctStrData, dataSource);
        }

        private static void FillStringCache<T>(CT_StrData cache, IChartDataSource<T> dataSource)
        {
            int numOfPoints = dataSource.PointCount;
            cache.AddNewPtCount().val = (uint)(numOfPoints);
            for (int i = 0; i < numOfPoints; ++i)
            {
                object value = dataSource.GetPointAt(i);
                if (value != null)
                {
                    CT_StrVal ctStrVal = cache.AddNewPt();
                    ctStrVal.idx = (uint)(i);
                    ctStrVal.v = (value.ToString());
                }
            }

        }

        private static void FillNumCache<T>(CT_NumData cache, IChartDataSource<T> dataSource)
        {
            int numOfPoints = dataSource.PointCount;
            cache.AddNewPtCount().val = (uint)(numOfPoints);
            for (int i = 0; i < numOfPoints; ++i)
            {
                double value = Convert.ToDouble(dataSource.GetPointAt(i));
                if (!double.IsNaN(value))
                {
                    CT_NumVal ctNumVal = cache.AddNewPt();
                    ctNumVal.idx = (uint)(i);
                    ctNumVal.v = (value.ToString());
                }
            }
        }
    }
}