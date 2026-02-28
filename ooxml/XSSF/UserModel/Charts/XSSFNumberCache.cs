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

using NPOI.SS.Util;
using NPOI.SS.Util.CellWalk;
using NPOI.SS.UserModel;
using System;
using NPOI.OpenXmlFormats.Dml.Chart;
namespace NPOI.XSSF.UserModel.Charts
{

    /**
     * Package private class to fill chart's number reference with cached
     * numeric values. If a formula-typed cell referenced by data marker,
     * cell's value will be calculated and placed to cache. Numeric cells
     * will be placed to cache as is. Non-numeric cells will be ignored.
     *
     * @author Roman Kashitsyn
     */
    internal sealed class XSSFNumberCache
    {

        private CT_NumData ctNumData;

        internal XSSFNumberCache(CT_NumData ctNumData)
        {
            this.ctNumData = ctNumData;
        }

        /**
         * Builds new numeric cache Container.
         * @param marker data marker to use for cache Evaluation
         * @param ctNumRef parent number reference
         * @return numeric cache instance
         */
        internal static XSSFNumberCache BuildCache(DataMarker marker, CT_NumRef ctNumRef)
        {
            CellRangeAddress range = marker.Range;
            int numOfPoints = range.NumberOfCells;

            if (numOfPoints == 0)
            {
                // Nothing to do.
                return null;
            }

            XSSFNumberCache cache = new XSSFNumberCache(ctNumRef.AddNewNumCache());
            cache.SetPointCount(numOfPoints);

            IWorkbook wb = marker.Sheet.Workbook;
            IFormulaEvaluator Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            CellWalk cellWalk = new CellWalk(marker);
            NumCacheCellHandler numCacheHandler = new NumCacheCellHandler(Evaluator, cache.ctNumData);
            cellWalk.Traverse(numCacheHandler);
            return cache;
        }

        /**
         * Returns total count of points in cache. Some (or even all) of
         * them might be empty.
         * @return total count of points in cache
         */
        public long PointCount
        {
            get
            {
                CT_UnsignedInt pointCount = ctNumData.ptCount;
                if (pointCount != null)
                {
                    return pointCount.val;
                }
                else
                {
                    return 0L;
                }
            }
        }

        /**
         * Returns cache value at specified index.
         * @param index index of the point in cache
         * @return point value
         */
        public double GetValueAt(int index)
        {
            /* TODO: consider more effective algorithm. Left as is since
             * this method should be invoked mostly in Tests. */
            foreach (CT_NumVal pt in ctNumData.pt)
            {
                if (pt.idx == index)
                {
                    return Convert.ToDouble(pt.v);
                }
            }
            return 0.0;
        }

        private void SetPointCount(int numOfPoints)
        {
            ctNumData.AddNewPtCount().val = (uint)numOfPoints;
        }

        private sealed class NumCacheCellHandler : ICellHandler
        {

            private IFormulaEvaluator Evaluator;
            private CT_NumData ctNumData;

            public NumCacheCellHandler(IFormulaEvaluator Evaluator, CT_NumData ctnumdata)
            {
                this.Evaluator = Evaluator;
                this.ctNumData = ctnumdata;
            }

            public void OnCell(ICell cell, ICellWalkContext ctx)
            {
                double pointValue = GetOrEvalCellValue(cell);
                /* Silently ignore non-numeric values.
                 * This is Office default behaviour. */
                if (Double.IsNaN(pointValue))
                {
                    return;
                }

                CT_NumVal point = this.ctNumData.AddNewPt();
                point.idx = (uint)ctx.OrdinalNumber;
                point.v = (NumberToTextConverter.ToText(pointValue));
            }

            private double GetOrEvalCellValue(ICell cell)
            {
                CellType cellType = cell.CellType;

                if (cellType == CellType.NUMERIC)
                {
                    return cell.NumericCellValue;
                }
                else if (cellType == CellType.FORMULA)
                {
                    CellValue value = Evaluator.Evaluate(cell);
                    if (value.CellType == CellType.NUMERIC)
                    {
                        return value.NumberValue;
                    }
                }
                return Double.NaN;
            }

        }
    }
}

