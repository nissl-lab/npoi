/* ====================================================================
   Licensed To the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file To You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed To in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.SS.Formula
{

    using System;
    using System.Text;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.Util;

    /**
     * Provides Lazy Evaluation to 3D Ranges
     */
    public class LazyAreaEval : AreaEvalBase
    {

        private SheetRangeEvaluator _evaluator;

        public LazyAreaEval(AreaI ptg, SheetRangeEvaluator evaluator)
            : base(ptg, evaluator)
        {
            
            _evaluator = evaluator;
        }

        public LazyAreaEval(int firstRowIndex, int firstColumnIndex, int lastRowIndex,
                int lastColumnIndex, SheetRangeEvaluator evaluator) :
            base(evaluator, firstRowIndex, firstColumnIndex, lastRowIndex, lastColumnIndex)
        {
            
            _evaluator = evaluator;
        }

        public override ValueEval GetRelativeValue(int relativeRowIndex, int relativeColumnIndex)
        {
            return GetRelativeValue(FirstSheetIndex, relativeRowIndex, relativeColumnIndex);
        }
        public override ValueEval GetRelativeValue(int sheetIndex, int relativeRowIndex, int relativeColumnIndex)
        {
            int rowIx = (relativeRowIndex + FirstRow);
            int colIx = (relativeColumnIndex + FirstColumn);

            return _evaluator.GetEvalForCell(sheetIndex, rowIx, colIx);
        }

        public override TwoDEval GetRow(int rowIndex)
        {
            if (rowIndex >= Height)
            {
                throw new ArgumentException("Invalid rowIndex " + rowIndex
                        + ".  Allowable range is (0.." + Height + ").");
            }
            int absRowIx = FirstRow + rowIndex;
            return new LazyAreaEval(absRowIx, FirstColumn, absRowIx, LastColumn, _evaluator);
        }
        public override TwoDEval GetColumn(int columnIndex)
        {
            if (columnIndex >= Width)
            {
                throw new ArgumentException("Invalid columnIndex " + columnIndex
                        + ".  Allowable range is (0.." + Width + ").");
            }
            int absColIx = FirstColumn + columnIndex;
            return new LazyAreaEval(FirstRow, absColIx, LastRow, absColIx, _evaluator);
        }

        public override AreaEval Offset(int relFirstRowIx, int relLastRowIx, int relFirstColIx, int relLastColIx)
        {
            AreaI area = new OffsetArea(FirstRow, FirstColumn,
                    relFirstRowIx, relLastRowIx, relFirstColIx, relLastColIx);

            return new LazyAreaEval(area, _evaluator);
        }
        public override String ToString()
        {
            CellReference crA = new CellReference(FirstRow, FirstColumn);
            CellReference crB = new CellReference(LastRow, LastColumn);
            StringBuilder sb = new StringBuilder();
            sb.Append(GetType().Name).Append("[");
            sb.Append(_evaluator.SheetNameRange);
            sb.Append('!');
            sb.Append(crA.FormatAsString());
            sb.Append(':');
            sb.Append(crB.FormatAsString());
            sb.Append("]");
            return sb.ToString();
        }
        /**
        * @return  whether cell at rowIndex and columnIndex is a subtotal
        */
        public override bool IsSubTotal(int rowIndex, int columnIndex)
        {
            // delegate the query to the sheet evaluator which has access to internal ptgs
            SheetRefEvaluator _sre = _evaluator.GetSheetEvaluator(_evaluator.FirstSheetIndex);
            return _sre.IsSubTotal(FirstRow + rowIndex, FirstColumn + columnIndex);
        }
    }
}