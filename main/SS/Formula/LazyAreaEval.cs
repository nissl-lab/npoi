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
    using NPOI.HSSF.Record.Formula;
    using NPOI.HSSF.Record.Formula.Eval;
    using NPOI.SS.Util;

    /**
     *
     * @author Josh Micich 
     */
    public class LazyAreaEval : AreaEvalBase
    {

        private SheetRefEvaluator _evaluator;

        public LazyAreaEval(AreaI ptg, SheetRefEvaluator evaluator):base(ptg)
        {
            
            _evaluator = evaluator;
        }

        public LazyAreaEval(int firstRowIndex, int firstColumnIndex, int lastRowIndex,
                int lastColumnIndex, SheetRefEvaluator evaluator):
            base(firstRowIndex, firstColumnIndex, lastRowIndex, lastColumnIndex)
        {
            
            _evaluator = evaluator;
        }

        public override ValueEval GetRelativeValue(int relativeRowIndex, int relativeColumnIndex)
        {

            int rowIx = (relativeRowIndex + FirstRow) & 0xFFFF;
            int colIx = (relativeColumnIndex + FirstColumn) & 0x00FF;

            return _evaluator.GetEvalForCell(rowIx, colIx);
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
            sb.Append(_evaluator.SheetName);
            sb.Append('!');
            sb.Append(crA.FormatAsString());
            sb.Append(':');
            sb.Append(crB.FormatAsString());
            sb.Append("]");
            return sb.ToString();
        }
    }
}