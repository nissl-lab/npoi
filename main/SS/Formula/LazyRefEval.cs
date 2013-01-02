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
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Util;
    using NPOI.SS.Formula.PTG;

    /**
    *
    * @author Josh Micich 
    */
    public class LazyRefEval : RefEvalBase
    {

        private SheetRefEvaluator _evaluator;
        public LazyRefEval(int rowIndex, int columnIndex, SheetRefEvaluator sre)
            :base(rowIndex, columnIndex)
        {
           
            if (sre == null)
            {
                throw new ArgumentException("sre must not be null");
            }
            _evaluator = sre;
        }
        public LazyRefEval(Ref3DPtg ptg, SheetRefEvaluator sre)
            : base(ptg.Row, ptg.Column)
        {

            _evaluator = sre;
        }

        public override ValueEval InnerValueEval
        {
            get
            {
                return _evaluator.GetEvalForCell(Row, Column);
            }
        }

        public override AreaEval Offset(int relFirstRowIx, int relLastRowIx, int relFirstColIx, int relLastColIx)
        {

            AreaI area = new OffsetArea(Row, Column,
                    relFirstRowIx, relLastRowIx, relFirstColIx, relLastColIx);

            return new LazyAreaEval(area, _evaluator);
        }

        public override String ToString()
        {
            CellReference cr = new CellReference(Row, Column);
            StringBuilder sb = new StringBuilder();
            sb.Append(GetType().Name).Append("[");
            sb.Append(_evaluator.SheetName);
            sb.Append('!');
            sb.Append(cr.FormatAsString());
            sb.Append("]");
            return sb.ToString();
        }
    }
}