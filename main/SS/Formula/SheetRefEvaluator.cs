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
    using NPOI.HSSF.Record.Formula.Eval;
    /**
     * 
     * 
     * @author Josh Micich
     */
    public class SheetRefEvaluator
    {

        private WorkbookEvaluator _bookEvaluator;
        private EvaluationTracker _tracker;
        private EvaluationSheet _sheet;
        private int _sheetIndex;

        public SheetRefEvaluator(WorkbookEvaluator bookEvaluator, EvaluationTracker tracker, int sheetIndex)
        {
            if (sheetIndex < 0)
            {
                throw new ArgumentException("Invalid sheetIndex: " + sheetIndex + ".");
            }
            _bookEvaluator = bookEvaluator;
            _tracker = tracker;
            _sheetIndex = sheetIndex;
        }

        public String SheetName
        {
            get
            {
                return _bookEvaluator.GetSheetName(_sheetIndex);
            }
        }

        public ValueEval GetEvalForCell(int rowIndex, int columnIndex)
        {
            return _bookEvaluator.EvaluateReference(this.Sheet, _sheetIndex, rowIndex, columnIndex, _tracker);
        }

        private EvaluationSheet Sheet
        {
            get
            {
                if (_sheet == null)
                {
                    _sheet = _bookEvaluator.GetSheet(_sheetIndex);
                }
                return _sheet;
            }
        }
    }
}