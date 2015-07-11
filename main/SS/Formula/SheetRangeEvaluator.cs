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

namespace NPOI.SS.Formula
{
    using System;

    using NPOI.SS.Formula.Eval;
    using System.Text;

    /**
     * Evaluator for returning cells or sheets for a range of sheets
     */
    public class SheetRangeEvaluator : ISheetRange
    {
        private int _firstSheetIndex;
        private int _lastSheetIndex;
        private SheetRefEvaluator[] _sheetEvaluators;

        public SheetRangeEvaluator(int firstSheetIndex, int lastSheetIndex, SheetRefEvaluator[] sheetEvaluators)
        {
            if (firstSheetIndex < 0)
            {
                throw new ArgumentException("Invalid firstSheetIndex: " + firstSheetIndex + ".");
            }
            if (lastSheetIndex < firstSheetIndex)
            {
                throw new ArgumentException("Invalid lastSheetIndex: " + lastSheetIndex + " for firstSheetIndex: " + firstSheetIndex + ".");
            }
            _firstSheetIndex = firstSheetIndex;
            _lastSheetIndex = lastSheetIndex;
            _sheetEvaluators = sheetEvaluators;
        }
        public SheetRangeEvaluator(int onlySheetIndex, SheetRefEvaluator sheetEvaluator)
            : this(onlySheetIndex, onlySheetIndex, new SheetRefEvaluator[] { sheetEvaluator })
        {
            
        }

        public SheetRefEvaluator GetSheetEvaluator(int sheetIndex)
        {
            if (sheetIndex < _firstSheetIndex || sheetIndex > _lastSheetIndex)
            {
                throw new ArgumentException("Invalid SheetIndex: " + sheetIndex +
                        " - Outside range " + _firstSheetIndex + " : " + _lastSheetIndex);
            }
            return _sheetEvaluators[sheetIndex - _firstSheetIndex];
        }

        public int FirstSheetIndex
        {
            get
            {
                return _firstSheetIndex;
            }
        }
        public int LastSheetIndex
        {
            get
            {
                return _lastSheetIndex;
            }
        }

        public String GetSheetName(int sheetIndex)
        {
            return GetSheetEvaluator(sheetIndex).SheetName;
        }
        public String SheetNameRange
        {
            get
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(GetSheetName(_firstSheetIndex));
                if (_firstSheetIndex != _lastSheetIndex)
                {
                    sb.Append(':');
                    sb.Append(GetSheetName(_lastSheetIndex));
                }
                return sb.ToString();
            }
        }

        public ValueEval GetEvalForCell(int sheetIndex, int rowIndex, int columnIndex)
        {
            return GetSheetEvaluator(sheetIndex).GetEvalForCell(rowIndex, columnIndex);
        }
    }

}