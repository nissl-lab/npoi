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

using System;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;
namespace NPOI.SS.Formula.Eval.Forked
{

    /**
     * Represents a cell being used for forked Evaluation that has had a value Set different from the
     * corresponding cell in the shared master workbook.
     *
     * @author Josh Micich
     */
    class ForkedEvaluationCell : IEvaluationCell
    {

        private IEvaluationSheet _sheet;
        /** corresponding cell from master workbook */
        private IEvaluationCell _masterCell;
        private bool _boolValue;
        private CellType _cellType;
        private int _errorValue;
        private double _numberValue;
        private String _stringValue;

        public ForkedEvaluationCell(ForkedEvaluationSheet sheet, IEvaluationCell masterCell)
        {
            _sheet = sheet;
            _masterCell = masterCell;
            // start with value blank, but expect construction to be immediately
            SetValue(BlankEval.instance); // followed by a proper call to SetValue()
        }

        public Object IdentityKey
        {
            get
            {
                return _masterCell.IdentityKey;
            }
        }

        public void SetValue(ValueEval value)
        {
            Type cls = value.GetType();

            if (cls == typeof(NumberEval))
            {
                _cellType = CellType.Numeric;
                _numberValue = ((NumberEval)value).NumberValue;
                return;
            }
            if (cls == typeof(StringEval))
            {
                _cellType = CellType.String;
                _stringValue = ((StringEval)value).StringValue;
                return;
            }
            if (cls == typeof(BoolEval))
            {
                _cellType = CellType.Boolean;
                _boolValue = ((BoolEval)value).BooleanValue;
                return;
            }
            if (cls == typeof(ErrorEval))
            {
                _cellType = CellType.Error;
                _errorValue = ((ErrorEval)value).ErrorCode;
                return;
            }
            if (cls == typeof(BlankEval))
            {
                _cellType = CellType.Blank;
                return;
            }
            throw new ArgumentException("Unexpected value class (" + cls.Name + ")");
        }
        public void CopyValue(ICell destCell)
        {
            switch (_cellType)
            {
                case CellType.Blank: destCell.SetCellType(CellType.Blank); return;
                case CellType.Numeric: destCell.SetCellValue(_numberValue); return;
                case CellType.Boolean: destCell.SetCellValue(_boolValue); return;
                case CellType.String: destCell.SetCellValue(_stringValue); return;
                case CellType.Error: destCell.SetCellErrorValue((byte)_errorValue); return;
            }
            throw new InvalidOperationException("Unexpected data type (" + _cellType + ")");
        }

        private void CheckCellType(CellType expectedCellType)
        {
            if (_cellType != expectedCellType)
            {
                throw new Exception("Wrong data type (" + _cellType + ")");
            }
        }
        public CellType CellType
        {
            get
            {
                return _cellType;
            }
        }
        public bool BooleanCellValue
        {
            get
            {
                CheckCellType(CellType.Boolean);
                return _boolValue;
            }
        }
        public int ErrorCellValue
        {
            get
            {
                CheckCellType(CellType.Error);
                return _errorValue;
            }
        }
        public double NumericCellValue
        {
            get
            {
                CheckCellType(CellType.Numeric);
                return _numberValue;
            }
        }
        public String StringCellValue
        {
            get
            {
                CheckCellType(CellType.String);
                return _stringValue;
            }
        }
        public IEvaluationSheet Sheet
        {
            get
            {
                return _sheet;
            }
        }
        public int RowIndex
        {
            get
            {
                return _masterCell.RowIndex;
            }
        }
        public int ColumnIndex
        {
            get
            {
                return _masterCell.ColumnIndex;
            }
        }
        public CellType CachedFormulaResultType
        {
            get { return _masterCell.CachedFormulaResultType; }
        }
    }
}