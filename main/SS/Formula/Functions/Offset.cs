/*
* Licensed to the Apache Software Foundation (ASF) Under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You Under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed Under the License is distributed on an "AS Is" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations Under the License.
*/


namespace NPOI.SS.Formula.Functions
{
    using System;
    using NPOI.SS.Formula.Eval;
    using System.Text;

    /**
     * Implementation for Excel function OFFSet()<p/>
     * 
     * OFFSet returns an area reference that Is a specified number of rows and columns from a 
     * reference cell or area.<p/>
     * 
     * <b>Syntax</b>:<br/>
     * <b>OFFSet</b>(<b>reference</b>, <b>rows</b>, <b>cols</b>, height, width)<p/>
     * <b>reference</b> Is the base reference.<br/>
     * <b>rows</b> Is the number of rows up or down from the base reference.<br/>
     * <b>cols</b> Is the number of columns left or right from the base reference.<br/>
     * <b>height</b> (default same height as base reference) Is the row Count for the returned area reference.<br/>
     * <b>width</b> (default same width as base reference) Is the column Count for the returned area reference.<br/>
     * 
     * @author Josh Micich
     */
    public class Offset : Function
    {
        // These values are specific to BIFF8
        private const int LAST_VALID_ROW_INDEX = 0xFFFF;
        private const int LAST_VALID_COLUMN_INDEX = 0xFF;


        /**
         * Exceptions are used within this class to help simplify flow control when error conditions
         * are enCountered 
         */
        [Serializable]
        private class EvalEx : Exception
        {
            private ErrorEval _error;

            public EvalEx(ErrorEval error)
            {
                _error = error;
            }
            public ErrorEval GetError()
            {
                return _error;
            }
        }

        /** 
         * A one dimensional base + offset.  Represents either a row range or a column range.
         * Two instances of this class toGether specify an area range.
         */
        /* package */
        public class LinearOffsetRange
        {

            private int _offset;
            private int _Length;

            public LinearOffsetRange(int offset, int length)
            {
                if (length == 0)
                {
                    // handled that condition much earlier
                    throw new ArgumentException("Length may not be zero");
                }
                _offset = offset;
                _Length = length;
            }

            public short FirstIndex
            {
                get
                {
                    return (short)_offset;
                }
            }
            public short LastIndex
            {
                get
                {
                    return (short)(_offset + _Length - 1);
                }
            }
            /**
             * Moves the range by the specified translation amount.<p/>
             * 
             * This method also 'normalises' the range: Excel specifies that the width and height 
             * parameters (Length field here) cannot be negative.  However, OFFSet() does produce
             * sensible results in these cases.  That behavior Is replicated here. <p/>
             * 
             * @param translationAmount may be zero negative or positive
             * 
             * @return the equivalent <c>LinearOffsetRange</c> with a positive Length, moved by the
             * specified translationAmount.
             */
            public LinearOffsetRange NormaliseAndTranslate(int translationAmount)
            {
                if (_Length > 0)
                {
                    if (translationAmount == 0)
                    {
                        return this;
                    }
                    return new LinearOffsetRange(translationAmount + _offset, _Length);
                }
                return new LinearOffsetRange(translationAmount + _offset + _Length + 1, -_Length);
            }

            public bool IsOutOfBounds(int lowValidIx, int highValidIx)
            {
                if (_offset < lowValidIx)
                {
                    return true;
                }
                if (LastIndex > highValidIx)
                {
                    return true;
                }
                return false;
            }
            public override String ToString()
            {
                StringBuilder sb = new StringBuilder(64);
                sb.Append(GetType().Name).Append(" [");
                sb.Append(_offset).Append("...").Append(LastIndex);
                sb.Append("]");
                return sb.ToString();
            }
        }


        /**
         * Encapsulates either an area or cell reference which may be 2d or 3d.
         */
        private class BaseRef
        {
            private const int INVALID_SHEET_INDEX = -1;
            private int _firstRowIndex;
            private int _firstColumnIndex;
            private int _width;
            private int _height;
            private RefEval _refEval;
		    private AreaEval _areaEval;

            public BaseRef(RefEval re)
            {
                _refEval = re;
                _areaEval = null;
                _firstRowIndex = re.Row;
                _firstColumnIndex = re.Column;
                _height = 1;
                _width = 1;
            }

            public BaseRef(AreaEval ae)
            {
                _refEval = null;
                _areaEval = ae;
                _firstRowIndex = ae.FirstRow;
                _firstColumnIndex = ae.FirstColumn;
                _height = ae.LastRow - ae.FirstRow + 1;
                _width = ae.LastColumn - ae.FirstColumn + 1;
            }

            public int Width
            {
                get
                {
                    return _width;
                }
            }

            public int Height
            {
                get
                {
                    return _height;
                }
            }

            public int FirstRowIndex
            {
                get
                {
                    return _firstRowIndex;
                }
            }

            public int FirstColumnIndex
            {
                get
                {
                    return _firstColumnIndex;
                }
            }
            public AreaEval Offset(int relFirstRowIx, int relLastRowIx,int relFirstColIx, int relLastColIx)
            {
                if (_refEval == null)
                {
                    return _areaEval.Offset(relFirstRowIx, relLastRowIx, relFirstColIx, relLastColIx);
                }
                return _refEval.Offset(relFirstRowIx, relLastRowIx, relFirstColIx, relLastColIx);
            }
        }

        public ValueEval Evaluate(ValueEval[] args, int srcCellRow, int srcCellCol)
        {

            if (args.Length < 3 || args.Length > 5)
            {
                return ErrorEval.VALUE_INVALID;
            }

            try
            {
                BaseRef baseRef = EvaluateBaseRef(args[0]);
                int rowOffset = EvaluateIntArg(args[1], srcCellRow, srcCellCol);
                int columnOffset = EvaluateIntArg(args[2], srcCellRow, srcCellCol);
                int height = baseRef.Height;
                int width = baseRef.Width;
                switch (args.Length)
                {
                    case 5:
                        width = EvaluateIntArg(args[4], srcCellRow, srcCellCol);
                        break;
                    case 4:
                        height = EvaluateIntArg(args[3], srcCellRow, srcCellCol);
                        break;
                }
                // Zero height or width raises #REF! error
                if (height == 0 || width == 0)
                {
                    return ErrorEval.REF_INVALID;
                }
                LinearOffsetRange rowOffsetRange = new LinearOffsetRange(rowOffset, height);
                LinearOffsetRange colOffsetRange = new LinearOffsetRange(columnOffset, width);
                return CreateOffset(baseRef, rowOffsetRange, colOffsetRange);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }



        private static AreaEval CreateOffset(BaseRef baseRef,
            LinearOffsetRange orRow, LinearOffsetRange orCol)
        {

            LinearOffsetRange absRows = orRow.NormaliseAndTranslate(baseRef.FirstRowIndex);
            LinearOffsetRange absCols = orCol.NormaliseAndTranslate(baseRef.FirstColumnIndex);

            if (absRows.IsOutOfBounds(0, LAST_VALID_ROW_INDEX))
            {
                throw new EvaluationException(ErrorEval.REF_INVALID);
            }
            if (absCols.IsOutOfBounds(0, LAST_VALID_COLUMN_INDEX))
            {
                throw new EvaluationException(ErrorEval.REF_INVALID);
            }
            return baseRef.Offset(orRow.FirstIndex, orRow.LastIndex, orCol.FirstIndex, orCol.LastIndex);
        }


        private static BaseRef EvaluateBaseRef(ValueEval eval)
        {

            if (eval is RefEval)
            {
                return new BaseRef((RefEval)eval);
            }
            if (eval is AreaEval)
            {
                return new BaseRef((AreaEval)eval);
            }
            if (eval is ErrorEval)
            {
                throw new EvalEx((ErrorEval)eval);
            }
            throw new EvalEx(ErrorEval.VALUE_INVALID);
        }


        /**
         * OFFSet's numeric arguments (2..5) have similar Processing rules
         */
        public static int EvaluateIntArg(ValueEval eval, int srcCellRow, int srcCellCol)
        {

            double d = EvaluateDoubleArg(eval, srcCellRow, srcCellCol);
            return ConvertDoubleToInt(d);
        }

        /**
         * Fractional values are silently truncated by Excel.
         * Truncation Is toward negative infinity.
         */
        /* package */
        public static int ConvertDoubleToInt(double d)
        {
            // Note - the standard java type conversion from double to int truncates toward zero.
            // but Math.floor() truncates toward negative infinity
            return (int)Math.Floor(d);
        }


        private static double EvaluateDoubleArg(ValueEval eval, int srcCellRow, int srcCellCol)
        {
            ValueEval ve = OperandResolver.GetSingleValue(eval, srcCellRow, srcCellCol);

            if (ve is NumericValueEval)
            {
                return ((NumericValueEval)ve).NumberValue;
            }
            if (ve is StringEval)
            {
                StringEval se = (StringEval)ve;
                double d = OperandResolver.ParseDouble(se.StringValue);
                if (double.IsNaN(d))
                {
                    throw new EvalEx(ErrorEval.VALUE_INVALID);
                }
                return d;
            }
            if (ve is BoolEval)
            {
                // in the context of OFFSet, bools resolve to 0 and 1.
                if (((BoolEval)ve).BooleanValue)
                {
                    return 1;
                }
                return 0;
            }
            throw new Exception("Unexpected eval type (" + ve.GetType().Name + ")");
        }
    }
}