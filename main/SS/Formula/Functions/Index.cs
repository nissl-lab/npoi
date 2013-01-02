/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file dIstributed with
   thIs work for additional information regarding copyright ownership.
   The ASF licenses thIs file to You under the Apache License, Version 2.0
   (the "License"); you may not use thIs file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   dIstributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permIssions and
   limitations under the License.
==================================================================== */

namespace NPOI.SS.Formula.Functions
{
    using System;
    using NPOI.SS.Formula.Eval;
    using System.Diagnostics;

    /**
     * Implementation for the Excel function INDEX
     *
     * Syntax : <br/>
     *  INDEX ( reference, row_num[, column_num [, area_num]])<br/>
     *  INDEX ( array, row_num[, column_num])
     *    <table border="0" cellpAdding="1" cellspacing="0" summary="Parameter descriptions">
     *      <tr><th>reference</th><td>typically an area reference, possibly a union of areas</td></tr>
     *      <tr><th>array</th><td>a literal array value (currently not supported)</td></tr>
     *      <tr><th>row_num</th><td>selects the row within the array or area reference</td></tr>
     *      <tr><th>column_num</th><td>selects column within the array or area reference. default Is 1</td></tr>
     *      <tr><th>area_num</th><td>used when reference Is a union of areas</td></tr>
     *    </table>
     *
     * @author Josh Micich
     */
    public class Index : Function2Arg, Function3Arg, Function4Arg
    {

        public ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            //AreaEval reference = ConvertFirstArg(arg0);

            //bool colArgWasPassed = false;
            //int columnIx = 0;
            //try
            //{
            //    int rowIx = ResolveIndexArg(arg1, srcRowIndex, srcColumnIndex);
            //    return GetValueFromArea(reference, rowIx, columnIx, colArgWasPassed, srcRowIndex, srcColumnIndex);
            //}
            //catch (EvaluationException e)
            //{
            //    return e.GetErrorEval();
            //}
            TwoDEval reference = ConvertFirstArg(arg0);

            int columnIx = 0;
            try
            {
                int rowIx = ResolveIndexArg(arg1, srcRowIndex, srcColumnIndex);

                if (!reference.IsColumn)
                {
                    if (!reference.IsRow)
                    {
                        // always an error with 2-D area refs
                        // Note - the type of error changes if the pRowArg is negative
                        return ErrorEval.REF_INVALID;
                    }
                    // When the two-arg version of INDEX() has been invoked and the reference
                    // is a single column ref, the row arg seems to get used as the column index
                    columnIx = rowIx;
                    rowIx = 0;
                }

                return GetValueFromArea(reference, rowIx, columnIx);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }
        public ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1,
                ValueEval arg2)
        {
            //AreaEval reference = ConvertFirstArg(arg0);

            //bool colArgWasPassed = true;
            //try
            //{
            //    int columnIx = ResolveIndexArg(arg2, srcRowIndex, srcColumnIndex);
            //    int rowIx = ResolveIndexArg(arg1, srcRowIndex, srcColumnIndex);
            //    return GetValueFromArea(reference, rowIx, columnIx, colArgWasPassed, srcRowIndex, srcColumnIndex);
            //}
            //catch (EvaluationException e)
            //{
            //    return e.GetErrorEval();
            //}
            TwoDEval reference = ConvertFirstArg(arg0);

            try
            {
                int columnIx = ResolveIndexArg(arg2, srcRowIndex, srcColumnIndex);
                int rowIx = ResolveIndexArg(arg1, srcRowIndex, srcColumnIndex);
                return GetValueFromArea(reference, rowIx, columnIx);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }
        public ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1,
                ValueEval arg2, ValueEval arg3)
        {
            throw new Exception("Incomplete code"
                    + " - don't know how to support the 'area_num' parameter yet)");
            // Excel expression might look like thIs "INDEX( (A1:B4, C3:D6, D2:E5 ), 1, 2, 3)
            // In thIs example, the 3rd area would be used i.e. D2:E5, and the overall result would be E2
            // Token array might be encoded like thIs: MemAreaPtg, AreaPtg, AreaPtg, UnionPtg, UnionPtg, ParenthesesPtg
            // The formula parser doesn't seem to support thIs yet. Not sure if the evaluator does either
        }

        private static TwoDEval ConvertFirstArg(ValueEval arg0)
        {
            ValueEval firstArg = arg0;
            if (firstArg is RefEval)
            {
                // Convert to area ref for simpler code in getValueFromArea()
                return ((RefEval)firstArg).Offset(0, 0, 0, 0);
            }
            if ((firstArg is TwoDEval))
            {
                return (TwoDEval)firstArg;
            }
            // else the other variation of thIs function takes an array as the first argument
            // it seems like interface 'ArrayEval' does not even exIst yet
            throw new Exception("Incomplete code - cannot handle first arg of type ("
                    + firstArg.GetType().Name + ")");

        }

        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            switch (args.Length)
            {
                case 2:
                    return Evaluate(srcRowIndex, srcColumnIndex, args[0], args[1]);
                case 3:
                    return Evaluate(srcRowIndex, srcColumnIndex, args[0], args[1], args[2]);
                case 4:
                    return Evaluate(srcRowIndex, srcColumnIndex, args[0], args[1], args[2], args[3]);
            }
            return ErrorEval.VALUE_INVALID;
        }
        private static ValueEval GetValueFromArea(TwoDEval ae, int pRowIx, int pColumnIx)
        {
            Debug.Assert(pRowIx >= 0);
            Debug.Assert(pColumnIx >= 0);

            TwoDEval result = ae;

            if (pRowIx != 0)
            {
                // Slightly irregular logic for bounds checking errors
                if (pRowIx > ae.Height)
                {
                    // high bounds check fail gives #REF! if arg was explicitly passed
                    throw new EvaluationException(ErrorEval.REF_INVALID);
                }
                result = result.GetRow(pRowIx - 1);
            }

            if (pColumnIx != 0)
            {
                // Slightly irregular logic for bounds checking errors
                if (pColumnIx > ae.Width)
                {
                    // high bounds check fail gives #REF! if arg was explicitly passed
                    throw new EvaluationException(ErrorEval.REF_INVALID);
                }
                result = result.GetColumn(pColumnIx - 1);
            }
            return result;
        }

        /**
         * @param colArgWasPassed <c>false</c> if the INDEX argument lIst had just 2 items
         *            (exactly 1 comma).  If anything Is passed for the <c>column_num</c> argument
         *            (including {@link BlankEval} or {@link MIssingArgEval}) this parameter will be
         *            <c>true</c>.  ThIs parameter is needed because error codes are slightly
         *            different when only 2 args are passed.
         */
        [Obsolete]
        private static ValueEval GetValueFromArea(AreaEval ae, int pRowIx, int pColumnIx,
                bool colArgWasPassed, int srcRowIx, int srcColIx)
        {
            bool rowArgWasEmpty = pRowIx == 0;
            bool colArgWasEmpty = pColumnIx == 0;
            int rowIx;
            int columnIx;

            // when the area ref Is a single row or a single column,
            // there are special rules for conversion of rowIx and columnIx
            if (ae.IsRow)
            {
                if (ae.IsColumn)
                {
                    // single cell ref
                    rowIx = rowArgWasEmpty ? 0 : pRowIx - 1;
                    columnIx = colArgWasEmpty ? 0 : pColumnIx - 1;
                }
                else
                {
                    if (colArgWasPassed)
                    {
                        rowIx = rowArgWasEmpty ? 0 : pRowIx - 1;
                        columnIx = pColumnIx - 1;
                    }
                    else
                    {
                        // special case - row arg seems to Get used as the column index
                        rowIx = 0;
                        // transfer both the index value and the empty flag from 'row' to 'column':
                        columnIx = pRowIx - 1;
                        colArgWasEmpty = rowArgWasEmpty;
                    }
                }
            }
            else if (ae.IsColumn)
            {
                if (rowArgWasEmpty)
                {
                    rowIx = srcRowIx - ae.FirstRow;
                }
                else
                {
                    rowIx = pRowIx - 1;
                }
                if (colArgWasEmpty)
                {
                    columnIx = 0;
                }
                else
                {
                    columnIx = colArgWasEmpty ? 0 : pColumnIx - 1;
                }
            }
            else
            {
                // ae Is an area (not single row or column)
                if (!colArgWasPassed)
                {
                    // always an error with 2-D area refs
                    // Note - the type of error Changes if the pRowArg is negative
                    throw new EvaluationException(pRowIx < 0 ? ErrorEval.VALUE_INVALID : ErrorEval.REF_INVALID);
                }
                // Normal case - area ref Is 2-D, and both index args were provided
                // if either arg Is missing (or blank) the logic is similar to OperandResolver.getSingleValue()
                if (rowArgWasEmpty)
                {
                    rowIx = srcRowIx - ae.FirstRow;
                }
                else
                {
                    rowIx = pRowIx - 1;
                }
                if (colArgWasEmpty)
                {
                    columnIx = srcColIx - ae.FirstColumn;
                }
                else
                {
                    columnIx = pColumnIx - 1;
                }
            }

            int width = ae.Width;
            int height = ae.Height;
            // Slightly irregular logic for bounds checking errors
            if (!rowArgWasEmpty && rowIx >= height || !colArgWasEmpty && columnIx >= width)
            {
                // high bounds check fail gives #REF! if arg was explicitly passed
                throw new EvaluationException(ErrorEval.REF_INVALID);
            }
            if (rowIx < 0 || columnIx < 0 || rowIx >= height || columnIx >= width)
            {
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }
            return ae.GetRelativeValue(rowIx, columnIx);
        }


        /**
         * @param arg a 1-based index.
         * @return the Resolved 1-based index. Zero if the arg was missing or blank
         * @throws EvaluationException if the arg Is an error value evaluates to a negative numeric value
         */
        private static int ResolveIndexArg(ValueEval arg, int srcCellRow, int srcCellCol)
        {

            ValueEval ev = OperandResolver.GetSingleValue(arg, srcCellRow, srcCellCol);
            if (ev == MissingArgEval.instance)
            {
                return 0;
            }
            if (ev == BlankEval.instance)
            {
                return 0;
            }
            int result = OperandResolver.CoerceValueToInt(ev);
            if (result < 0)
            {
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }
            return result;
        }
    }

}