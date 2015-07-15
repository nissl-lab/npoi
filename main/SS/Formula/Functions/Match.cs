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
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.Formula;

    public class SingleValueVector : ValueVector
    {

        private ValueEval _value;

        public SingleValueVector(ValueEval value)
        {
            _value = value;
        }

        public ValueEval GetItem(int index)
        {
            if (index != 0)
            {
                throw new ArgumentException("Invalid index ("
                        + index + ") only zero is allowed");
            }
            return _value;
        }

        public int Size
        {
            get
            {
                return 1;
            }
        }
    }
    /**
     * Implementation for the MATCH() Excel function.<p/>
     * 
     * <b>Syntax:</b><br/>
     * <b>MATCH</b>(<b>lookup_value</b>, <b>lookup_array</b>, match_type)<p/>
     * 
     * Returns a 1-based index specifying at what position in the <b>lookup_array</b> the specified 
     * <b>lookup_value</b> Is found.<p/>
     * 
     * Specific matching behaviour can be modified with the optional <b>match_type</b> parameter.
     * 
     *    <table border="0" cellpAdding="1" cellspacing="0" summary="match_type parameter description">
     *      <tr><th>Value</th><th>Matching Behaviour</th></tr>
     *      <tr><td>1</td><td>(default) Find the largest value that Is less than or equal to lookup_value.
     *        The lookup_array must be in ascending <i>order</i>*.</td></tr>
     *      <tr><td>0</td><td>Find the first value that Is exactly equal to lookup_value.
     *        The lookup_array can be in any order.</td></tr>
     *      <tr><td>-1</td><td>Find the smallest value that Is greater than or equal to lookup_value.
     *        The lookup_array must be in descending <i>order</i>*.</td></tr>
     *    </table>
     * 
     * * Note regarding <i>order</i> - For the <b>match_type</b> cases that require the lookup_array to
     *  be ordered, MATCH() can produce incorrect results if this requirement Is not met.  Observed
     *  behaviour in Excel Is to return the lowest index value for which every item after that index
     *  breaks the match rule.<br/>
     *  The (ascending) sort order expected by MATCH() Is:<br/>
     *  numbers (low to high), strings (A to Z), bool (FALSE to TRUE)<br/>
     *  MATCH() ignores all elements in the lookup_array with a different type to the lookup_value. 
     *  Type conversion of the lookup_array elements Is never performed.
     *  
     *  
     * @author Josh Micich
     */
    public class Match : Function
    {


        public ValueEval Evaluate(ValueEval[] args, int srcCellRow, int srcCellCol)
        {

            double match_type = 1; // default

            switch (args.Length)
            {
                case 3:
                    try
                    {
                        match_type = EvaluateMatchTypeArg(args[2], srcCellRow, srcCellCol);
                    }
                    catch (EvaluationException)
                    {
                        // Excel/MATCH() seems to have slightly abnormal handling of errors with
                        // the last parameter.  Errors do not propagate up.  Every error Gets
                        // translated into #REF!
                        return ErrorEval.REF_INVALID;
                    }
                    break;
                case 2:
                    break;
                default:
                    return ErrorEval.VALUE_INVALID;
            }

            bool matchExact = match_type == 0;
            // Note - Excel does not strictly require -1 and +1
            bool FindLargestLessThanOrEqual = match_type > 0;


            try
            {
                ValueEval lookupValue = OperandResolver.GetSingleValue(args[0], srcCellRow, srcCellCol);
                ValueVector lookupRange = EvaluateLookupRange(args[1]);
                int index = FindIndexOfValue(lookupValue, lookupRange, matchExact, FindLargestLessThanOrEqual);
                return new NumberEval(index + 1); // +1 to Convert to 1-based
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }

        private static ValueVector EvaluateLookupRange(ValueEval eval)
        {
            if (eval is RefEval)
            {
                RefEval re = (RefEval)eval;
                if (re.NumberOfSheets == 1)
                {
                    return new SingleValueVector(re.GetInnerValueEval(re.FirstSheetIndex));
                }
                else
                {
                    return LookupUtils.CreateVector(re);
                }
            }
            if (eval is TwoDEval)
            {
                ValueVector result = LookupUtils.CreateVector((TwoDEval)eval);
                if (result == null)
                {
                    throw new EvaluationException(ErrorEval.NA);
                }
                return result;
            }

            // Error handling for lookup_range arg Is also Unusual
            if (eval is NumericValueEval)
            {
                throw new EvaluationException(ErrorEval.NA);
            }
            if (eval is StringEval)
            {
                StringEval se = (StringEval)eval;
                double d = OperandResolver.ParseDouble(se.StringValue);
                if (double.IsNaN(d))
                {
                    // plain string
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
                }
                // else looks like a number
                throw new EvaluationException(ErrorEval.NA);
            }
            throw new Exception("Unexpected eval type (" + eval.GetType().Name + ")");
        }



        private static double EvaluateMatchTypeArg(ValueEval arg, int srcCellRow, int srcCellCol)
        {
            ValueEval match_type = OperandResolver.GetSingleValue(arg, srcCellRow, srcCellCol);

            if (match_type is ErrorEval)
            {
                throw new EvaluationException((ErrorEval)match_type);
            }
            if (match_type is NumericValueEval)
            {
                NumericValueEval ne = (NumericValueEval)match_type;
                return ne.NumberValue;
            }
            if (match_type is StringEval)
            {
                StringEval se = (StringEval)match_type;
                double d = OperandResolver.ParseDouble(se.StringValue);
                if (double.IsNaN(d))
                {
                    // plain string
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
                }
                // if the string Parses as a number, it Is OK
                return d;
            }
            throw new Exception("Unexpected match_type type (" + match_type.GetType().Name + ")");
        }

        /**
         * @return zero based index
         */
        private static int FindIndexOfValue(ValueEval lookupValue, ValueVector lookupRange,
                bool matchExact, bool FindLargestLessThanOrEqual)
        {

            LookupValueComparer lookupComparer = CreateLookupComparer(lookupValue, matchExact);

            if (matchExact)
            {
                for (int i = 0; i < lookupRange.Size; i++)
                {
                    if (lookupComparer.CompareTo(lookupRange.GetItem(i)).IsEqual)
                    {
                        return i;
                    }
                }
                throw new EvaluationException(ErrorEval.NA);
            }

            if (FindLargestLessThanOrEqual)
            {
                // Note - backward iteration
                for (int i = lookupRange.Size - 1; i >= 0; i--)
                {
                    CompareResult cmp = lookupComparer.CompareTo(lookupRange.GetItem(i));
                    if (cmp.IsTypeMismatch)
                    {
                        continue;
                    }
                    if (!cmp.IsLessThan)
                    {
                        return i;
                    }
                }
                throw new EvaluationException(ErrorEval.NA);
            }

            // else - Find smallest greater than or equal to
            // TODO - Is binary search used for (match_type==+1) ?
            for (int i = 0; i < lookupRange.Size; i++)
            {
                CompareResult cmp = lookupComparer.CompareTo(lookupRange.GetItem(i));
                if (cmp.IsEqual)
                {
                    return i;
                }
                if (cmp.IsGreaterThan)
                {
                    if (i < 1)
                    {
                        throw new EvaluationException(ErrorEval.NA);
                    }
                    return i - 1;
                }
            }

            throw new EvaluationException(ErrorEval.NA);
        }

        private static LookupValueComparer CreateLookupComparer(ValueEval lookupValue, bool matchExact)
        {
            return LookupUtils.CreateLookupComparer(lookupValue, matchExact, true);
        }

        private static bool IsLookupValueWild(String stringValue)
        {
            if (stringValue.IndexOf('?') >= 0 || stringValue.IndexOf('*') >= 0)
            {
                return true;
            }
            return false;
        }
    }
}