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
    using System.Text;
    using System.Text.RegularExpressions;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;
    using System.Globalization;

    /**
     * Implementation for the function COUNTIF<p/>
     *
     * Syntax: COUNTIF ( range, criteria )
     *    <table border="0" cellpAdding="1" cellspacing="0" summary="Parameter descriptions">
     *      <tr><th>range </th><td>is the range of cells to be Counted based on the criteria</td></tr>
     *      <tr><th>criteria</th><td>is used to determine which cells to Count</td></tr>
     *    </table>
     * <p/>
     *
     * @author Josh Micich
     */
    public class Countif : Fixed2ArgFunction
    {
        public class CmpOp
        {
            public const int NONE = 0;
            public const int EQ = 1;
            public const int NE = 2;
            public const int LE = 3;
            public const int LT = 4;
            public const int GT = 5;
            public const int GE = 6;

            public static readonly CmpOp OP_NONE = op("", NONE);
            public static readonly CmpOp OP_EQ = op("=", EQ);
            public static readonly CmpOp OP_NE = op("<>", NE);
            public static readonly CmpOp OP_LE = op("<=", LE);
            public static readonly CmpOp OP_LT = op("<", LT);
            public static readonly CmpOp OP_GT = op(">", GT);
            public static readonly CmpOp OP_GE = op(">=", GE);
            private String _representation;
            private int _code;

            private static CmpOp op(String rep, int code)
            {
                return new CmpOp(rep, code);
            }
            private CmpOp(String representation, int code)
            {
                _representation = representation;
                _code = code;
            }
            /**
             * @return number of characters used to represent this operator
             */
            public int Length
            {
                get
                {
                    return _representation.Length;
                }
            }
            public int Code
            {
                get
                {
                    return _code;
                }
            }
            public static CmpOp GetOperator(String value)
            {
                int len = value.Length;
                if (len < 1)
                {
                    return OP_NONE;
                }

                char firstChar = value[0];

                switch (firstChar)
                {
                    case '=':
                        return OP_EQ;
                    case '>':
                        if (len > 1)
                        {
                            switch (value[1])
                            {
                                case '=':
                                    return OP_GE;
                            }
                        }
                        return OP_GT;
                    case '<':
                        if (len > 1)
                        {
                            switch (value[1])
                            {
                                case '=':
                                    return OP_LE;
                                case '>':
                                    return OP_NE;
                            }
                        }
                        return OP_LT;
                }
                return OP_NONE;
            }
            public bool Evaluate(bool cmpResult)
            {
                switch (_code)
                {
                    case NONE:
                    case EQ:
                        return cmpResult;
                    case NE:
                        return !cmpResult;
                }
                throw new Exception("Cannot call bool Evaluate on non-equality operator '"
                        + _representation + "'");
            }
            public bool Evaluate(int cmpResult)
            {
                switch (_code)
                {
                    case NONE:
                    case EQ:
                        return cmpResult == 0;
                    case NE: return cmpResult != 0;
                    case LT: return cmpResult < 0;
                    case LE: return cmpResult <= 0;
                    case GT: return cmpResult > 0;
                    case GE: return cmpResult >= 0;
                }
                throw new Exception("Cannot call bool Evaluate on non-equality operator '"
                        + _representation + "'");
            }
            public override String ToString()
            {
                StringBuilder sb = new StringBuilder(64);
                sb.Append(this.GetType().Name);
                sb.Append(" [").Append(_representation).Append("]");
                return sb.ToString();
            }
            public String Representation
            {
                get
                {
                    return _representation;
                }
            }
        }
        public abstract class MatcherBase : IMatchPredicate
        {
            private CmpOp _operator;

            public MatcherBase(CmpOp operator1)
            {
                _operator = operator1;
            }
            protected int Code
            {
                get
                {
                    return _operator.Code;
                }
            }
            protected bool Evaluate(int cmpResult)
            {
                return _operator.Evaluate(cmpResult);
            }
            protected bool Evaluate(bool cmpResult)
            {
                return _operator.Evaluate(cmpResult);
            }
            public override String ToString()
            {
                StringBuilder sb = new StringBuilder(64);
                sb.Append(this.GetType().Name).Append(" [");
                sb.Append(_operator.Representation);
                sb.Append(ValueText);
                sb.Append("]");
                return sb.ToString();
            }
            protected abstract String ValueText { get; }

            public abstract bool Matches(ValueEval x);
        }

        public class ErrorMatcher : MatcherBase
        {

            private int _value;

            public ErrorMatcher(int errorCode, CmpOp operator1)
                : base(operator1)
            {
                ;
                _value = errorCode;
            }
            protected override String ValueText
            {
                get
                {
                    return FormulaError.ForInt(_value).String;
                }
            }

            public override bool Matches(ValueEval x)
            {
                if (x is ErrorEval)
                {
                    int testValue = ((ErrorEval)x).ErrorCode;
                    return Evaluate(testValue - _value);
                }
                return false;
            }
            public int Value
            {
                get
                {
                    return _value;
                }
            }
        }
        private class NumberMatcher : MatcherBase
        {

            private double _value;

            public NumberMatcher(double value, CmpOp optr)
                : base(optr)
            {
                _value = value;
            }

            public override bool Matches(ValueEval x)
            {
                double testValue;
                if (x is StringEval)
                {
                    // if the target(x) is a string, but parses as a number
                    // it may still count as a match, only for the equality operator
                    switch (Code)
                    {
                        case CmpOp.EQ:
                        case CmpOp.NONE:
                            break;
                        case CmpOp.NE:
                            // Always matches (inconsistent with above two cases).
                            // for example '<>123' matches '123', '4', 'abc', etc
                            return true;
                        default:
                            // never matches (also inconsistent with above three cases).
                            // for example '>5' does not match '6',
                            return false;
                    }
                    StringEval se = (StringEval)x;
                    Double val = OperandResolver.ParseDouble(se.StringValue);
                    if (double.IsNaN(val))
                    {
                        // x is text that is not a number
                        return false;
                    }
                    return _value == val;
                }
                else if ((x is NumberEval))
                {
                    NumberEval ne = (NumberEval)x;
                    testValue = ne.NumberValue;
                }
                else if ((x is BlankEval))
                {
                    switch (Code)
                    {
                        case CmpOp.NE:
                            // Excel counts blank values in range as not equal to any value. See Bugzilla 51498
                            return true;
                        default:
                            return false;
                    }
                }
                else
                {
                    return false;
                }
                return Evaluate(testValue.CompareTo(_value));
            }

            protected override string ValueText
            {
                get { return _value.ToString(CultureInfo.InvariantCulture); }
            }
        }
        private class BooleanMatcher : MatcherBase
        {

            private int _value;

            public BooleanMatcher(bool value, CmpOp optr)
                : base(optr)
            {
                _value = BoolToInt(value);
            }

            private static int BoolToInt(bool value)
            {
                return value == true ? 1 : 0;
            }

            public override bool Matches(ValueEval x)
            {
                int testValue;
                if (x is StringEval)
                {
#if !HIDE_UNREACHABLE_CODE
                    if (true)
                    { // change to false to observe more intuitive behaviour
                        // Note - Unlike with numbers, it seems that COUNTIF never matches
                        // boolean values when the target(x) is a string
                        return false;
                    }
                    StringEval se = (StringEval)x;
                    Boolean? val = ParseBoolean(se.StringValue);
                    if (val == null)
                    {
                        // x is text that is not a boolean
                        return false;
                    }
                    testValue = BoolToInt(val.Value);
#else
                    return false;
#endif
                }
                else if ((x is BoolEval))
                {
                    BoolEval be = (BoolEval)x;
                    testValue = BoolToInt(be.BooleanValue);
                }
                else if ((x is BlankEval))
                {
                    switch (Code)
                    {
                        case CmpOp.NE:
                            // Excel counts blank values in range as not equal to any value. See Bugzilla 51498
                            return true;
                        default:
                            return false;
                    }
                }
                else if ((x is NumberEval))
                {
                    switch (Code)
                {
                    case CmpOp.NE:
                        // not-equals comparison of a number to boolean always returnes false
                        return true;
                    default:
                        return false;
                }
            }
                else
                {
                    return false;
                }
                return Evaluate(testValue - _value);
            }

            protected override string ValueText
            {
                get { return _value == 1 ? "TRUE" : "FALSE"; }
            }
        }
        internal class StringMatcher : MatcherBase
        {

            private String _value;
            private CmpOp _operator;
            private Regex _pattern;

            public StringMatcher(String value, CmpOp optr):base(optr)
            {
                _value = value;
                _operator = optr;
                switch (optr.Code)
                {
                    case CmpOp.NONE:
                    case CmpOp.EQ:
                    case CmpOp.NE:
                        _pattern = GetWildCardPattern(value);
                        break;
                    default:
                        _pattern = null;
                        break;
                }
            }
            public override bool Matches(ValueEval x)
            {
                if (x is BlankEval)
                {
                    switch (_operator.Code)
                    {
                        case CmpOp.NONE:
                        case CmpOp.EQ:
                            return _value.Length == 0;
                        case CmpOp.NE:
                            // pred '<>' matches empty string but not blank cell
                            // pred '<>ABC'  matches blank and 'not ABC'
                            return _value.Length != 0;
                    }
                    // no other criteria matches a blank cell
                    return false;
                }
                if (!(x is StringEval))
                {
                    if (_operator.Code==CmpOp.NE) return true;
                    // must almost always be string
                    // even if match str is wild, but contains only digits
                    // e.g. '4*7', NumberEval(4567) does not match
                    return false;
                }
                String testedValue = ((StringEval)x).StringValue;
                if ((testedValue.Length < 1 && _value.Length < 1))
                {
                    // odd case: criteria '=' behaves differently to criteria ''

                    switch (_operator.Code)
                    {
                        case CmpOp.NONE: return true;
                        case CmpOp.EQ: return false;
                        case CmpOp.NE: return true;
                    }
                    return false;
                }
                if (_pattern != null)
                {
                    return Evaluate(_pattern.IsMatch(testedValue));
                }
                //return Evaluate(testedValue.CompareTo(_value));
                return Evaluate(string.Compare(testedValue, _value, StringComparison.CurrentCultureIgnoreCase));
            }

            /// <summary>
            /// Translates Excel countif wildcard strings into .NET regex strings
            /// </summary>
            /// <param name="value">Excel wildcard expression</param>
            /// <returns>return null if the specified value contains no special wildcard characters.</returns>
            internal static Regex GetWildCardPattern(String value)
            {
                int len = value.Length;
                StringBuilder sb = new StringBuilder(len);
                sb.Append("^");
                bool hasWildCard = false;
                for (int i = 0; i < len; i++)
                {
                    char ch = value[i];
                    switch (ch)
                    {
                        case '?':
                            hasWildCard = true;
                            // match exactly one character
                            sb.Append('.');
                            continue;
                        case '*':
                            hasWildCard = true;
                            // match one or more occurrences of any character
                            sb.Append(".*");
                            continue;
                        case '~':
                            if (i + 1 < len)
                            {
                                ch = value[i + 1];
                                switch (ch)
                                {
                                    case '?':
                                    case '*':
                                        hasWildCard = true;
                                        //sb.Append("\\").Append(ch);
                                        sb.Append('[').Append(ch).Append(']');
                                        i++; // Note - incrementing loop variable here
                                        continue;
                                }
                            }
                            // else not '~?' or '~*'
                            sb.Append('~'); // just plain '~'
                            continue;
                        case '.':
                        case '$':
                        case '^':
                        case '[':
                        case ']':
                        case '(':
                        case ')':
                            // escape literal characters that would have special meaning in regex
                            sb.Append("\\").Append(ch);
                            continue;
                    }
                    sb.Append(ch);
                }
                sb.Append("$");
                if (hasWildCard)
                {
                    return new Regex(sb.ToString(), RegexOptions.IgnoreCase);
                }
                return null;
            }

            protected override string ValueText
            {
                get
                {
                    if (_pattern == null)
                    {
                        return _value;
                    }
                    return _pattern.ToString();
                }
            }
        }


        /**
     * @return the number of evaluated cells in the range that match the specified criteria
     */
        private double CountMatchingCellsInArea(ValueEval rangeArg, IMatchPredicate criteriaPredicate)
        {
            if (rangeArg is RefEval)
            {
                return CountUtils.CountMatchingCellsInRef((RefEval)rangeArg, criteriaPredicate);
            }
            else if (rangeArg is ThreeDEval)
            {
                return CountUtils.CountMatchingCellsInArea((ThreeDEval)rangeArg, criteriaPredicate);
            }
            else
            {
                throw new ArgumentException("Bad range arg type (" + rangeArg.GetType().Name + ")");
            }
        }


        /**
     *
     * @return the de-referenced criteria arg (possibly {@link ErrorEval})
     */
        private static ValueEval EvaluateCriteriaArg(ValueEval arg, int srcRowIndex, int srcColumnIndex)
        {
            try
            {
                return OperandResolver.GetSingleValue(arg, srcRowIndex, (short)srcColumnIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }
        /**
     * When the second argument is a string, many things are possible
     */
        private static IMatchPredicate CreateGeneralMatchPredicate(StringEval stringEval)
        {
            String value = stringEval.StringValue;
            CmpOp operator1 = CmpOp.GetOperator(value);
            value = value.Substring(operator1.Length);

            bool? booleanVal = ParseBoolean(value);
            if (booleanVal != null)
            {
                return new BooleanMatcher(booleanVal.Value, operator1);
            }

            Double doubleVal = OperandResolver.ParseDouble(value);
            if (!double.IsNaN(doubleVal))
            {
                return new NumberMatcher(doubleVal, operator1);
            }
            ErrorEval ee = ParseError(value);
            if (ee != null)
            {
                return new ErrorMatcher(ee.ErrorCode, operator1);
            }

            //else - just a plain string with no interpretation.
            return new StringMatcher(value, operator1);
        }
        /**
     * Creates a criteria predicate object for the supplied criteria arg
     * @return <code>null</code> if the arg evaluates to blank.
     */
        public static IMatchPredicate CreateCriteriaPredicate(ValueEval arg, int srcRowIndex, int srcColumnIndex)
        {

            ValueEval evaluatedCriteriaArg = EvaluateCriteriaArg(arg, srcRowIndex, srcColumnIndex);

            if (evaluatedCriteriaArg is NumberEval)
            {
                return new NumberMatcher(((NumberEval)evaluatedCriteriaArg).NumberValue, CmpOp.OP_NONE);
            }
            if (evaluatedCriteriaArg is BoolEval)
            {
                return new BooleanMatcher(((BoolEval)evaluatedCriteriaArg).BooleanValue, CmpOp.OP_NONE);
            }

            if (evaluatedCriteriaArg is StringEval)
            {
                return CreateGeneralMatchPredicate((StringEval)evaluatedCriteriaArg);
            }
            if (evaluatedCriteriaArg is ErrorEval)
            {
                return new ErrorMatcher(((ErrorEval)evaluatedCriteriaArg).ErrorCode, CmpOp.OP_NONE);
            }
            if (evaluatedCriteriaArg == BlankEval.instance)
            {
                return null;
            }
            throw new Exception("Unexpected type for criteria ("
                    + evaluatedCriteriaArg.GetType().Name + ")");
        }
        private static ErrorEval ParseError(String value)
        {
            if (value.Length < 4 || value[0] != '#')
            {
                return null;
            }
            if (value.Equals("#NULL!")) return ErrorEval.NULL_INTERSECTION;
            if (value.Equals("#DIV/0!")) return ErrorEval.DIV_ZERO;
            if (value.Equals("#VALUE!")) return ErrorEval.VALUE_INVALID;
            if (value.Equals("#REF!")) return ErrorEval.REF_INVALID;
            if (value.Equals("#NAME?")) return ErrorEval.NAME_INVALID;
            if (value.Equals("#NUM!")) return ErrorEval.NUM_ERROR;
            if (value.Equals("#N/A")) return ErrorEval.NA;

            return null;
        }
        /**
         * bool literals ('TRUE', 'FALSE') treated similarly but NOT same as numbers.
         */
        /* package */
        public static bool? ParseBoolean(String strRep)
        {
            if (strRep.Length < 1)
            {
                return null;
            }
            switch (strRep[0])
            {
                case 't':
                case 'T':
                    if ("TRUE".Equals(strRep, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                    break;
                case 'f':
                case 'F':
                    if ("FALSE".Equals(strRep, StringComparison.OrdinalIgnoreCase))
                    {
                        return false;
                    }
                    break;
            }
            return null;
        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            IMatchPredicate mp = CreateCriteriaPredicate(arg1, srcRowIndex, srcColumnIndex);
            if (mp == null)
            {
                // If the criteria arg is a reference to a blank cell, countif always returns zero.
                return NumberEval.ZERO;
            }
            double result = CountMatchingCellsInArea(arg0, mp);
            return new NumberEval(result);
        }
    }
}