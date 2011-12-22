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

    /**
     * Implementation for the function COUNTIF<p/>
     * 
     * Syntax: COUNTIF ( range, criteria )
     *    <table border="0" cellpAdding="1" cellspacing="0" summary="Parameter descriptions">
     *      <tr><th>range&nbsp;&nbsp;&nbsp;</th><td>is the range of cells to be Counted based on the criteria</td></tr>
     *      <tr><th>criteria</th><td>is used to determine which cells to Count</td></tr>
     *    </table>
     * <p/>
     * 
     * @author Josh Micich
     */
    public class Countif : Function
    {
        private class CmpOp
        {
            public const int NONE = 0;
            public const int EQ = 1;
            public const int NE = 2;
            public const int LE = 3;
            public const int LT = 4;
            public const int GT = 5;
            public const int GE = 6;

            public static CmpOp OP_NONE = op("", NONE);
            public static CmpOp OP_EQ = op("=", EQ);
            public static CmpOp OP_NE = op("<>", NE);
            public static CmpOp OP_LE = op("<=", LE);
            public static CmpOp OP_LT = op("<", LT);
            public static CmpOp OP_GT = op(">", GT);
            public static CmpOp OP_GE = op(">=", GE);
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
                    case NE: return cmpResult == 0;
                    case LT: return cmpResult < 0;
                    case LE: return cmpResult <= 0;
                    case GT: return cmpResult > 0;
                    case GE: return cmpResult <= 0;
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
        }

        private class NumberMatcher : I_MatchPredicate
        {

            private double _value;
		    private CmpOp _operator;

		    public NumberMatcher(double value, CmpOp optr) {
			    _value = value;
			    _operator = optr;
		    }

            public bool Matches(ValueEval x)
            {
                double testValue;
                if (x is StringEval)
                {
                    // if the tarGet(x) Is a string, but Parses as a number
                    // it may still Count as a match
                    StringEval se = (StringEval)x;
                    double val = OperandResolver.ParseDouble(se.StringValue);
                    if (double.IsNaN(val))
                    {
                        // x Is text that Is not a number
                        return false;
                    }
                    testValue = val;
                }
                else if (x is NumberEval)
                {
                    NumberEval ne = (NumberEval)x;
                    testValue = ne.NumberValue;
                }
                else 
                {
                    return false;
                }
                return _operator.Evaluate(testValue.CompareTo(_value));
            }
        }
        private class BooleanMatcher : I_MatchPredicate
        {

            private int _value;
            private CmpOp _operator;

            public BooleanMatcher(bool value,CmpOp optr)
            {
                _value = BoolToInt(value);
                _operator = optr;
            }

            private static int BoolToInt(bool? value)
            {
                if (value == null)
                {
                    throw new ArgumentException("null Boolean cannot be converted to Integer");
                }
                return value==true ? 1 : 0;
            }

            public bool Matches(ValueEval x)
            {
                int testValue=0;
                if (x is StringEval)
                {
                    if (true)
                    { // Change to false to observe more intuitive behaviour
                        // Note - Unlike with numbers, it seems that COUNTA never Matches 
                        // bool values when the tarGet(x) Is a string
                        return false;
                    }
#if !HIDE_UNREACHABLE_CODE
                    StringEval se = (StringEval)x;
                    bool? val = ParseBoolean(se.StringValue);
                    
                    if (val==null)
                    {
                        // x is text that is not a boolean
                        return false;
                    }
                    testValue = BoolToInt(val);
#endif
                }
                else if (x is BoolEval)
                {
                    BoolEval be = (BoolEval)x;
                    testValue = BoolToInt(be.BooleanValue);
                }
                else
                {
                    return false;
                }
                return _operator.Evaluate(testValue - _value);
            }
        }
        private class StringMatcher : I_MatchPredicate
        {

            private String _value;
            private CmpOp _operator;
            private Regex _pattern;

            public StringMatcher(String value, CmpOp optr)
            {
			    _value = value;
			    _operator = optr;
			    switch(optr.Code) {
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

            public bool Matches(ValueEval x)
            {
			    if (x is BlankEval) {
				    switch(_operator.Code) {
					    case CmpOp.NONE:
					    case CmpOp.EQ:
						    return _value.Length == 0;
				    }
				    // no other criteria matches a blank cell
				    return false;
			    }
			    if(!(x is StringEval)) {
				    // must always be string
				    // even if match str is wild, but contains only digits
				    // e.g. '4*7', NumberEval(4567) does not match
				    return false;
			    }
			    String testedValue = ((StringEval) x).StringValue;
			    if ((testedValue.Length < 1 && _value.Length < 1)
                    ||_value=="<>") 
                {
				    // odd case: criteria '=' behaves differently to criteria ''

				    switch(_operator.Code) {
					    case CmpOp.NONE: return true;
					    case CmpOp.EQ:   return false;
					    case CmpOp.NE:   return true;
				    }
				    return false;
			    }
			    if (_pattern != null) {
				    return _operator.Evaluate(_pattern.IsMatch(testedValue));
			    }
                return _operator.Evaluate(testedValue.CompareTo(_value));
            }

            /// <summary>
            /// Translates Excel countif wildcard strings into .NET regex strings
            /// </summary>
            /// <param name="value">Excel wildcard expression</param>
            /// <returns>return null if the specified value contains no special wildcard characters.</returns>
            private static Regex GetWildCardPattern(String value)
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
                                        sb.Append("\\").Append(ch);
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
                    return new Regex(sb.ToString());
                }
                return null;
            }
        }

        public ValueEval Evaluate(ValueEval[] args, int srcCellRow, int srcCellCol)
        {
            switch (args.Length)
            {
                case 2:
                    // expected
                    break;
                default:
                    // TODO - it doesn't seem to be possible to enter COUNTIF() into Excel with the wrong arg Count
                    // perhaps this should be an exception
                    return ErrorEval.VALUE_INVALID;
            }

            ValueEval range = (ValueEval)args[0];
            ValueEval criteriaArg = args[1];
            if (criteriaArg is RefEval)
            {
                // criteria Is not a literal value, but a cell reference
                // for example COUNTIF(B2:D4, E1)
                RefEval re = (RefEval)criteriaArg;
                criteriaArg = re.InnerValueEval;
            }
            else
            {
                // other non literal tokens such as function calls, have been fully Evaluated
                // for example COUNTIF(B2:D4, COLUMN(E1))
            }
            if (criteriaArg is BlankEval)
            {
                return NumberEval.ZERO;
            }
            I_MatchPredicate mp = CreateCriteriaPredicate(criteriaArg);
            return CountMatchingCellsInArea(range, mp);
        }
        /**
         * @return the number of Evaluated cells in the range that match the specified criteria
         */
        private ValueEval CountMatchingCellsInArea(ValueEval rangeArg, I_MatchPredicate criteriaPredicate)
        {
            int result;
            if (rangeArg is RefEval)
            {
                result = CountUtils.CountMatchingCell((RefEval)rangeArg, criteriaPredicate);
            }
            else if (rangeArg is AreaEval)
            {
                result = CountUtils.CountMatchingCellsInArea((AreaEval)rangeArg, criteriaPredicate);
            }
            else
            {
                throw new ArgumentException("Bad range arg type (" + rangeArg.GetType().Name + ")");
            }
            return new NumberEval(result);
        }

        public static I_MatchPredicate CreateCriteriaPredicate(ValueEval evaluatedCriteriaArg)
        {
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
            if (evaluatedCriteriaArg == BlankEval.instance)
            {
                return null;
            }
            throw new Exception("Unexpected type for criteria ("
                    + evaluatedCriteriaArg.GetType().Name + ")");
        }

        /**
         * When the second argument Is a string, many things are possible
         */
        private static I_MatchPredicate CreateGeneralMatchPredicate(StringEval stringEval)
        {
            String value = stringEval.StringValue;
            CmpOp optr = CmpOp.GetOperator(value);
            value = value.Substring(optr.Length);

            bool? boolVal= ParseBoolean(value);
            if(boolVal!=null)
            {
			    return new BooleanMatcher(boolVal==true?true:false, optr);
		    }


		    Double doubleVal = OperandResolver.ParseDouble(value);
		    if(!double.IsNaN(doubleVal)) {
			    return new NumberMatcher(doubleVal, optr);
		    }

		    //else - just a plain string with no interpretation.
            return new StringMatcher(value, optr);
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
                    if ("TRUE".Equals(strRep,StringComparison.InvariantCultureIgnoreCase))
                    {
                        return true;
                    }
                    break;
                case 'f':
                case 'F':
                    if ("FALSE".Equals(strRep, StringComparison.InvariantCultureIgnoreCase))
                    {
                        return false;
                    }
                    break;
            }
            return null ;
        }
    }
}