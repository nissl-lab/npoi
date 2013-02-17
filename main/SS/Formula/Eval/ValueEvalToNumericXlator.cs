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
/*
 * Created on May 14, 2005
 *
 */
namespace NPOI.SS.Formula.Eval
{
    using System;
    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *
     */
    public class ValueEvalToNumericXlator
    {

        public const int STRING_IS_PARSED = 0x0001;
        public const int BOOL_IS_PARSED = 0x0002;
        public const int BLANK_IS_PARSED = 0x0004; // => blanks are not ignored, Converted to 0

        public const int REF_STRING_IS_PARSED = 0x0008;
        public const int REF_BOOL_IS_PARSED = 0x0010;
        public const int REF_BLANK_IS_PARSED = 0x0020;

        public const int STRING_IS_INVALID_VALUE = 0x0800;

        private int flags;


        public ValueEvalToNumericXlator(int flags)
        {
            this.flags = flags;
        }

        /**
         * returned value can be either A NumericValueEval, BlankEval or ErrorEval.
         * The params can be either NumberEval, BoolEval, StringEval, or
         * RefEval
         * @param eval
         */
        public ValueEval AttemptXlateToNumeric(ValueEval eval)
        {
            ValueEval retval = null;

            if (eval == null)
            {
                retval = BlankEval.instance;
            }

            // most common case - least worries :)
            else if (eval is NumberEval)
            {
                retval = eval;
            }

            // booleval
            else if (eval is BoolEval)
            {
                retval = ((flags & BOOL_IS_PARSED) > 0)
                    ? (NumericValueEval)eval
                    : XlateBlankEval(BLANK_IS_PARSED);
            }

            // stringeval 
            else if (eval is StringEval)
            {
                retval = XlateStringEval((StringEval)eval); // TODO: recursive call needed
            }

            // refeval
            else if (eval is RefEval)
            {
                retval = XlateRefEval((RefEval)eval);
            }

            // erroreval
            else if (eval is ErrorEval)
            {
                retval = eval;
            }

            else if (eval is BlankEval)
            {
                retval = XlateBlankEval(BLANK_IS_PARSED);
            }

            // probably AreaEval? then not acceptable.
            else
            {
                throw new Exception("Invalid ValueEval type passed for conversion: " + eval.GetType());
            }

            return retval;
        }

        /**
         * no args are required since BlankEval has only one 
         * instance. If flag is Set, a zero
         * valued numbereval is returned, else BlankEval.INSTANCE
         * is returned.
         */
        private ValueEval XlateBlankEval(int flag)
        {
            return ((flags & flag) > 0)
                    ? (ValueEval)NumberEval.ZERO
                    : BlankEval.instance;
        }

        /**
         * uses the relevant flags to decode the supplied RefVal
         * @param eval
         */
        private ValueEval XlateRefEval(RefEval reval)
        {
            ValueEval eval = reval.InnerValueEval;

            // most common case - least worries :)
            if (eval is NumberEval)
            {
                return eval;
            }

            if (eval is BoolEval)
            {
                return ((flags & REF_BOOL_IS_PARSED) > 0)
                        ? (ValueEval)eval
                        : BlankEval.instance;
            }

            if (eval is StringEval)
            {
                return XlateRefStringEval((StringEval)eval);
            }

            if (eval is ErrorEval)
            {
                return eval;
            }

            if (eval is BlankEval)
            {
                return XlateBlankEval(REF_BLANK_IS_PARSED);
            }

            throw new Exception("Invalid ValueEval type passed for conversion: ("
                    + eval.GetType().Name + ")");
        }

        /**
         * uses the relevant flags to decode the StringEval
         * @param eval
         */
        private ValueEval XlateStringEval(StringEval eval)
        {

            if ((flags & STRING_IS_PARSED) > 0)
            {
                String s = eval.StringValue;
                double d = OperandResolver.ParseDouble(s);
                if (double.IsNaN(d))
                {
                    return ErrorEval.VALUE_INVALID;
                }
                return new NumberEval(d);
            }
            // strings are errors?
            if ((flags & STRING_IS_INVALID_VALUE) > 0)
            {
                return ErrorEval.VALUE_INVALID;
            }

            // ignore strings
            return XlateBlankEval(BLANK_IS_PARSED);
        }

        /**
         * uses the relevant flags to decode the StringEval
         * @param eval
         */
        private ValueEval XlateRefStringEval(StringEval sve)
        {
            if ((flags & REF_STRING_IS_PARSED) > 0)
            {
                String s = sve.StringValue;
                double d = OperandResolver.ParseDouble(s);
                if (double.IsNaN(d))
                {
                    return ErrorEval.VALUE_INVALID;
                }
                return new NumberEval(d);
            }
            // strings are blanks
            return BlankEval.instance;
        }
    }
}