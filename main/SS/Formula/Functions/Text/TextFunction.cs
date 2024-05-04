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
 * Created on May 22, 2005
 *
 */
namespace NPOI.SS.Formula.Functions
{
    using System;
    using NPOI.SS.Formula.Eval;

    public abstract class SingleArgTextFunc : TextFunction
    {

        protected SingleArgTextFunc()
        {
            // no fields to initialise
        }
        public override ValueEval EvaluateFunc(ValueEval[] args, int srcCellRow, int srcCellCol)
        {
            if (args.Length != 1)
            {
                return ErrorEval.VALUE_INVALID;
            }
            String arg = EvaluateStringArg(args[0], srcCellRow, srcCellCol);
            return Evaluate(arg);
        }
        public abstract ValueEval Evaluate(String arg);
    }

    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     */
    public abstract class TextFunction : Function
    {

        protected static String EMPTY_STRING = "";

        public static String EvaluateStringArg(ValueEval eval, int srcRow, int srcCol)
        {
            ValueEval ve = OperandResolver.GetSingleValue(eval, srcRow, srcCol);
            return OperandResolver.CoerceValueToString(ve);
        }
        public static int EvaluateIntArg(ValueEval arg, int srcCellRow, int srcCellCol)
        {
            ValueEval ve = OperandResolver.GetSingleValue(arg, srcCellRow, srcCellCol);
            return OperandResolver.CoerceValueToInt(ve);
        }
        public static double EvaluateDoubleArg(ValueEval arg, int srcCellRow, int srcCellCol) {
            ValueEval ve = OperandResolver.GetSingleValue(arg, srcCellRow, srcCellCol);
            return OperandResolver.CoerceValueToDouble(ve);
        }

        public ValueEval Evaluate(ValueEval[] args, int srcCellRow, int srcCellCol)
        {
            try
            {
                return EvaluateFunc(args, srcCellRow, srcCellCol);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }
        internal static bool IsPrintable(char c)
        {
            int charCode = (int)c;
            return charCode >= 32;
        }
        public abstract ValueEval EvaluateFunc(ValueEval[] args, int srcCellRow, int srcCellCol);

        /* ---------------------------------------------------------------------- */



        public static readonly Function LEN = new Len();
        public static readonly Function LOWER = new Lower();
        public static readonly Function UPPER = new Upper();
        /**
         * @author Manda Wilson &lt; wilson at c bio dot msk cc dot org &gt;
         */
        ///<summary>
        ///An implementation of the TRIM function:
        ///<para>
        /// Removes leading and trailing spaces from value if evaluated operand value is string.
        ///</para>
        ///</summary>
        public static readonly Function TRIM = new Trim();

        /*
         * @author Manda Wilson &lt; wilson at c bio dot msk cc dot org &gt;
        */
        ///<summary>
        ///An implementation of the MID function
        ///
        ///MID returns a specific number of
        ///characters from a text string, starting at the specified position.
        ///
        /// Syntax: MID(text, start_num, num_chars)
        ///</summary>
        public static readonly Function MID = new Mid();



        public static readonly Function LEFT = new LeftRight(true);
        public static readonly Function RIGHT = new LeftRight(false);

        public static readonly Function CONCATENATE = new Concatenate();

        public static readonly Function EXACT = new Exact();

        public static readonly Function TEXT = new Text();
        /**
         * @author Torstein Tauno Svendsen (torstei@officenet.no)
         */
        ///<summary>
        ///Implementation of the FIND() function.
        ///<para>
        /// Syntax: FIND(Find_text, within_text, start_num)
        ///</para>
        ///<para> FIND returns the character position of the first (case sensitive) occurrence of
        /// Find_text inside within_text.  The third parameter,
        /// start_num, is optional (default=1) and specifies where to start searching
        /// from.  Character positions are 1-based.</para>
        ///</summary>
        public static readonly Function FIND = new SearchFind(true);
        ///<summary>
        ///Implementation of the FIND() function. SEARCH is a case-insensitive version of FIND()
        ///<para>
        /// Syntax: SEARCH(Find_text, within_text, start_num)
        ///</para>
        ///</summary>
        public static readonly Function SEARCH = new SearchFind(false);

        public static readonly Function CLEAN = new Clean();
        public static readonly Function CHAR = new CHAR();
        public static readonly Function PROPER = new Proper();

    }
}