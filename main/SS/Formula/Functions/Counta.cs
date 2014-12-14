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
    using NPOI.SS.Formula.Eval;

    /**
     * Counts the number of cells that contain data within the list of arguments. 
     *
     * Excel Syntax
     * COUNTA(value1,value2,...)
     * Value1, value2, ...   are 1 to 30 arguments representing the values or ranges to be Counted.
     * 
     * @author Josh Micich
     */
    public class Counta : Function
    {
        private IMatchPredicate _predicate;

        public Counta()
        {
            _predicate = defaultPredicate;
        }

        private Counta(IMatchPredicate criteriaPredicate)
        {
            _predicate = criteriaPredicate;
        }
        public ValueEval Evaluate(ValueEval[] args, int srcCellRow, int srcCellCol)
        {
            int nArgs = args.Length;
            if (nArgs < 1)
            {
                // too few arguments
                return ErrorEval.VALUE_INVALID;
            }

            if (nArgs > 30)
            {
                // too many arguments
                return ErrorEval.VALUE_INVALID;
            }

            int temp = 0;
            // Note - observed behavior of Excel:
            // Error values like #VALUE!, #REF!, #DIV/0!, #NAME? etc don't cause this COUNTA to return an error
            // in fact, they seem to Get Counted

            for (int i = 0; i < nArgs; i++)
            {
                temp += CountUtils.CountArg(args[i], _predicate);

            }
            return new NumberEval(temp);
        }

        private static IMatchPredicate defaultPredicate = new DefaultPredicate();
        public class DefaultPredicate : IMatchPredicate
        {
            public bool Matches(ValueEval valueEval)
            {
                // Note - observed behavior of Excel:
                // Error values like #VALUE!, #REF!, #DIV/0!, #NAME? etc don't cause this COUNTA to return an error
                // in fact, they seem to get counted

                if (valueEval == BlankEval.instance)
                {
                    return false;
                }
                // Note - everything but BlankEval counts
                return true;
            }
        }
        private static IMatchPredicate subtotalPredicate = new SubtotalPredicate();
        public class SubtotalPredicate : I_MatchAreaPredicate
        {
            public bool Matches(ValueEval valueEval)
            {
                return defaultPredicate.Matches(valueEval);
            }

            /**
             * don't count cells that are subtotals
             */
            public bool Matches(TwoDEval areEval, int rowIndex, int columnIndex)
            {
                return !areEval.IsSubTotal(rowIndex, columnIndex);
            }

        }
        public static Counta SubtotalInstance()
        {
            return new Counta(subtotalPredicate);
        }
    }
}