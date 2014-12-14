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
 * Created on May 15, 2005
 *
 */
namespace NPOI.SS.Formula.Functions
{
    using NPOI.SS.Formula.Eval;

    /**
     * Counts the number of cells that contain numeric data within
     *  the list of arguments. 
     *
     * Excel Syntax
     * COUNT(value1,value2,...)
     * Value1, value2, ...   are 1 to 30 arguments representing the values or ranges to be Counted.
     * 
     * TODO: Check this properly Matches excel on edge cases
     *  like formula cells, error cells etc
     */
    public class Count : Function
    {
        private IMatchPredicate _predicate;

        public Count()
        {
            _predicate = defaultPredicate;
        }

        private Count(IMatchPredicate criteriaPredicate)
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

            for (int i = 0; i < nArgs; i++)
            {
                temp += CountUtils.CountArg(args[i], _predicate);

            }
            return new NumberEval(temp);
        }

        private class DefaultPredicate : IMatchPredicate
        {
            public bool Matches(ValueEval valueEval)
            {

                if (valueEval is NumberEval)
                {
                    // only numbers are counted
                    return true;
                }
                if (valueEval == MissingArgEval.instance)
                {
                    // oh yeah, and missing arguments
                    return true;
                }

                // error values and string values not counted
                return false;
            }
        }
        private static IMatchPredicate defaultPredicate = new DefaultPredicate();
        private class SubtotalPredicate : I_MatchAreaPredicate
        {
            public bool Matches(ValueEval valueEval)
            {
                return defaultPredicate.Matches(valueEval);
            }


            public bool Matches(TwoDEval x, int rowIndex, int columnIndex)
            {
                return !x.IsSubTotal(rowIndex, columnIndex);
            }

        }

        private static IMatchPredicate subtotalPredicate = new SubtotalPredicate();
        /**
     *  Create an instance of Count to use in {@link Subtotal}
     * <p>
     *     If there are other subtotals within argument refs (or nested subtotals),
     *     these nested subtotals are ignored to avoid double counting.
     * </p>
     *
     *  @see Subtotal
     */
        public static Count SubtotalInstance()
        {
            return new Count(subtotalPredicate);
        }
    }
}