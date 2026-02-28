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


namespace NPOI.SS.Formula.Functions
{
    using NPOI.SS.Formula.Eval;

    /**
     * Implementation for the Excel function MINIFS<p>
     *
     * Syntax : <p>
     *  MINIFS ( <b>min_range</b>, <b>criteria_range1</b>, <b>criteria1</b>,
     *  [<b>criteria_range2</b>,  <b>criteria2</b>], ...)
     *    <ul>
     *      <li><b>min_range</b> Required. One or more cells to determine the minimum value of, including numbers or names, ranges,
     *      or cell references that contain numbers. Blank and text values are ignored.</li>
     *      <li><b>criteria1_range</b> Required. The first range in which
     *      to evaluate the associated criteria.</li>
     *      <li><b>criteria1</b> Required. The criteria in the form of a number, expression,
     *        cell reference, or text that define which cells in the criteria_range1
     *        argument will be added</li>
     *      <li><b> criteria_range2, criteria2, ...</b>    Optional. Additional ranges and their associated criteria.
     *      Up to 127 range/criteria pairs are allowed.
     *    </ul>
     */

    public class Minifs : Baseifs
    {
        public static FreeRefFunction instance = new Minifs();
        protected override IAggregator CreateAggregator()
        {
            return new MyAggregator();
        }

        /**
         * https://support.microsoft.com/en-us/office/maxifs-function-dfd611e6-da2c-488a-919b-9b6376b28883
         * MAXIFS(max_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
         * need at least 3 arguments and need to have an odd number of arguments (max-range plus x*(criteria_range, criteria))
         */
        protected override bool HasInitialRange => true;

        private sealed class MyAggregator : IAggregator
        {
            double? accumulator = null;

            public void AddValue(ValueEval value)
            {
                double d = (value is NumberEval eval) ? eval.NumberValue : 0.0;
                if(accumulator == null || accumulator > d)
                {
                    accumulator = d;
                }
            }

            public ValueEval GetResult()
            {
                return new NumberEval(accumulator == null ? 0.0 : accumulator.Value);
            }
        }
    }

}
