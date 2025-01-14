/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */
using System;
using NPOI.SS.Formula.Eval;

namespace NPOI.SS.Formula.Functions
{
    /**
     * Implementation for the Excel function SUMIFS<br/>
     * <p>
     * Syntax : <br/>
     *  SUMIFS ( <b>sum_range</b>, <b>criteria_range1</b>, <b>criteria1</b>,
     *  [<b>criteria_range2</b>,  <b>criteria2</b>], ...) <br/>
     *    <ul>
     *      <li><b>sum_range</b> Required. One or more cells to sum, including numbers or names, ranges,
     *      or cell references that contain numbers. Blank and text values are ignored.</li>
     *      <li><b>criteria1_range</b> Required. The first range in which
     *      to evaluate the associated criteria.</li>
     *      <li><b>criteria1</b> Required. The criteria in the form of a number, expression,
     *        cell reference, or text that define which cells in the criteria_range1
     *        argument will be added</li>
     *      <li><b> criteria_range2, criteria2, ...</b>    Optional. Additional ranges and their associated criteria.
     *      Up to 127 range/criteria pairs are allowed.</li>
     *    </ul>
     * </p>
     *
     * @author Yegor Kozlov
     */
    public class Sumifs : Baseifs
    {
        public static FreeRefFunction instance = new Sumifs();

        protected override bool HasInitialRange => true;

        protected override IAggregator CreateAggregator()
        {
            return new MyAggregator();
        }

        private class MyAggregator : IAggregator
        {
            double accumulator = 0.0;

            public void AddValue(ValueEval value)
            {
                accumulator += (value is NumberEval) ? ((NumberEval) value).NumberValue : 0.0;
            }

            public ValueEval GetResult()
            {
                return new NumberEval(accumulator);
            }
        }

    }

}
