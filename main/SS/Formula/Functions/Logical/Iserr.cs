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
using NPOI.SS.Formula.Eval;

namespace NPOI.SS.Formula.Functions
{
    /**
     * contribute by Pavel Egorov 
     * https://github.com/xoposhiy/npoi/commit/27b34a2389030c7115a666ace65daafda40d61af
     */
    /**
     * Implementation of Excel <tt>ISERR()</tt> function.<p/>
     *
     * <b>Syntax</b>:<br/>
     * <b>ISERR</b>(<b>value</b>)<p/>
     *
     * <b>value</b>  The value to be tested<p/>
     *
     * Returns the logical value <tt>TRUE</tt> if value refers to any error value except
     * <tt>'#N/A'</tt>; otherwise, it returns <tt>FALSE</tt>.
     */
    public class Iserr : LogicalFunction
    {
        protected override bool Evaluate(ValueEval arg)
        {
            return arg is ErrorEval && arg != ErrorEval.NA;

        }
    }
}
