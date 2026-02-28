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

    /**
     * An implementation of the TRIM function:
     * Removes leading and trailing spaces from value if Evaluated operand
     *  value Is string.
     * @author Manda Wilson &lt; wilson at c bio dot msk cc dot org &gt;
     */
    public class Trim : SingleArgTextFunc
    {

        public override ValueEval Evaluate(String arg)
        {
            return new StringEval(arg.Trim());
        }
    }
}