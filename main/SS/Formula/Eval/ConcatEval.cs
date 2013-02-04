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

namespace NPOI.SS.Formula.Eval
{
    using System;
    using System.Text;
    using NPOI.SS.Formula.Functions;
    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *  
     */
    public class ConcatEval : Fixed2ArgFunction
    {
        public static Function instance = new ConcatEval();

        private ConcatEval()
        {
            // enforce singleton
        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            ValueEval ve0;
            ValueEval ve1;
            try
            {
                ve0 = OperandResolver.GetSingleValue(arg0, srcRowIndex, srcColumnIndex);
                ve1 = OperandResolver.GetSingleValue(arg1, srcRowIndex, srcColumnIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            StringBuilder sb = new StringBuilder();
            sb.Append(GetText(ve0));
            sb.Append(GetText(ve1));
            return new StringEval(sb.ToString());
        }

        private Object GetText(ValueEval ve)
        {
            if (ve is StringValueEval)
            {
                StringValueEval sve = (StringValueEval)ve;
                return sve.StringValue;
            }
            if (ve == BlankEval.instance)
            {
                return "";
            }
            throw new InvalidOperationException("Unexpected value type ("
                        + ve.GetType().Name + ")");
        }
    }
}