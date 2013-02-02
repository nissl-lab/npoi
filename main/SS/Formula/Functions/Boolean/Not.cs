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
 * Created on May 9, 2005
 *
 */
namespace NPOI.SS.Formula.Functions
{
    using NPOI.SS.Formula.Eval;


    /*
     * @author Amol S. Deshmukh &lt; amol at apache dot org &gt;
     * The NOT bool function. Returns negation of specified value
     * (treated as a bool). If the specified arg Is a number,
     * then it Is true <=> 'number Is non-zero'
     */
    public class Not : Fixed1ArgFunction
    {
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0)
        {
            bool boolArgVal;
            try
            {
                ValueEval ve = OperandResolver.GetSingleValue(arg0, srcRowIndex, srcColumnIndex);
                bool? b = OperandResolver.CoerceValueToBoolean(ve, false);
                boolArgVal = b == null ? false : (bool)b;
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }

            return BoolEval.ValueOf(!boolArgVal);
        }
    }
}