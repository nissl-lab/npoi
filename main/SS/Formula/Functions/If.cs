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
 * Created on Nov 25, 2006
 *
 */
namespace NPOI.SS.Formula.Functions
{
    using NPOI.SS.Formula.Eval;

    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     * 
     */
    public class If : Var2or3ArgFunction
    {

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            bool b;
            try
            {
                b = EvaluateFirstArg(arg0, srcRowIndex, srcColumnIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            if (b)
            {
                if (arg1 == MissingArgEval.instance)
                {
                    return BlankEval.instance;
                }
                return arg1;
            }
            return BoolEval.FALSE;
        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1,
                ValueEval arg2)
        {
            bool b;
            try
            {
                b = EvaluateFirstArg(arg0, srcRowIndex, srcColumnIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            if (b)
            {
                if (arg1 == MissingArgEval.instance)
                {
                    return BlankEval.instance;
                }
                return arg1;
            }
            if (arg2 == MissingArgEval.instance)
            {
                return BlankEval.instance;
            }
            return arg2;
        }

        public static bool EvaluateFirstArg(ValueEval arg, int srcCellRow, int srcCellCol)
        {
            ValueEval ve = OperandResolver.GetSingleValue(arg, srcCellRow, srcCellCol);
            bool? b = OperandResolver.CoerceValueToBoolean(ve, false);
            if (b == null)
            {
                return false;
            }
            return (bool)b;
        }
    }
}