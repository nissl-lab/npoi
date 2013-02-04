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
     * @author Josh Micich
     */
    public class Choose : Function
    {

        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if (args.Length < 2)
            {
                return ErrorEval.VALUE_INVALID;
            }

            try
            {
                int ix = EvaluateFirstArg(args[0], srcRowIndex, srcColumnIndex);
                if (ix < 1 || ix >= args.Length)
                {
                    return ErrorEval.VALUE_INVALID;
                }
                ValueEval result = OperandResolver.GetSingleValue(args[ix], srcRowIndex, srcColumnIndex);
                if (result == MissingArgEval.instance)
                {
                    return BlankEval.instance;
                }
                return result;
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }

        public static int EvaluateFirstArg(ValueEval arg0, int srcRowIndex, int srcColumnIndex)
        {
            ValueEval ev = OperandResolver.GetSingleValue(arg0, srcRowIndex, srcColumnIndex);
            return OperandResolver.CoerceValueToInt(ev);
        }
    }
}

