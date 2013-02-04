/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.SS.Formula.Eval
{
    using NPOI.SS.Formula.Functions;
    /**
     * Implementation of Excel formula token '%'. <p/>
     * @author Josh Micich
     */
    public class PercentEval : Fixed1ArgFunction
    {
        public static Function instance = new PercentEval();

        private PercentEval()
        {
        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0)
        {
            double d0;
            try
            {
                ValueEval ve = OperandResolver.GetSingleValue(arg0, srcRowIndex, srcColumnIndex);
                d0 = OperandResolver.CoerceValueToDouble(ve);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            if (d0 == 0.0)
            { // this '==' matches +0.0 and -0.0
                return NumberEval.ZERO;
            }
            return new NumberEval(d0 / 100);
        }
    }
}