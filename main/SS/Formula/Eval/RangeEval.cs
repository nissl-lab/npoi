/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace NPOI.SS.Formula.Eval
{
    using System;

    using NPOI.SS.Formula.Functions;
    /**
     * 
     * @author Josh Micich 
     */
    public class RangeEval : Fixed2ArgFunction
    {

        public static Function instance = new RangeEval();

        private RangeEval()
        {
        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            try
            {
                AreaEval reA = EvaluateRef(arg0);
                AreaEval reB = EvaluateRef(arg1);
                return ResolveRange(reA, reB);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }

        private static AreaEval ResolveRange(AreaEval aeA, AreaEval aeB)
        {
            int aeAfr = aeA.FirstRow;
            int aeAfc = aeA.FirstColumn;

            int top = Math.Min(aeAfr, aeB.FirstRow);
            int bottom = Math.Max(aeA.LastRow, aeB.LastRow);
            int left = Math.Min(aeAfc, aeB.FirstColumn);
            int right = Math.Max(aeA.LastColumn, aeB.LastColumn);

            return aeA.Offset(top - aeAfr, bottom - aeAfr, left - aeAfc, right - aeAfc);
        }

        private static AreaEval EvaluateRef(ValueEval arg)
        {
            if (arg is AreaEval)
            {
                return (AreaEval)arg;
            }
            if (arg is RefEval)
            {
                return ((RefEval)arg).Offset(0, 0, 0, 0);
            }
            if (arg is ErrorEval)
            {
                throw new EvaluationException((ErrorEval)arg);
            }
            throw new ArgumentException("Unexpected ref arg class (" + arg.GetType().Name + ")");
        }
    }

}