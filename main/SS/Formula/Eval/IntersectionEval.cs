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
using System;
using NPOI.SS.Formula.Functions;
using NPOI.SS.Formula.Eval;

namespace NPOI.SS.Formula.Eval
{

    /**
     * @author Josh Micich
     */
    public class IntersectionEval : Fixed2ArgFunction
    {

        public static NPOI.SS.Formula.Functions.Function instance = new IntersectionEval();

        private IntersectionEval()
        {
            // enforces Singleton
        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {

            try
            {
                AreaEval reA = EvaluateRef(arg0);
                AreaEval reB = EvaluateRef(arg1);
                AreaEval result = ResolveRange(reA, reB);
                if (result == null)
                {
                    return ErrorEval.NULL_INTERSECTION;
                }
                return result;
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }

        /**
         * @return simple rectangular {@link AreaEval} which represents the intersection of areas
         * <c>aeA</c> and <c>aeB</c>. If the two areas do not intersect, the result is <code>null</code>.
         */
        private static AreaEval ResolveRange(AreaEval aeA, AreaEval aeB)
        {

            int aeAfr = aeA.FirstRow;
            int aeAfc = aeA.FirstColumn;
            int aeBlc = aeB.LastColumn;
            if (aeAfc > aeBlc)
            {
                return null;
            }
            int aeBfc = aeB.FirstColumn;
            if (aeBfc > aeA.LastColumn)
            {
                return null;
            }
            int aeBlr = aeB.LastRow;
            if (aeAfr > aeBlr)
            {
                return null;
            }
            int aeBfr = aeB.FirstRow;
            int aeAlr = aeA.LastRow;
            if (aeBfr > aeAlr)
            {
                return null;
            }


            int top = Math.Max(aeAfr, aeBfr);
            int bottom = Math.Min(aeAlr, aeBlr);
            int left = Math.Max(aeAfc, aeBfc);
            int right = Math.Min(aeA.LastColumn, aeBlc);

            return aeA.Offset(top - aeAfr, bottom - aeAfr, left - aeAfc, right - aeAfc);
        }

        private AreaEval EvaluateRef(ValueEval arg)
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

