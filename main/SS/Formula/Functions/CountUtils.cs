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

namespace NPOI.SS.Formula.Functions
{
    using System;
    using NPOI.SS.Formula.Eval;
    /**
 * Common interface for the matching criteria.
 */
    public interface IMatchPredicate
    {
        bool Matches(ValueEval x);
    }
    public interface I_MatchAreaPredicate : IMatchPredicate
    {
        bool Matches(TwoDEval x, int rowIndex, int columnIndex);
    }
    /**
     * Common logic for COUNT, COUNTA and COUNTIF
     *
     * @author Josh Micich 
     */
    class CountUtils
    {

        private CountUtils()
        {
            // no instances of this class
        }

        /**
         * @return the number of evaluated cells in the range that match the specified criteria
         */
        public static int CountMatchingCellsInRef(RefEval refEval, IMatchPredicate criteriaPredicate)
        {
            int result = 0;

            for (int sIx = refEval.FirstSheetIndex; sIx <= refEval.LastSheetIndex; sIx++)
            {
                ValueEval ve = refEval.GetInnerValueEval(sIx);
                if (criteriaPredicate.Matches(ve))
                {
                    result++;
                }
            }
            return result;
        }
        public static int CountArg(ValueEval eval, IMatchPredicate criteriaPredicate)
        {
            if (eval == null)
            {
                throw new ArgumentException("eval must not be null");
            }
            if (eval is ThreeDEval)
            {
                return CountUtils.CountMatchingCellsInArea((ThreeDEval)eval, criteriaPredicate);
            }
            if (eval is TwoDEval)
            {
                throw new ArgumentException("Count requires 3D Evals, 2D ones aren't supported");
            }
            if (eval is RefEval)
            {
                return CountUtils.CountMatchingCellsInRef((RefEval)eval, criteriaPredicate);
            }
            return criteriaPredicate.Matches(eval) ? 1 : 0;
        }
        /**
     * @return the number of evaluated cells in the range that match the specified criteria
     */
        public static int CountMatchingCellsInArea(ThreeDEval areaEval, IMatchPredicate criteriaPredicate)
        {
            int result = 0;
            for (int sIx = areaEval.FirstSheetIndex; sIx <= areaEval.LastSheetIndex; sIx++)
            {
                int height = areaEval.Height;
                int width = areaEval.Width;
                for (int rrIx = 0; rrIx < height; rrIx++)
                {
                    for (int rcIx = 0; rcIx < width; rcIx++)
                    {
                        ValueEval ve = areaEval.GetValue(sIx, rrIx, rcIx);

                        if (criteriaPredicate is I_MatchAreaPredicate)
                        {
                            I_MatchAreaPredicate areaPredicate = (I_MatchAreaPredicate)criteriaPredicate;
                            if (!areaPredicate.Matches(areaEval, rrIx, rcIx)) continue;
                        }

                        if (criteriaPredicate.Matches(ve))
                        {
                            result++;
                        }
                    }
                }
            }
            return result;
        }
    }
}