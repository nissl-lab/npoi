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
    public interface I_MatchPredicate
    {
        bool Matches(ValueEval x);
    }
    public interface I_MatchAreaPredicate : I_MatchPredicate
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
         * @return 1 if the evaluated cell matches the specified criteria
         */
        public static int CountMatchingCell(RefEval refEval, I_MatchPredicate criteriaPredicate)
        {
            if (criteriaPredicate.Matches(refEval.InnerValueEval))
            {
                return 1;
            }
            return 0;
        }
        public static int CountArg(ValueEval eval, I_MatchPredicate criteriaPredicate)
        {
            if (eval == null)
            {
                throw new ArgumentException("eval must not be null");
            }
            if (eval is TwoDEval)
            {
                return CountUtils.CountMatchingCellsInArea((TwoDEval)eval, criteriaPredicate);
            }
            if (eval is RefEval)
            {
                return CountUtils.CountMatchingCell((RefEval)eval, criteriaPredicate);
            }
            return criteriaPredicate.Matches(eval) ? 1 : 0;
        }
        /**
	 * @return the number of evaluated cells in the range that match the specified criteria
	 */
        public static int CountMatchingCellsInArea(TwoDEval areaEval, I_MatchPredicate criteriaPredicate)
        {
            int result = 0;

            int height = areaEval.Height;
            int width = areaEval.Width;
            for (int rrIx = 0; rrIx < height; rrIx++)
            {
                for (int rcIx = 0; rcIx < width; rcIx++)
                {
                    ValueEval ve = areaEval.GetValue(rrIx, rcIx);

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
            return result;
        }
    }
}