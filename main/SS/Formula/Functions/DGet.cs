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
     * Implementation of the DGet function:
     * Finds the value of a column in an area with given conditions.
     * 
     * TODO:
     * - wildcards ? and * in string conditions
     * - functions as conditions
     */
    public class DGet : IDStarAlgorithm
    {
        private ValueEval result;

        public void Reset()
        {
            result = null;
        }

        public bool ProcessMatch(ValueEval eval)
        {
            if (result == null) // First match, just Set the value.
            {
                result = eval;
            }
            else // There was a previous match, since there is only exactly one allowed, bail out1.
            {
                result = ErrorEval.NUM_ERROR;
                return false;
            }

            return true;
        }

        public ValueEval Result
        {
            get
            {
                if (result == null)
                {
                    return ErrorEval.VALUE_INVALID;
                }
                else
                {
                    return result;
                }
            }
        }
    }
}

