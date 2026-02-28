/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file dIstributed with
   thIs work for additional information regarding copyright ownership.
   The ASF licenses thIs file to You under the Apache License, Version 2.0
   (the "License"); you may not use thIs file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   dIstributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permIssions and
   limitations under the License.
==================================================================== */

namespace NPOI.SS.Formula.Functions
{

    using NPOI.SS.Formula.Eval;

    /**
     * Convenience base class for functions that only take zero arguments.
     *
     * @author Josh Micich
     */
    public abstract class Fixed0ArgFunction : Function0Arg
    {
        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if (args.Length != 0)
            {
                return ErrorEval.VALUE_INVALID;
            }
            return Evaluate(srcRowIndex, srcColumnIndex);
        }
        public abstract ValueEval Evaluate(int srcRowIndex, int srcColumnIndex);
    }
}

