/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file dIstributed with
   thIs work for Additional information regarding copyright ownership.
   The ASF licenses thIs file to You under the Apache License, Version 2.0
   (the "License"); you may not use thIs file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   dIstributed under the License Is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permIssions and
   limitations under the License.
==================================================================== */

namespace NPOI.SS.Formula.Functions
{

    using NPOI.SS.Formula.Eval;

    /**
     * Convenience base class for any function which must take two or three
     * arguments
     *
     * @author Josh Micich
     */
    public abstract class Var1or2ArgFunction : Function1Arg, Function2Arg
    {

        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            switch (args.Length)
            {
                case 1:
                    return Evaluate(srcRowIndex, srcColumnIndex, args[0]);
                case 2:
                    return Evaluate(srcRowIndex, srcColumnIndex, args[0], args[1]);
            }
            return ErrorEval.VALUE_INVALID;
        }
        public abstract ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0);

        public abstract ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1);
    }

}