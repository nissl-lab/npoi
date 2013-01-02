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
     * Convenience base class for any function which must take three or four
     * arguments
     *
     * @author Josh Micich
     */
    public abstract class Var3or4ArgFunction : Function3Arg, Function4Arg
    {

        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            switch (args.Length)
            {
                case 3:
                    return Evaluate(srcRowIndex, srcColumnIndex, args[0], args[1], args[2]);
                case 4:
                    return Evaluate(srcRowIndex, srcColumnIndex, args[0], args[1], args[2], args[3]);
            }
            return ErrorEval.VALUE_INVALID;
        }

        public abstract ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1, ValueEval arg2);

        public abstract ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1, ValueEval arg2, ValueEval arg3);
    }

}