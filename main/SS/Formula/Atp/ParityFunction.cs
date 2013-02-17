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

namespace NPOI.SS.Formula.Atp
{
    using System;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.Formula;
    /**
     * Implementation of Excel 'Analysis ToolPak' function ISEVEN() ISODD()<br/>
     * 
     * @author Josh Micich
     */
    public class ParityFunction : FreeRefFunction
    {

        public static readonly FreeRefFunction IS_EVEN = new ParityFunction(0);
        public static readonly FreeRefFunction IS_ODD = new ParityFunction(1);
        private int _desiredParity;

        private ParityFunction(int desiredParity)
        {
            _desiredParity = desiredParity;
        }
        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec) 
        {
            if (args.Length != 1)
            {
                return ErrorEval.VALUE_INVALID;
            }

            int val;
            try
            {
                val = EvaluateArgParity(args[0], ec.RowIndex, ec.ColumnIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }

            return BoolEval.ValueOf(val == _desiredParity);
        }

        private static int EvaluateArgParity(ValueEval arg, int srcCellRow, int srcCellCol)
        {
            ValueEval ve = OperandResolver.GetSingleValue(arg, srcCellRow, srcCellCol);

            double d = OperandResolver.CoerceValueToDouble(ve);
            if (d < 0)
            {
                d = -d;
            }
            long v = (long)Math.Floor(d);
            return (int)(v & 0x0001);
        }
    }
}