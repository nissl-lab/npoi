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
using System.Collections.Generic;
using System.Text;
using NPOI.SS.Formula.Functions;
using NPOI.SS.Formula.Eval;

namespace NPOI.SS.Formula.Atp
{
    class IfError : FreeRefFunction
    {

        public static FreeRefFunction Instance = new IfError();

        private IfError()
        {
            // enforce singleton
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length != 2)
            {
                return ErrorEval.VALUE_INVALID;
            }

            ValueEval val;
            try
            {
                val = EvaluateInternal(args[0], args[1], ec.RowIndex, ec.ColumnIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }

            return val;
        }

        private static ValueEval EvaluateInternal(ValueEval arg, ValueEval iferror, int srcCellRow, int srcCellCol)
        {
            arg = WorkbookEvaluator.DereferenceResult(arg, srcCellRow, srcCellCol);
            if (arg is ErrorEval)
            {
                return iferror;
            }
            else
            {
                return arg;
            }
        }
    }
}
