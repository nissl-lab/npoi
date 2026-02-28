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
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.Formula.Eval;
    using System;


    /**
     * Implementation of Excel 'Analysis ToolPak' function MROUND()<br/>
     *
     * Returns a number rounded to the desired multiple.<p/>
     *
     * <b>Syntax</b><br/>
     * <b>MROUND</b>(<b>number</b>, <b>multiple</b>)
     *
     * <p/>
     *
     * @author Yegor Kozlov
     */
    class MRound : FreeRefFunction
    {

        public static FreeRefFunction Instance = new MRound();

        private MRound()
        {
            // enforce singleton
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            double number, multiple, result;

            if (args.Length != 2)
            {
                return ErrorEval.VALUE_INVALID;
            }

            try
            {
                number = OperandResolver.CoerceValueToDouble(OperandResolver.GetSingleValue(args[0], ec.RowIndex, ec.ColumnIndex));
                multiple = OperandResolver.CoerceValueToDouble(OperandResolver.GetSingleValue(args[1], ec.RowIndex, ec.ColumnIndex));

                if (multiple == 0.0)
                {
                    result = 0.0;
                }
                else
                {
                    if (number * multiple < 0)
                    {
                        // Returns #NUM! because the number and the multiple have different signs
                        throw new EvaluationException(ErrorEval.NUM_ERROR);
                    }
                    result = multiple * Math.Round(number / multiple, MidpointRounding.AwayFromZero);
                    
                }
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }
    }
}