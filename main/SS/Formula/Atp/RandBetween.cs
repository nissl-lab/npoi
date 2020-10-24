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
     * Implementation of Excel 'Analysis ToolPak' function RANDBETWEEN()<br/>
     *
     * Returns a random integer number between the numbers you specify.<p/>
     *
     * <b>Syntax</b><br/>
     * <b>RANDBETWEEN</b>(<b>bottom</b>, <b>top</b>)<p/>
     *
     * <b>bottom</b> is the smallest integer RANDBETWEEN will return.<br/>
     * <b>top</b> is the largest integer RANDBETWEEN will return.<br/>

     * @author Brendan Nolan
     */
    class RandBetween : FreeRefFunction
    {
        private Random _rnd;
        
        public static FreeRefFunction Instance = new RandBetween();

        private RandBetween()
        {
            //enforces singleton
            _rnd = new Random();
        }

        /**
         * Evaluate for RANDBETWEEN(). Must be given two arguments. Bottom must be greater than top.
         * Bottom is rounded up and top value is rounded down. After rounding top has to be set greater
         * than top.
         * 
         * @see org.apache.poi.ss.formula.functions.FreeRefFunction#evaluate(org.apache.poi.ss.formula.eval.ValueEval[], org.apache.poi.ss.formula.OperationEvaluationContext)
         */
        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {

            double bottom, top;

            if (args.Length != 2)
            {
                return ErrorEval.VALUE_INVALID;
            }

            try
            {
                bottom = OperandResolver.CoerceValueToDouble(OperandResolver.GetSingleValue(args[0], ec.RowIndex, ec.ColumnIndex));
                top = OperandResolver.CoerceValueToDouble(OperandResolver.GetSingleValue(args[1], ec.RowIndex, ec.ColumnIndex));
                if (bottom > top)
                {
                    return ErrorEval.NUM_ERROR;
                }
            }
            catch (EvaluationException)
            {
                return ErrorEval.VALUE_INVALID;
            }

            bottom = Math.Ceiling(bottom);
            top = Math.Floor(top);

            if (bottom > top)
            {
                top = bottom;
            }

            return new NumberEval((bottom + (int)(_rnd.NextDouble() * ((top - bottom) + 1))));

        }
    }
}
