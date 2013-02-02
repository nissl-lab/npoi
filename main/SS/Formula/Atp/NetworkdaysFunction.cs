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

using NPOI.SS.Formula.Functions;
using NPOI.SS.Formula.Eval;

namespace NPOI.SS.Formula.Atp
{
    /**
     * Implementation of Excel 'Analysis ToolPak' function NETWORKDAYS()<br/>
     * Returns the number of workdays given a starting and an ending date, considering an interval of holidays. A workday is any non
     * saturday/sunday date.
     * <p/>
     * <b>Syntax</b><br/>
     * <b>NETWORKDAYS</b>(<b>startDate</b>, <b>endDate</b>, holidays)
     * <p/>
     * 
     * @author jfaenomoto@gmail.com
     */
    public class NetworkdaysFunction : FreeRefFunction
    {

        public static FreeRefFunction instance = new NetworkdaysFunction(ArgumentsEvaluator.instance);

        private ArgumentsEvaluator evaluator;

        /**
         * Constructor.
         * 
         * @param anEvaluator an injected {@link ArgumentsEvaluator}.
         */
        private NetworkdaysFunction(ArgumentsEvaluator anEvaluator)
        {
            // enforces singleton
            this.evaluator = anEvaluator;
        }

        /**
         * Evaluate for NETWORKDAYS. Given two dates and a optional date or interval of holidays, determines how many working days are there
         * between those dates.
         * 
         * @return {@link ValueEval} for the number of days between two dates.
         */
        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 2 || args.Length > 3)
            {
                return ErrorEval.VALUE_INVALID;
            }

            int srcCellRow = ec.RowIndex;
            int srcCellCol = ec.ColumnIndex;

            double start, end;
            double[] holidays;
            try
            {
                start = this.evaluator.EvaluateDateArg(args[0], srcCellRow, srcCellCol);
                end = this.evaluator.EvaluateDateArg(args[1], srcCellRow, srcCellCol);
                if (start > end)
                {
                    return ErrorEval.NAME_INVALID;
                }
                ValueEval holidaysCell = args.Length == 3 ? args[2] : null;
                holidays = this.evaluator.EvaluateDatesArg(holidaysCell, srcCellRow, srcCellCol);
                return new NumberEval(WorkdayCalculator.instance.CalculateWorkdays(start, end, holidays));
            }
            catch (EvaluationException)
            {
                return ErrorEval.VALUE_INVALID;
            }
        }
    }
}
