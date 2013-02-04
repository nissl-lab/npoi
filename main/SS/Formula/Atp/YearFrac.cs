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
     * Implementation of Excel 'Analysis ToolPak' function YEARFRAC()<br/>
     * 
     * Returns the fraction of the year spanned by two dates.<p/>
     * 
     * <b>Syntax</b><br/>
     * <b>YEARFRAC</b>(<b>startDate</b>, <b>endDate</b>, basis)<p/>
     * 
     * The <b>basis</b> optionally specifies the behaviour of YEARFRAC as follows:
     * 
     * <table border="0" cellpadding="1" cellspacing="0" summary="basis parameter description">
     *   <tr><th>Value</th><th>Days per Month</th><th>Days per Year</th></tr>
     *   <tr align='center'><td>0 (default)</td><td>30</td><td>360</td></tr>
     *   <tr align='center'><td>1</td><td>actual</td><td>actual</td></tr>
     *   <tr align='center'><td>2</td><td>actual</td><td>360</td></tr>
     *   <tr align='center'><td>3</td><td>actual</td><td>365</td></tr>
     *   <tr align='center'><td>4</td><td>30</td><td>360</td></tr>
     * </table>
     * 
     */
    class YearFrac : FreeRefFunction
    {

        public static FreeRefFunction instance = new YearFrac();

        private YearFrac()
        {
            // enforce singleton
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            int srcCellRow = ec.RowIndex;
            int srcCellCol = ec.ColumnIndex;
            double result;
            try
            {
                int basis = 0; // default
                switch (args.Length)
                {
                    case 3:
                        basis = EvaluateIntArg(args[2], srcCellRow, srcCellCol);
                        break;
                    case 2:
                        //basis = EvaluateIntArg(args[2], srcCellRow, srcCellCol);
                        break;
                    default:
                        return ErrorEval.VALUE_INVALID;
                }
                double startDateVal = EvaluateDateArg(args[0], srcCellRow, srcCellCol);
                double endDateVal = EvaluateDateArg(args[1], srcCellRow, srcCellCol);
                result = YearFracCalculator.Calculate(startDateVal, endDateVal, basis);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }

            return new NumberEval(result);
        }

        private static double EvaluateDateArg(ValueEval arg, int srcCellRow, int srcCellCol)
        {
            ValueEval ve = OperandResolver.GetSingleValue(arg, srcCellRow, (short)srcCellCol);

            if (ve is StringEval)
            {
                String strVal = ((StringEval)ve).StringValue;
                Double dVal = OperandResolver.ParseDouble(strVal);
                if (!double.IsNaN(dVal))
                {
                    return dVal;
                }
                DateTime date = DateParser.ParseDate(strVal);
                return NPOI.SS.UserModel.DateUtil.GetExcelDate(date, false);
            }
            return OperandResolver.CoerceValueToDouble(ve);
        }


        private static int EvaluateIntArg(ValueEval arg, int srcCellRow, int srcCellCol)
        {
            ValueEval ve = OperandResolver.GetSingleValue(arg, srcCellRow, srcCellCol);
            return OperandResolver.CoerceValueToInt(ve);
        }
    }
}