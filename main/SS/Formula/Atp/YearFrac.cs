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
    
    using System.Globalization;
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
                DateTime date = ParseDate(strVal);
                return NPOI.SS.UserModel.DateUtil.GetExcelDate(date, false);
            }
            return OperandResolver.CoerceValueToDouble(ve);
        }

        private static DateTime ParseDate(String strVal)
        {
            String[] parts = strVal.Split(new char[] { '/' });
            if (parts.Length != 3)
            {
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }
            String part2 = parts[2];
            int spacePos = part2.IndexOf(' ');
            if (spacePos > 0)
            {
                // drop time portion if present
                part2 = part2.Substring(0, spacePos);
            }
            int f0;
            int f1;
            int f2;
            try
            {
                f0 = int.Parse(parts[0], CultureInfo.InvariantCulture);
                f1 = int.Parse(parts[1], CultureInfo.InvariantCulture);
                f2 = int.Parse(part2, CultureInfo.InvariantCulture);
            }
            catch (FormatException)
            {
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }
            if (f0 < 0 || f1 < 0 || f2 < 0 || (f0 > 12 && f1 > 12 && f2 > 12))
            {
                // easy to see this cannot be a valid date
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }

            if (f0 >= 1900 && f0 < 9999)
            {
                // when 4 digit value appears first, the format is YYYY/MM/DD, regardless of OS settings
                return MakeDate(f0, f1, f2);
            }
            //// otherwise the format seems to depend on OS settings (default date format)
            //if (false)
            //{
            //    // MM/DD/YYYY is probably a good guess, if the in the US
            //    return MakeDate(f2, f0, f1);
            //}
            // TODO - find a way to choose the correct date format
            throw new InvalidCastException("Unable to determine date format for text '" + strVal + "'");
        }

        /**
         * @param month 1-based
         */
        private static DateTime MakeDate(int year, int month, int day)
        {
            return new DateTime(year, month - 1, day, 0, 0, 0,0);
        }

        private static int EvaluateIntArg(ValueEval arg, int srcCellRow, int srcCellCol)
        {
            ValueEval ve = OperandResolver.GetSingleValue(arg, srcCellRow, srcCellCol);
            return OperandResolver.CoerceValueToInt(ve);
        }
    }
}