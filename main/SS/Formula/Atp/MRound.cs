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
    using NPOI.SS.Util;
    using System;
    using System.Globalization;


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
                    result = RoundToMultiple(number, multiple);
                }
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }

        private static double RoundToMultiple(double number, double multiple)
        {
            // like POI, round on Excel's 15-digit view of both arguments in decimal arithmetic:
            // MROUND(0.7+0.1, 1.6) is 1.6 and MROUND(0.45, 0.1*3) is 0.6
            double quotient = Math.Abs(number / multiple);
            if (quotient < 0.25)
            {
                return 0.0;
            }
            if (quotient >= 1E25)
            {
                // the multiple is far below the precision of a double here
                return ToDouble(NumberToTextConverter.ToText(number));
            }

            // both are scaled by the same power of ten to stay within the range of decimal
            int shift = -(int)Math.Floor(Math.Log10(Math.Abs(multiple)));
            decimal decMultiple = ToDecimal(multiple, shift);
            decimal rounded = decMultiple * Math.Round(ToDecimal(number, shift) / decMultiple, MidpointRounding.AwayFromZero);
            return ToDouble(rounded.ToString(CultureInfo.InvariantCulture) + "E" + (-shift).ToString(CultureInfo.InvariantCulture));
        }

        private static decimal ToDecimal(double d, int shift)
        {
            string text = NumberToTextConverter.ToText(d);
            int exponent = shift;
            int e = text.IndexOf('E');
            if (e >= 0)
            {
                exponent += int.Parse(text.Substring(e + 1), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
                text = text.Substring(0, e);
            }
            return decimal.Parse(text + "E" + exponent.ToString(CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture);
        }

        private static double ToDouble(string text)
        {
            return double.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture);
        }
    }
}