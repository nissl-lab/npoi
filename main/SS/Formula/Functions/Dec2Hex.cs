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

using NPOI.SS.Formula.Eval;
using System;
namespace NPOI.SS.Formula.Functions
{

    /**
     * Implementation for Excel DELTA() function.<p/>
     * <p/>
     * <b>Syntax</b>:<br/> <b>DEC2HEX  </b>(<b>number</b>,<b>places</b> )<br/>
     * <p/>
     * Converts a decimal number to hexadecimal.
     *
     * The decimal integer you want to Convert. If number is negative, places is ignored
     * and this function returns a 10-character (40-bit) hexadecimal number in which the
     * most significant bit is the sign bit. The remaining 39 bits are magnitude bits.
     * Negative numbers are represented using two's-complement notation.
     *
     * <ul>
     * <li>If number &lt; -549,755,813,888 or if number &gt; 549,755,813,887, this function returns the #NUM! error value.</li>
     * <li>If number is nonnumeric, this function returns the #VALUE! error value.</li>
     * </ul>
     *
     * <h2>places</h2>
     *
     * The number of characters to use. The places argument is useful for pAdding the
     * return value with leading 0s (zeros).
     *
     * <ul>
     * <li>If this argument is omitted, this function uses the minimum number of characters necessary.</li>
     * <li>If this function requires more than places characters, it returns the #NUM! error value.</li>
     * <li>If this argument is nonnumeric, this function returns the #VALUE! error value.</li>
     * <li>If this argument is negative, this function returns the #NUM! error value.</li>
     * <li>If this argument Contains a decimal value, this function ignores the numbers to the right side of the decimal point.</li>
     * </ul>
     *
     * @author cedric dot walter @ gmail dot com
     */
    public class Dec2Hex : Var1or2ArgFunction, FreeRefFunction
    {
        public static FreeRefFunction instance = new Dec2Hex();
        private static long MinValue = -549755813888;
        private static long MaxValue = 549755813887;
        private static int DEFAULT_PLACES_VALUE = 10;


        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval number, ValueEval places)
        {
            ValueEval veText1;
            try
            {
                veText1 = OperandResolver.GetSingleValue(number, srcRowIndex, srcColumnIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            String strText1 = OperandResolver.CoerceValueToString(veText1);
            Double number1 = OperandResolver.ParseDouble(strText1);

            //If this number argument is non numeric, this function returns the #VALUE! error value.
            if (double.IsNaN(number1))
            {
                return ErrorEval.VALUE_INVALID;
            }

            //If number < -549,755,813,888 or if number > 549,755,813,887, this function returns the #NUM! error value.
            if (number1 < MinValue || number1 > MaxValue)
            {
                return ErrorEval.NUM_ERROR;
            }

            int placesNumber = 0;
            if (number1 < 0)
            {
                placesNumber = DEFAULT_PLACES_VALUE;
            }
            else if (places != null)
            {
                ValueEval placesValueEval;
                try
                {
                    placesValueEval = OperandResolver.GetSingleValue(places, srcRowIndex, srcColumnIndex);
                }
                catch (EvaluationException e)
                {
                    return e.GetErrorEval();
                }
                String placesStr = OperandResolver.CoerceValueToString(placesValueEval);
                Double placesNumberDouble = OperandResolver.ParseDouble(placesStr);

                //non numeric value
                if (double.IsNaN(placesNumberDouble))
                {
                    return ErrorEval.VALUE_INVALID;
                }

                //If this argument Contains a decimal value, this function ignores the numbers to the right side of the decimal point.
                placesNumber = (int)placesNumberDouble;

                if (placesNumber < 0)
                {
                    return ErrorEval.NUM_ERROR;
                }
            }

            String hex = "";
            if (placesNumber != 0)
            {
                hex = String.Format("{0:X" + placesNumber + "}", (int)number1);
            }
            else
            {
                hex = String.Format("{0:X}", (int)number1);
            }
            if (number1 < 0)
            {
                hex = "FF" + hex.Substring(2);
            }
            return new StringEval(hex.ToUpper());
        }


        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0)
        {
            return this.Evaluate(srcRowIndex, srcColumnIndex, arg0, null);
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length == 1)
            {
                return Evaluate(ec.RowIndex, ec.ColumnIndex, args[0]);
            }
            if (args.Length == 2)
            {
                return Evaluate(ec.RowIndex, ec.ColumnIndex, args[0], args[1]);
            }
            return ErrorEval.VALUE_INVALID;
        }
    }
}