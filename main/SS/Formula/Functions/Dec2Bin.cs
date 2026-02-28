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
namespace NPOI.SS.Formula.Functions
{
    using System;

    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;

    /**
     * Implementation for Excel Bin2Dec() function.<p/>
     * <p/>
     * <b>Syntax</b>:<br/> <b>Bin2Dec  </b>(<b>number</b>,<b>[places]</b> )<br/>
     * <p/>
     * Converts a decimal number to binary.
     * <p/>
     * The DEC2BIN function syntax has the following arguments:
     * <ul>
     * <li>Number    Required. The decimal integer you want to Convert. If number is negative, valid place values are ignored and DEC2BIN returns a 10-character (10-bit) binary number in which the most significant bit is the sign bit. The remaining 9 bits are magnitude bits. Negative numbers are represented using two's-complement notation.</li>
     * <li>Places    Optional. The number of characters to use. If places is omitted, DEC2BIN uses the minimum number of characters necessary. Places is useful for pAdding the return value with leading 0s (zeros).</li>
     * </ul>
     * <p/>
     * Remarks
     * <ul>
     * <li>If number &lt; -512 or if number &gt; 511, DEC2BIN returns the #NUM! error value.</li>
     * <li>If number is nonnumeric, DEC2BIN returns the #VALUE! error value.</li>
     * <li>If DEC2BIN requires more than places characters, it returns the #NUM! error value.</li>
     * <li>If places is not an integer, it is tRuncated.</li>
     * <li>If places is nonnumeric, DEC2BIN returns the #VALUE! error value.</li>
     * <li>If places is zero or negative, DEC2BIN returns the #NUM! error value.</li>
     * </ul>
     *
     * @author cedric dot walter @ gmail dot com
     */
    public class Dec2Bin : Var1or2ArgFunction, FreeRefFunction
    {

        public static FreeRefFunction instance = new Dec2Bin();

        private static long MinValue = -512;
        private static long MaxValue = 511;
        private static int DEFAULT_PLACES_VALUE = 10;

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval numberVE, ValueEval placesVE)
        {
            ValueEval veText1;
            try
            {
                veText1 = OperandResolver.GetSingleValue(numberVE, srcRowIndex, srcColumnIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            String strText1 = OperandResolver.CoerceValueToString(veText1);
            Double number = OperandResolver.ParseDouble(strText1);

            //If this number argument is non numeric, this function returns the #VALUE! error value.
            if (double.IsNaN(number))
            {
                return ErrorEval.VALUE_INVALID;
            }

            //If number < -512 or if number > 512, this function returns the #NUM! error value.
            if (number< MinValue || number > MaxValue)
            {
                return ErrorEval.NUM_ERROR;
            }

            int placesNumber;
            if (number < 0 || placesVE == null)
            {
                placesNumber = DEFAULT_PLACES_VALUE;
            }
            else
            {
                ValueEval placesValueEval;
                try
                {
                    placesValueEval = OperandResolver.GetSingleValue(placesVE, srcRowIndex, srcColumnIndex);
                }
                catch (EvaluationException e)
                {
                    return e.GetErrorEval();
                }
                String placesStr = OperandResolver.CoerceValueToString(placesValueEval);
                Double placesNumberDouble = OperandResolver.ParseDouble(placesStr);

                //non numeric value
                if (double.IsNaN( placesNumberDouble))
                {
                    return ErrorEval.VALUE_INVALID;
                }

                //If this argument Contains a decimal value, this function ignores the numbers to the right side of the decimal point.
                placesNumber = (int)Math.Floor(placesNumberDouble);

                if (placesNumber < 0 || placesNumber == 0)
                {
                    return ErrorEval.NUM_ERROR;
                }
            }
            String binary = Convert.ToString((int)Math.Floor(number), 2);

            if (binary.Length > DEFAULT_PLACES_VALUE)
            {
                binary = binary.Substring(binary.Length - DEFAULT_PLACES_VALUE);
            }
            //If DEC2BIN requires more than places characters, it returns the #NUM! error value.
            if (binary.Length > placesNumber)
            {
                return ErrorEval.NUM_ERROR;
            }

            return new StringEval(binary);
        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval numberVE)
        {
            return this.Evaluate(srcRowIndex, srcColumnIndex, numberVE, null);
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