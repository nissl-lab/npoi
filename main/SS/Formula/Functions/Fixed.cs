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
    using NPOI.SS.Formula.Eval;
    using System;

    public class Fixed : Function1Arg, Function2Arg, Function3Arg
    {

        public ValueEval Evaluate(
                int srcRowIndex, int srcColumnIndex,
                ValueEval arg0, ValueEval arg1, ValueEval arg2)
        {
            return doFixed(arg0, arg1, arg2, srcRowIndex, srcColumnIndex);
        }


        public ValueEval Evaluate(
                int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            return doFixed(arg0, arg1, BoolEval.FALSE, srcRowIndex, srcColumnIndex);
        }


        public ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0)
        {
            return doFixed(arg0, new NumberEval(2), BoolEval.FALSE, srcRowIndex, srcColumnIndex);
        }


        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            switch (args.Length)
            {
                case 1:
                    return doFixed(args[0], new NumberEval(2), BoolEval.FALSE,
                            srcRowIndex, srcColumnIndex);
                case 2:
                    return doFixed(args[0], args[1], BoolEval.FALSE,
                            srcRowIndex, srcColumnIndex);
                case 3:
                    return doFixed(args[0], args[1], args[2], srcRowIndex, srcColumnIndex);
            }
            return ErrorEval.VALUE_INVALID;
        }

        private ValueEval doFixed(
                ValueEval numberParam, ValueEval placesParam,
                ValueEval skipThousandsSeparatorParam,
                int srcRowIndex, int srcColumnIndex)
        {
            try
            {
                ValueEval numberValueEval =
                        OperandResolver.GetSingleValue(
                        numberParam, srcRowIndex, srcColumnIndex);
                decimal number = (decimal)OperandResolver.CoerceValueToDouble(numberValueEval);
                ValueEval placesValueEval =
                        OperandResolver.GetSingleValue(
                        placesParam, srcRowIndex, srcColumnIndex);
                int places = OperandResolver.CoerceValueToInt(placesValueEval);
                ValueEval skipThousandsSeparatorValueEval =
                        OperandResolver.GetSingleValue(
                        skipThousandsSeparatorParam, srcRowIndex, srcColumnIndex);
                bool? skipThousandsSeparator =
                        OperandResolver.CoerceValueToBoolean(
                        skipThousandsSeparatorValueEval, false);

                // Round number to respective places.
                //number = number.SetScale(places, RoundingMode.HALF_UP);
                if (places < 0)
                {
                    number = number / (decimal)Math.Pow(10, -places);
                    number = Math.Round(number, 0);
                    number = number * (decimal)Math.Pow(10, -places);
                }
                else
                    number = Math.Round(number, places);

                // Format number conditionally using a thousands separator.
                /*NumberFormat nf = NumberFormat.GetNumberInstance(Locale.US);
                DecimalFormat formatter = (DecimalFormat)nf;
                formatter.setGroupingUsed(!skipThousandsSeparator);
                formatter.setMinimumFractionDigits(places >= 0 ? places : 0);
                formatter.setMaximumFractionDigits(places >= 0 ? places : 0);
                String numberString = formatter.Format(number);*/
                //System.Globalization.CultureInfo culture = System.Globalization.CultureInfo.CurrentCulture;
                string numberString = skipThousandsSeparator!=null && skipThousandsSeparator.Value ?
                    number.ToString(places > 0 ? "F" + places : "F0")
                    : number.ToString(places > 0 ? "N" + places : "N0", System.Globalization.CultureInfo.InvariantCulture);
                // Return the result as a StringEval.
                
                return new StringEval(numberString);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }
    }

}