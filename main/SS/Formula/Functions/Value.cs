/*
* Licensed to the Apache Software Foundation (ASF) Under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You Under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed Under the License Is distributed on an "AS Is" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations Under the License.
*/
/*
 * Created on May 15, 2005
 *
 */
namespace NPOI.SS.Formula.Functions
{
    using System;
    using System.Text;
    using NPOI.SS.Formula.Eval;
    using System.Globalization;

    public class Value : Fixed1ArgFunction
    {
        /** "1,0000" is valid, "1,00" is not */
        private const int MIN_DISTANCE_BETWEEN_THOUSANDS_SEPARATOR = 4;
        private const double ZERO = 0.0;

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0)
        {
            ValueEval veText;
            try
            {
                veText = OperandResolver.GetSingleValue(arg0, srcRowIndex, srcColumnIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            String strText = OperandResolver.CoerceValueToString(veText);
            Double result = ConvertTextToNumber(strText);
            if (Double.IsNaN(result))
            {
                return ErrorEval.VALUE_INVALID;
            }
            return new NumberEval(result);
        }

        /**
         * TODO see if the same functionality is needed in {@link OperandResolver#parseDouble(String)}
         *
         * @return <code>null</code> if there is any problem converting the text
         */
        private static Double ConvertTextToNumber(String strText)
        {
            bool foundCurrency = false;
            bool foundUnaryPlus = false;
            bool foundUnaryMinus = false;
            bool foundPercentage = false;

            int len = strText.Length;
            int i;
            for (i = 0; i < len; i++)
            {
                char ch = strText[i];
                if (Char.IsDigit(ch) || ch == '.')
                {
                    break;
                }
                switch (ch)
                {
                    case ' ':
                        // intervening spaces between '$', '-', '+' are OK
                        continue;
                    case '$':
                        if (foundCurrency)
                        {
                            // only one currency symbols is allowed
                            return Double.NaN;
                        }
                        foundCurrency = true;
                        continue;
                    case '+':
                        if (foundUnaryMinus || foundUnaryPlus)
                        {
                            return Double.NaN;
                        }
                        foundUnaryPlus = true;
                        continue;
                    case '-':
                        if (foundUnaryMinus || foundUnaryPlus)
                        {
                            return Double.NaN;
                        }
                        foundUnaryMinus = true;
                        continue;
                    default:
                        // all other characters are illegal
                        return Double.NaN;
                }
            }
            if (i >= len)
            {
                // didn't find digits or '.'
                if (foundCurrency || foundUnaryMinus || foundUnaryPlus)
                {
                    return Double.NaN;
                }
                return ZERO;
            }

            // remove thousands separators

            bool foundDecimalPoint = false;
            int lastThousandsSeparatorIndex = short.MinValue;

            StringBuilder sb = new StringBuilder(len);
            for (; i < len; i++)
            {
                char ch = strText[i];
                if (Char.IsDigit(ch))
                {
                    sb.Append(ch);
                    continue;
                }
                switch (ch)
                {
                    case ' ':
                        String remainingTextTrimmed = strText.Substring(i).Trim();
                        // support for value[space]%
                        if (remainingTextTrimmed.Equals("%")) {
                            foundPercentage= true;
                            break;
                        }
                        if (remainingTextTrimmed.Length > 0)
                        {
                            // intervening spaces not allowed once the digits start
                            return Double.NaN;
                        }
                        break;
                    case '.':
                        if (foundDecimalPoint)
                        {
                            return Double.NaN;
                        }
                        if (i - lastThousandsSeparatorIndex < MIN_DISTANCE_BETWEEN_THOUSANDS_SEPARATOR)
                        {
                            return Double.NaN;
                        }
                        foundDecimalPoint = true;
                        sb.Append('.');
                        continue;
                    case ',':
                        if (foundDecimalPoint)
                        {
                            // thousands separators not allowed after '.' or 'E'
                            return Double.NaN;
                        }
                        int distanceBetweenThousandsSeparators = i - lastThousandsSeparatorIndex;
                        // as long as there are 3 or more digits between
                        if (distanceBetweenThousandsSeparators < MIN_DISTANCE_BETWEEN_THOUSANDS_SEPARATOR)
                        {
                            return Double.NaN;
                        }
                        lastThousandsSeparatorIndex = i;
                        // don't append ','
                        continue;

                    case 'E':
                    case 'e':
                        if (i - lastThousandsSeparatorIndex < MIN_DISTANCE_BETWEEN_THOUSANDS_SEPARATOR)
                        {
                            return Double.NaN;
                        }
                        // append rest of strText and skip to end of loop
                        sb.Append(strText.Substring(i));
                        i = len;
                        break;
                    case '%':
                        foundPercentage = true;
                        break;
                    default:
                        // all other characters are illegal
                        return Double.NaN;
                }
            }
            if (!foundDecimalPoint)
            {
                if (i - lastThousandsSeparatorIndex < MIN_DISTANCE_BETWEEN_THOUSANDS_SEPARATOR)
                {
                    return Double.NaN;
                }
            }
            double d;
            try
            {
                d = Double.Parse(sb.ToString(), CultureInfo.InvariantCulture);
            }
            catch (FormatException)
            {
                // still a problem parsing the number - probably out of range
                return Double.NaN;
            }
            double result = foundUnaryMinus ? -d : d;
            return foundPercentage ? result / 100 : result;
        }
    }
}
