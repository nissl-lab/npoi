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

    /**
     * <p>Some utils for Converting from and to any base</p>
     *
     * @author cedric dot walter @ gmail dot com
     */
    public class BaseNumberUtils
    {
        public static double ConvertToDecimal(String value, int base1, int maxNumberOfPlaces)
        {
            if (string.IsNullOrEmpty(value))
            {
                return 0.0;
            }

            long stringLength = value.Length;
            if (stringLength > maxNumberOfPlaces)
            {
                throw new ArgumentException();
            }

            double decimalValue = 0.0;

            long signedDigit = 0;
            bool hasSignedDigit = true;
            char[] characters = value.ToCharArray();
            foreach (char character in characters)
            {
                long digit;

                if ('0' <= character && character <= '9')
                {
                    digit = character - '0';
                }
                else if ('A' <= character && character <= 'Z')
                {
                    digit = 10 + (character - 'A');
                }
                else if ('a' <= character && character <= 'z')
                {
                    digit = 10 + (character - 'a');
                }
                else
                {
                    digit = base1;
                }

                if (digit < base1)
                {
                    if (hasSignedDigit)
                    {
                        hasSignedDigit = false;
                        signedDigit = digit;
                    }
                    decimalValue = decimalValue * base1 + digit;
                }
                else
                {
                    throw new ArgumentException("character not allowed");
                }
            }

            bool isNegative = (!hasSignedDigit && stringLength == maxNumberOfPlaces && (signedDigit >= base1 / 2));
            if (isNegative)
            {
                decimalValue = GetTwoComplement(base1, maxNumberOfPlaces, decimalValue);
                decimalValue = decimalValue * -1.0;
            }

            return decimalValue;
        }

        private static double GetTwoComplement(double base1, double maxNumberOfPlaces, double decimalValue)
        {
            return (Math.Pow(base1, maxNumberOfPlaces) - decimalValue);
        }
    }

}