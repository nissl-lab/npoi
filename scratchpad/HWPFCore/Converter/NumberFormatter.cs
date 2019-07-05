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

using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HWPF.Converter
{
    public class NumberFormatter
    {

        private static String[] ENGLISH_LETTERS = new String[] { "a", "b",
            "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o",
            "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z" };

        private static String[] ROMAN_LETTERS = { "m", "cm", "d", "cd", "c",
            "xc", "l", "xl", "x", "ix", "v", "iv", "i" };

        private static int[] ROMAN_VALUES = { 1000, 900, 500, 400, 100, 90,
            50, 40, 10, 9, 5, 4, 1 };

        private const int T_ARABIC = 0;
        private const int T_LOWER_LETTER = 4;
        private const int T_LOWER_ROMAN = 2;
        private const int T_ORDINAL = 5;
        private const int T_UPPER_LETTER = 3;
        private const int T_UPPER_ROMAN = 1;

        public static String GetNumber(int num, int style)
        {
            switch (style)
            {
                case T_UPPER_ROMAN:
                    return ToRoman(num).ToUpper();
                case T_LOWER_ROMAN:
                    return ToRoman(num);
                case T_UPPER_LETTER:
                    return ToLetters(num).ToUpper();
                case T_LOWER_LETTER:
                    return ToLetters(num);
                case T_ARABIC:
                case T_ORDINAL:
                default:
                    return num.ToString();
            }
        }

        private static String ToLetters(int number)
        {
            int letterBase = 26;

            if (number <= 0)
                throw new ArgumentOutOfRangeException("Unsupported number: " + number);

            if (number < letterBase + 1)
                return ENGLISH_LETTERS[number - 1];

            long toProcess = number;

            StringBuilder stringBuilder = new StringBuilder();
            int maxPower = 0;
            {
                int boundary = 0;
                while (toProcess > boundary)
                {
                    maxPower++;
                    boundary = boundary * letterBase + letterBase;

                    if (boundary > int.MaxValue)
                        throw new ArgumentOutOfRangeException("Unsupported number: "
                                + toProcess);
                }
            }
            maxPower--;

            for (int p = maxPower; p > 0; p--)
            {
                long boundary = 0;
                long shift = 1;
                for (int i = 0; i < p; i++)
                {
                    shift *= letterBase;
                    boundary = boundary * letterBase + letterBase;
                }

                int count = 0;
                while (toProcess > boundary)
                {
                    count++;
                    toProcess -= shift;
                }
                stringBuilder.Append(ENGLISH_LETTERS[count - 1]);
            }
            stringBuilder.Append(ENGLISH_LETTERS[(int)toProcess - 1]);
            return stringBuilder.ToString();
        }

        private static String ToRoman(int number)
        {
            if (number <= 0)
                throw new ArgumentOutOfRangeException("Unsupported number: " + number);

            StringBuilder result = new StringBuilder();

            for (int i = 0; i < ROMAN_LETTERS.Length; i++)
            {
                String letter = ROMAN_LETTERS[i];
                int value = ROMAN_VALUES[i];
                while (number >= value)
                {
                    number -= value;
                    result.Append(letter);
                }
            }
            return result.ToString();
        }
    }

}
