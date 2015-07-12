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
 * distributed Under the License is distributed on an "AS Is" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations Under the License.
 */

namespace NPOI.SS.Formula.Functions
{
    using System;
    using NPOI.SS.Formula.Eval;
    using System.Text;
    using System.Text.RegularExpressions;

    /// <summary>
    /// Implementation of the PROPER function:
    /// Normalizes all words (separated by non-word characters) by
    /// making the first letter upper and the rest lower case.
    /// </summary>
    public class Proper : SingleArgTextFunc
    {
        //Regex nonAlphabeticPattern = new Regex("\\P{IsL}");
        public override ValueEval Evaluate(String text)
        {
            StringBuilder sb = new StringBuilder();
            bool shouldMakeUppercase = true;
            String lowercaseText = text.ToLower();
            String uppercaseText = text.ToUpper();

            bool prevCharIsLetter = char.IsLetter(text[0]);
            sb.Append(uppercaseText[0]);

            for (int i = 1; i < text.Length; i++)
            {
                shouldMakeUppercase = !prevCharIsLetter;
                if (shouldMakeUppercase)
                {
                    sb.Append(uppercaseText[(i)]);
                }
                else
                {
                    sb.Append(lowercaseText[(i)]);
                }
                prevCharIsLetter = char.IsLetter(text[i]);
            }
            return new StringEval(sb.ToString());
        }
    }
}