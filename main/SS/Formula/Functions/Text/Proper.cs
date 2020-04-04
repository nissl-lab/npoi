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
    using System.Globalization;

    /// <summary>
    /// Implementation of the PROPER function:
    /// Normalizes all words (separated by non-word characters) by
    /// making the first letter upper and the rest lower case.
    /// </summary>
    public class Proper : SingleArgTextFunc
    {
        public override ValueEval Evaluate(String text)
        {
            StringBuilder sb = new StringBuilder();
            bool shouldMakeUppercase = true;

            foreach (char ch in text.ToCharArray())
            {

                // Note: we are using String.toUpperCase() here on purpose as it handles certain things
                // better than Character.toUpperCase(), e.g. German "scharfes s" is translated
                // to "SS" (i.e. two characters), if uppercased properly!
                if (shouldMakeUppercase)
                {
                    sb.Append(ch.ToString().ToUpper(CultureInfo.CurrentCulture));
                }
                else
                {
                    sb.Append(ch.ToString().ToLower(CultureInfo.CurrentCulture));
                }
                shouldMakeUppercase = !char.IsLetter(ch);
            }
            return new StringEval(sb.ToString());
        }
    }
}