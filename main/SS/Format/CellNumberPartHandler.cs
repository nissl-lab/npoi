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
namespace NPOI.SS.Format
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Text.RegularExpressions;

    /**
     * Internal helper class for CellNumberFormatter
     */

    public class CellNumberPartHandler : CellFormatPart.IPartHandler
    {
        private char insertSignForExponent;
        private double scale = 1;
        private CellNumberFormatter.Special decimalPoint;
        private CellNumberFormatter.Special slash;
        private CellNumberFormatter.Special exponent;
        private CellNumberFormatter.Special numerator;
        private List<CellNumberFormatter.Special> specials = new List<CellNumberFormatter.Special>();
        private bool improperFraction;

        public String HandlePart(Match m, String part, CellFormatType type, StringBuilder descBuf)
        {
            int pos = descBuf.Length;
            char firstCh = part[0];
            switch (firstCh)
            {
                case 'e':
                case 'E':
                    // See comment in WriteScientific -- exponent handling is complex.
                    // (1) When parsing the format, remove the sign from After the 'e' and
                    // Put it before the first digit of the exponent.
                    if (exponent == null && specials.Count > 0)
                    {
                        exponent = new CellNumberFormatter.Special('.', pos);
                        specials.Add(exponent);
                        insertSignForExponent = part[1];
                        return part.Substring(0, 1);
                    }
                    break;

                case '0':
                case '?':
                case '#':
                    if (insertSignForExponent != '\0')
                    {
                        specials.Add(new CellNumberFormatter.Special(insertSignForExponent, pos));
                        descBuf.Append(insertSignForExponent);
                        insertSignForExponent = '\0';
                        pos++;
                    }
                    for (int i = 0; i < part.Length; i++)
                    {
                        char ch = part[i];
                        specials.Add(new CellNumberFormatter.Special(ch, pos + i));
                    }
                    break;

                case '.':
                    if (decimalPoint == null && specials.Count > 0)
                    {
                        decimalPoint = new CellNumberFormatter.Special('.', pos);
                        specials.Add(decimalPoint);
                    }
                    break;

                case '/':
                    //!! This assumes there is a numerator and a denominator, but these are actually optional
                    if (slash == null && specials.Count > 0)
                    {
                        numerator = PreviousNumber();
                        // If the first number in the whole format is the numerator, the
                        // entire number should be printed as an improper fraction
                        improperFraction |= (numerator == FirstDigit(specials));
                        slash = new CellNumberFormatter.Special('.', pos);
                        specials.Add(slash);
                    }
                    break;

                case '%':
                    // don't need to remember because we don't need to do anything with these
                    scale *= 100;
                    break;

                default:
                    return null;
            }
            return part;
        }

        public double Scale
        {
            get
            {
                return scale;
            }
        }

        public CellNumberFormatter.Special DecimalPoint
        {
            get
            {
                return decimalPoint;
            }
        }

        public CellNumberFormatter.Special Slash
        {
            get
            {
                return slash;
            }
        }

        public CellNumberFormatter.Special Exponent
        {
            get
            {
                return exponent;
            }
        }

        public CellNumberFormatter.Special Numerator
        {
            get
            {
                return numerator;
            }
        }

        public List<CellNumberFormatter.Special> Specials
        {
            get
            {
                return specials;
            }
            
        }

        public bool IsImproperFraction
        {
            get
            {
                return improperFraction;
            }
        }

        private CellNumberFormatter.Special PreviousNumber()
        {
            for (int i = specials.Count - 1; i >= 0; i--)
            {
                CellNumberFormatter.Special s = specials[i];
                if (IsDigitFmt(s))
                {
                    //Special numStart = s;
                    CellNumberFormatter.Special last = s;
                    while (i >= 0)
                    {
                        s = specials[i];
                        if (last.pos - s.pos > 1 || !IsDigitFmt(s)) // it has to be continuous digits
                            break;
                        last = s;
                        i--;
                    }
                    return last;
                }
            }
            return null;
        }

        private static bool IsDigitFmt(CellNumberFormatter.Special s)
        {
            return s.ch == '0' || s.ch == '?' || s.ch == '#';
        }

        private static CellNumberFormatter.Special FirstDigit(List<CellNumberFormatter.Special> specials)
        {
            foreach (CellNumberFormatter.Special s in specials)
            {
                if (IsDigitFmt(s))
                {
                    return s;
                }
            }
            return null;
        }
    }

}