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
    using System.Text.RegularExpressions;
    using System.Text;
    using System.Collections.Generic;
    using NPOI.SS.Util;
    using System.Collections;
    using NPOI.Util;


    /**
     * This class : printing out a value using a number format.
     *
     * @author Ken Arnold, Industrious Media LLC
     */
    public class CellNumberFormatter : CellFormatter
    {
        private String desc;
        private String printfFmt;
        private double scale;
        private Special decimalPoint;
        private Special slash;
        private Special exponent;
        private Special numerator;
        private Special afterInteger;
        private Special afterFractional;
        private bool integerCommas;
        private List<Special> specials = new List<Special>();
        private List<Special> integerSpecials = new List<Special>();
        private List<Special> fractionalSpecials = new List<Special>();
        private List<Special> numeratorSpecials = new List<Special>();
        private List<Special> denominatorSpecials = new List<Special>();
        private List<Special> exponentSpecials = new List<Special>();
        private List<Special> exponentDigitSpecials = new List<Special>();
        private int maxDenominator;
        private String numeratorFmt;
        private String denominatorFmt;
        private bool improperFraction;
        private DecimalFormat decimalFmt;
        private static List<Special> EmptySpecialList = new List<Special>();

        private static readonly CellFormatter SIMPLE_NUMBER = new SimpleNumberCellFormatter("General");
        private static readonly CellFormatter SIMPLE_INT = new CellNumberFormatter("#");
        private static readonly CellFormatter SIMPLE_FLOAT = new CellNumberFormatter("#.#");

        /// <summary>
        /// The CellNumberFormatter.simpleValue() method uses the SIMPLE_NUMBER
        /// CellFormatter defined here. The CellFormat.GENERAL_FORMAT CellFormat
        /// no longer uses the SIMPLE_NUMBER CellFormatter.
        /// Note that the simpleValue()/SIMPLE_NUMBER CellFormatter format
        /// ("#" for integer values, and "#.#" for floating-point values) is
        /// different from the 'General' format for numbers ("#" for integer
        /// values and "#.#########" for floating-point values).
        /// </summary>
        private sealed class GeneralNumberFormatter : CellFormatter
        {
            private GeneralNumberFormatter()
                    : base("General")
            {

            }

            public override void FormatValue(StringBuilder toAppendTo, Object value)
            {
                if (value == null)
                {
                    return;
                }

                CellFormatter cf;
                //if (value is Number)
                if (NPOI.Util.Number.IsNumber(value))
                {
                    double num;
                    double.TryParse(value.ToString(), out num);
                    cf = (num % 1.0 == 0) ? SIMPLE_INT : SIMPLE_FLOAT;
                }
                else
                {
                    cf = CellTextFormatter.SIMPLE_TEXT;
                }
                cf.FormatValue(toAppendTo, value);
            }

            public override void SimpleValue(StringBuilder toAppendTo, Object value)
            {
                FormatValue(toAppendTo, value);
            }
        }

        private sealed class SimpleNumberCellFormatter : CellFormatter
        {
            public SimpleNumberCellFormatter(string format)
                : base(format)
            {

            }
            public override void FormatValue(StringBuilder toAppendTo, Object value)
            {
                if (value == null)
                    return;
                //if (value is Number) {
                if (NPOI.Util.Number.IsNumber(value))
                {
                    double num;
                    double.TryParse(value.ToString(), out num);
                    if (num % 1.0 == 0)
                        SIMPLE_INT.FormatValue(toAppendTo, value);
                    else
                        SIMPLE_FLOAT.FormatValue(toAppendTo, value);
                }
                else
                {
                    CellTextFormatter.SIMPLE_TEXT.FormatValue(toAppendTo, value);
                }
            }
            public override void SimpleValue(StringBuilder toAppendTo, Object value)
            {
                FormatValue(toAppendTo, value);
            }
        }



        /**
         * This class is used to mark where the special characters in the format
         * are, as opposed to the other characters that are simply printed.
         */
        public class Special
        {
            internal char ch;
            internal int pos;

            public Special(char ch, int pos)
            {
                this.ch = ch;
                this.pos = pos;
            }


            public override String ToString()
            {
                return "'" + ch + "' @ " + pos;
            }
        }

        /**
         * Creates a new cell number formatter.
         *
         * @param format The format to Parse.
         */
        public CellNumberFormatter(String format)
            : base(format)
        {
            CellNumberPartHandler ph = new CellNumberPartHandler();
            StringBuilder descBuf = CellFormatPart.ParseFormat(format, CellFormatType.NUMBER, ph);

            exponent = ph.Exponent;
            specials.AddRange(ph.Specials);
            improperFraction = ph.IsImproperFraction;

            // These are inconsistent settings, so ditch 'em
            if ((ph.DecimalPoint != null || ph.Exponent != null) && ph.Slash != null)
            {
                slash = null;
                numerator = null;
            }
            else
            {
                slash = ph.Slash;
                numerator = ph.Numerator;
            }

            int precision = interpretPrecision(ph.DecimalPoint, specials);
            int fractionPartWidth = 0;
            if (ph.DecimalPoint != null)
            {
                fractionPartWidth = 1 + precision;
                if (precision == 0)
                {
                    // This means the format has a ".", but that output should have no decimals after it.
                    // We just stop treating it specially
                    specials.Remove(ph.DecimalPoint);
                    decimalPoint = null;
                }
                else
                {
                    decimalPoint = ph.DecimalPoint;
                }
            }
            else
            {
                decimalPoint = null;
            }

            if (decimalPoint != null)
            {
                afterInteger = decimalPoint;
            }
            else if (exponent != null)
            {
                afterInteger = exponent;
            }
            else if (numerator != null)
            {
                afterInteger = numerator;
            }
            else
            {
                afterInteger = null;
            }

            if (exponent != null)
            {
                afterFractional = exponent;
            }
            else if (numerator != null)
            {
                afterFractional = numerator;
            }
            else
            {
                afterFractional = null;
            }

            double[] scaleByRef = { ph.Scale };
            integerCommas = interpretIntegerCommas(descBuf, specials, decimalPoint, integerEnd(), fractionalEnd(), scaleByRef);
            if (exponent == null)
            {
                scale = scaleByRef[0];
            }
            else
            {
                // in "e" formats,% and trailing commas have no scaling effect
                scale = 1;
            }

            if (precision != 0)
            {
                // TODO: if decimalPoint is null (-> index == -1), return the whole list?
                int startIndex = specials.IndexOf(decimalPoint) + 1;
                fractionalSpecials.AddRange(specials.GetRange(startIndex, fractionalEnd() - startIndex));
            }

            if (exponent != null)
            {
                int exponentPos = specials.IndexOf(exponent);
                exponentSpecials.AddRange(specialsFor(exponentPos, 2));
                exponentDigitSpecials.AddRange(specialsFor(exponentPos + 2));
            }

            if (slash != null)
            {
                if (numerator != null)
                {
                    numeratorSpecials.AddRange(specialsFor(specials.IndexOf(numerator)));
                }

                denominatorSpecials.AddRange(specialsFor(specials.IndexOf(slash) + 1));
                if (denominatorSpecials.Count == 0)
                {
                    // no denominator follows the slash, drop the fraction idea
                    numeratorSpecials.Clear();
                    maxDenominator = 1;
                    numeratorFmt = null;
                    denominatorFmt = null;
                }
                else
                {
                    maxDenominator = maxValue(denominatorSpecials);
                    numeratorFmt = SingleNumberFormat(numeratorSpecials);
                    denominatorFmt = SingleNumberFormat(denominatorSpecials);
                }
            }
            else
            {
                maxDenominator = 1;
                numeratorFmt = null;
                denominatorFmt = null;
            }

            integerSpecials.AddRange(specials.GetRange(0, integerEnd()));

            if (exponent == null)
            {
                StringBuilder fmtBuf = new StringBuilder();

                int integerPartWidth = calculateintPartWidth();
                //int totalWidth = integerPartWidth + fractionPartWidth;

                fmtBuf.Append('0', integerPartWidth).Append('.').Append('0', precision);

                //fmtBuf.Append("f");
                printfFmt = fmtBuf.ToString();

                //this number format is legal in C#;
                //printfFmt = fmt;
                decimalFmt = null;
            }
            else
            {
                StringBuilder fmtBuf = new StringBuilder();
                bool first = true;
                List<Special> specialList = integerSpecials;
                if (integerSpecials.Count == 1)
                {
                    // If we don't do this, we Get ".6e5" instead of "6e4"
                    fmtBuf.Append("0");
                    first = false;
                }
                else
                    foreach (Special s in specialList)
                    {
                        if (IsDigitFmt(s))
                        {
                            fmtBuf.Append(first ? '#' : '0');
                            first = false;
                        }
                    }
                if (fractionalSpecials.Count > 0)
                {
                    fmtBuf.Append('.');
                    foreach (Special s in fractionalSpecials)
                    {
                        if (IsDigitFmt(s))
                        {
                            if (!first)
                                fmtBuf.Append('0');
                            first = false;
                        }
                    }
                }
                fmtBuf.Append('E');
                placeZeros(fmtBuf, exponentSpecials.GetRange(2, exponentSpecials.Count - 2));
                decimalFmt = new DecimalFormat(fmtBuf.ToString());

                printfFmt = null;
            }

            desc = descBuf.ToString();
        }

        private static void placeZeros(StringBuilder sb, List<Special> specials)
        {
            foreach (Special s in specials)
            {
                if (IsDigitFmt(s))
                    sb.Append('0');
            }
        }

        private static Special firstDigit(List<Special> specials)
        {
            foreach (Special s in specials)
            {
                if (IsDigitFmt(s))
                    return s;
            }
            return null;
        }

        static CellNumberStringMod insertMod(Special special, string toAdd, int where)
        {
            return new CellNumberStringMod(special, toAdd, where);
        }

        static CellNumberStringMod deleteMod(Special start, bool startInclusive, Special end, bool endInclusive)
        {

            return new CellNumberStringMod(start, startInclusive, end, endInclusive);
        }

        static CellNumberStringMod ReplaceMod(Special start, bool startInclusive,
                Special end, bool endInclusive, char withChar)
        {
            return new CellNumberStringMod(start, startInclusive, end, endInclusive,
                    withChar);
        }

        private static String SingleNumberFormat(List<Special> numSpecials)
        {
            //return "%0" + numSpecials.Count + "d";
            return "D" + numSpecials.Count;
        }

        private static int maxValue(List<Special> s)
        {
            return (int)Math.Round(Math.Pow(10, s.Count) - 1);
        }

        private List<Special> specialsFor(int pos, int takeFirst)
        {
            if (pos >= specials.Count)
                return EmptySpecialList;
            IEnumerator<Special> it = specials.GetRange(pos + takeFirst, specials.Count - pos - takeFirst).GetEnumerator();
            //.ListIterator(pos + takeFirst);
            it.MoveNext();
            Special last = it.Current;
            int end = pos + takeFirst;
            while (it.MoveNext())
            {
                Special s = it.Current;
                if (!IsDigitFmt(s) || s.pos - last.pos > 1)
                    break;
                end++;
                last = s;
            }
            return specials.GetRange(pos, end + 1 - pos);
        }

        private List<Special> specialsFor(int pos)
        {
            return specialsFor(pos, 0);
        }

        private static bool IsDigitFmt(Special s)
        {
            return s.ch == '0' || s.ch == '?' || s.ch == '#';
        }


        private int calculateintPartWidth()
        {
            int digitCount = 0;
            foreach (Special s in specials)
            {
                //!! Handle fractions: The previous set of digits before that is the numerator, so we should stop short of that
                if (s == afterInteger)
                {
                    break;
                }
                else if (IsDigitFmt(s))
                {
                    digitCount++;
                }
            }
            return digitCount;
        }

        private static int interpretPrecision(Special decimalPoint, List<Special> specials)
        {
            int idx = specials.IndexOf(decimalPoint);
            int precision = 0;
            if (idx != -1)
            {
                // skip over the decimal point itself
                IEnumerator<Special> it = specials.GetRange(idx + 1, specials.Count - idx - 1).GetEnumerator();//.ListIterator(specials.IndexOf(decimalPoint));
                //if (it.HasNext())
                //     it.Next();  // skip over the decimal point itself
                while (it.MoveNext())
                {
                    Special s = it.Current;
                    if (!IsDigitFmt(s))
                    {
                        break;
                    }
                    precision++;
                }
            }
            return precision;
        }

        private static bool interpretIntegerCommas(StringBuilder sb, List<Special> specials, Special decimalPoint, int integerEnd, int fractionalEnd, double[] scale)
        {
            // In the integer part, commas at the end are scaling commas; other commas mean to show thousand-grouping commas
            List<Special> list = specials.GetRange(0, integerEnd);//.ListIterator();

            bool stillScaling = true;
            bool integerCommas = false;
            //while (it.HasPrevious())
            for (int i = list.Count - 1; i >= 0; i--)
            {
                Special s = list[i];
                if (s.ch != ',')
                {
                    stillScaling = false;
                }
                else
                {
                    if (stillScaling)
                    {
                        scale[0] /= 1000;
                    }
                    else
                    {
                        integerCommas = true;
                    }
                }
            }

            if (decimalPoint != null)
            {
                list = specials.GetRange(0, fractionalEnd);//.ListIterator(fractionalEnd());
                //while (it.HasPrevious())
                for (int i = list.Count - 1; i >= 0; i--)
                {
                    Special s = list[i];
                    if (s.ch != ',')
                    {
                        break;
                    }
                    else
                    {
                        scale[0] /= 1000;
                    }
                }
            }

            // Now strip them out -- we only need their interpretation, not their presence
            IEnumerator<Special> it = specials.GetEnumerator();
            int Removed = 0;
            List<Special> toRemove = new List<Special>();
            while (it.MoveNext())
            {
                Special s = it.Current;
                s.pos -= Removed;
                if (s.ch == ',')
                {
                    Removed++;
                    //it.Remove();
                    toRemove.Add(s);
                    sb.Remove(s.pos, 1);
                }
            }
            foreach (Special e in toRemove)
            {
                specials.Remove(e);
            }

            return integerCommas;
        }

        private int integerEnd()
        {
            return (afterInteger == null) ? specials.Count : specials.IndexOf(afterInteger);
        }

        private int fractionalEnd()
        {
            return (afterFractional == null) ? specials.Count : specials.IndexOf(afterFractional);
        }

        /** {@inheritDoc} */
        public override void FormatValue(StringBuilder toAppendTo, Object valueObject)
        {
            double value = ((double)valueObject);
            value *= scale;

            // For negative numbers:
            // - If the cell format has a negative number format, this method
            // is called with a positive value and the number format has
            // the negative formatting required, e.g. minus sign or brackets.
            // - If the cell format does not have a negative number format,
            // this method is called with a negative value and the number is
            // formatted with a minus sign at the start.
            bool negative = value < 0;
            if (negative)
                value = -value;

            // Split out the fractional part if we need to print a fraction
            double fractional = 0;
            if (slash != null)
            {
                if (improperFraction)
                {
                    fractional = value;
                    value = 0;
                }
                else
                {
                    fractional = value % 1.0;
                    //noinspection SillyAssignment
                    value = (long)value;
                }
            }

            SortedList<CellNumberStringMod, object> mods = new SortedList<CellNumberStringMod, object>();
            StringBuilder output = new StringBuilder(desc);

            if (exponent != null)
            {
                WriteScientific(value, output, mods);
            }
            else if (improperFraction)
            {
                WriteFraction(value, null, fractional, output, mods);
            }
            else
            {
                StringBuilder result = new StringBuilder();
                //Formatter f = new Formatter(result);
                //f.Format(LOCALE, printfFmt, value);
                result.Append(value.ToString(printfFmt));
                if (numerator == null)
                {
                    WriteFractional(result, output);
                    Writeint(result, output, integerSpecials, mods, integerCommas);
                }
                else
                {
                    WriteFraction(value, result, fractional, output, mods);
                }
            }

            // Now strip out any remaining '#'s and add any pending text ...
            IEnumerator<Special> it = specials.GetEnumerator();//.ListIterator();
            IEnumerator Changes = mods.Keys.GetEnumerator();
            CellNumberStringMod nextChange = (Changes.MoveNext() ? (CellNumberStringMod)Changes.Current : null);
            int adjust = 0;
            BitArray deletedChars = new BitArray(1024); // records chars already deleted
            foreach (Special s in specials)
            {
                int adjustedPos = s.pos + adjust;
                if (!deletedChars[(s.pos)] && output[adjustedPos] == '#')
                {
                    output.Remove(adjustedPos, 1);
                    adjust--;
                    deletedChars.Set(s.pos, true);
                }
                while (nextChange != null && s == nextChange.GetSpecial())
                {
                    int lenBefore = output.Length;
                    int modPos = s.pos + adjust;

                    switch (nextChange.Op)
                    {
                        case CellNumberStringMod.AFTER:
                            // ignore Adding a comma After a deleted char (which was a '#')
                            if (nextChange.ToAdd.Equals(",") && deletedChars.Get(s.pos))
                                break;
                            output.Insert(modPos + 1, nextChange.ToAdd);
                            break;
                        case CellNumberStringMod.BEFORE:
                            output.Insert(modPos, nextChange.ToAdd);
                            break;

                        case CellNumberStringMod.REPLACE:
                            // delete starting pos in original coordinates
                            int delPos = s.pos;
                            if (!nextChange.IsStartInclusive)
                            {
                                delPos++;
                                modPos++;
                            }

                            // Skip over anything already deleted
                            while (deletedChars.Get(delPos))
                            {
                                delPos++;
                                modPos++;
                            }

                            // delete end point in original
                            int delEndPos = nextChange.End.pos;
                            if (nextChange.IsEndInclusive)
                                delEndPos++;

                            // delete end point in current
                            int modEndPos = delEndPos + adjust;

                            if (modPos < modEndPos)
                            {
                                if (nextChange.ToAdd == "")
                                    output.Remove(modPos, modEndPos - modPos);
                                else
                                {
                                    char fillCh = nextChange.ToAdd[0];
                                    for (int i = modPos; i < modEndPos; i++)
                                        output[i] = fillCh;
                                }
                                for (int k = delPos; k < delEndPos; k++)
                                    deletedChars.Set(k, true);
                                //deletedChars.Set(delPos, delEndPos);
                            }
                            break;

                        default:
                            throw new InvalidOperationException("Unknown op: " + nextChange.Op);
                    }
                    adjust += output.Length - lenBefore;

                    if (Changes.MoveNext())
                        nextChange = (CellNumberStringMod)Changes.Current;
                    else
                        nextChange = null;
                }
            }

            // Finally, add it to the string
            if (negative)
                toAppendTo.Append('-');
            toAppendTo.Append(output);
        }

        private void WriteScientific(double value, StringBuilder output, SortedList<CellNumberStringMod, object> mods)
        {

            StringBuilder result = new StringBuilder();
            //FieldPosition fractionPos = new FieldPosition(DecimalFormat.FRACTION_FIELD;
            //decimalFmt.Format(value, result, fractionPos);

            //
            string pattern = decimalFmt.Pattern;
            int pos = 0;
            while (true)
            {
                if (pattern[pos] == '#' || pattern[pos] == '0')
                {
                    pos++;
                }
                else
                    break;
            }
            int integerNum = pos;
            if (pattern[0] == '#')
                integerNum--;
            if (integerNum >= 6 && value > 1)
            {
                pattern = pattern.Substring(1);
                result.Append(value.ToString(pattern));
            }
            else
            {
                result.Append(value.ToString("E"));
            }

            Writeint(result, output, integerSpecials, mods, integerCommas);
            WriteFractional(result, output);

            /*
            * Exponent sign handling is complex.
            *
            * In DecimalFormat, you never Put the sign in the format, and the sign only
            * comes out of the format if it is negative.
            *
            * In Excel, you always say whether to always show the sign ("e+") or only
            * show negative signs ("e-").
            *
            * Also in Excel, where you Put the sign in the format is NOT where it comes
            * out in the result.  In the format, the sign goes with the "e"; in the
            * output it goes with the exponent value.  That is, if you say "#e-|#" you
            * Get "1e|-5", not "1e-|5". This Makes sense I suppose, but it complicates
            * things.
            *
            * Finally, everything else in this formatting code assumes that the base of
            * the result is the original format, and that starting from that situation,
            * the indexes of the original special characters can be used to place the new
            * characters.  As just described, this is not true for the exponent's sign.
            * <p/>
            * So here is how we handle it:
            *
            * (1) When parsing the format, remove the sign from After the 'e' and Put it
            * before the first digit of the exponent (where it will be Shown).
            *
            * (2) Determine the result's sign.
            *
            * (3) If it's missing, Put the sign into the output to keep the result
            * lined up with the output. (In the result, "after the 'e'" and "before the
            * first digit" are the same because the result has no extra chars to be in
            * the way.)
            *
            * (4) In the output, remove the sign if it should not be Shown ("e-" was used
            * and the sign is negative) or Set it to the correct value.
            */

            // (2) Determine the result's sign.
            string tmp = result.ToString();
            int ePos = tmp.IndexOf("E");// fractionPos.EndIndex;
            int signPos = ePos + 1;
            char expSignRes = result[signPos];

            if (expSignRes != '-')
            {
                // not a sign, so it's a digit, and therefore a positive exponent
                expSignRes = '+';
                // (3) If it's missing, Put the sign into the output to keep the result
                // lined up with the output.
                if (tmp.IndexOf(expSignRes, ePos) < 0)
                    result.Insert(signPos, '+');
            }

            // Now the result lines up like it is supposed to with the specials' indexes
            IEnumerator<Special> it = exponentSpecials.GetEnumerator();//.ListIterator(1);
            it.MoveNext();
            it.MoveNext();
            Special expSign = it.Current;//.Next();
            char expSignFmt = expSign.ch;

            // (4) In the output, remove the sign if it should not be Shown or Set it to
            // the correct value.
            if (expSignRes == '-' || expSignFmt == '+')
                mods.Add(ReplaceMod(expSign, true, expSign, true, expSignRes), null);
            else
                mods.Add(deleteMod(expSign, true, expSign, true), null);

            StringBuilder exponentNum = new StringBuilder(result.ToString().Substring(signPos + 1));
            if (exponentNum.Length > 2 && exponentNum[0] == '0')
                exponentNum.Remove(0, 1);
            Writeint(exponentNum, output, exponentDigitSpecials, mods, false);
        }

        private void WriteFraction(double value, StringBuilder result,
                double fractional, StringBuilder output, SortedList<CellNumberStringMod, object> mods)
        {

            // Figure out if we are to suppress either the integer or fractional part.
            // With # the suppressed part is Removed; with ? it is Replaced with spaces.
            if (!improperFraction)
            {
                // If fractional part is zero, and numerator doesn't have '0', write out
                // only the integer part and strip the rest.
                if (fractional == 0 && !HasChar('0', numeratorSpecials))
                {
                    Writeint(result, output, integerSpecials, mods, false);

                    Special start = lastSpecial(integerSpecials);
                    Special end = lastSpecial(denominatorSpecials);
                    if (HasChar('?', integerSpecials, numeratorSpecials, denominatorSpecials))
                    {
                        //if any format has '?', then replace the fraction with spaces
                        mods.Add(ReplaceMod(start, false, end, true, ' '), null);
                    }
                    else
                    {
                        // otherwise, remove the fraction
                        mods.Add(deleteMod(start, false, end, true), null);
                    }

                    // That's all, just return
                    return;
                }
                else
                {
                    // New we check to see if we should remove the integer part
                    bool numNoZero = !HasChar('0', numeratorSpecials);
                    bool intNoZero = !HasChar('0', integerSpecials);
                    bool intOnlyHash = integerSpecials.Count == 0 || (integerSpecials.Count == 1 && HasChar('#', integerSpecials));

                    bool removeBecauseZero = fractional == 0 && (intOnlyHash || numNoZero);
                    bool removeBecauseFraction = fractional != 0 && intNoZero;

                    if (value == 0 && (removeBecauseZero || removeBecauseFraction))
                    {
                        Special start = lastSpecial(integerSpecials);
                        bool hasPlaceHolder = HasChar('?', integerSpecials, numeratorSpecials);
                        CellNumberStringMod sm = hasPlaceHolder
                            ? ReplaceMod(start, true, numerator, false, ' ')
                            : deleteMod(start, true, numerator, false);
                        mods.Add(sm, null);
                    }
                    else
                    {
                        // Not removing the integer part -- print it out
                        Writeint(result, output, integerSpecials, mods, false);
                    }
                }
            }

            // Calculate and print the actual fraction (improper or otherwise)
            try
            {
                int n;
                int d;
                // the "fractional % 1" captures integer values in improper fractions
                if (fractional == 0 || (improperFraction && fractional % 1 == 0))
                {
                    // 0 as a fraction is reported by excel as 0/1
                    n = (int)Math.Round(fractional);
                    d = 1;
                }
                else
                {
                    SimpleFraction frac = SimpleFraction.BuildFractionMaxDenominator(fractional, maxDenominator);
                    n = frac.Numerator;
                    d = frac.Denominator;
                }
                if (improperFraction)
                    n += (int)Math.Round(value * d);
                WriteSingleint(numeratorFmt, n, output, numeratorSpecials, mods);
                WriteSingleint(denominatorFmt, d, output, denominatorSpecials, mods);
            }
            catch (Exception ignored)
            {
                //ignored.PrintStackTrace();
                System.Console.WriteLine(ignored.StackTrace);
            }
        }
        //private static bool HasChar(char ch, List<Special> s1, List<Special> s2, List<Special> s3)
        //{
        //    return HasChar(ch, s1) || HasChar(ch, s2) || HasChar(ch, s3);
        //}
        //private static bool HasChar(char ch, List<Special> s1, List<Special> s2)
        //{
        //    return HasChar(ch, s1) || HasChar(ch, s2);
        //}
        private static bool HasChar(char ch, params List<Special>[] numSpecials)
        {
            foreach (List<Special> specials in numSpecials)
            {
                foreach (Special s in specials)
                {
                    if (s.ch == ch)
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private void WriteSingleint(String fmt, int num, StringBuilder output,
                List<Special> numSpecials, SortedList<CellNumberStringMod, object> mods)
        {

            StringBuilder sb = new StringBuilder();
            //Formatter formatter = new Formatter(sb);
            //formatter.Format(LOCALE, fmt, num);
            sb.Append(num.ToString(fmt));
            Writeint(sb, output, numSpecials, mods, false);
        }

        private void Writeint(StringBuilder result, StringBuilder output,
                List<Special> numSpecials, SortedList<CellNumberStringMod, object> mods,
                bool ShowCommas)
        {

            int pos = result.ToString().IndexOf(".") - 1;
            if (pos < 0)
            {
                if (exponent != null && numSpecials == integerSpecials)
                    pos = result.ToString().IndexOf("E") - 1;
                else
                    pos = result.Length - 1;
            }

            int strip;
            for (strip = 0; strip < pos; strip++)
            {
                char resultCh = result[strip];
                if (resultCh != '0' && resultCh != ',')
                    break;
            }

            //ListIterator<Special> it = numSpecials.ListIterator(numSpecials.Count);
            bool followWithComma = false;
            Special lastOutputintDigit = null;
            int digit = 0;
            //while (it.HasPrevious()) {
            for (int i = numSpecials.Count - 1; i >= 0; i--)
            {
                char resultCh;
                if (pos >= 0)
                    resultCh = result[pos];
                else
                {
                    // If result is shorter than field, pretend there are leading zeros
                    resultCh = '0';
                }
                Special s = numSpecials[i];
                followWithComma = ShowCommas && digit > 0 && digit % 3 == 0;
                bool zeroStrip = false;
                if (resultCh != '0' || s.ch == '0' || s.ch == '?' || pos >= strip)
                {
                    zeroStrip = (s.ch == '?' && pos < strip);
                    output[s.pos] = (zeroStrip ? ' ' : resultCh);
                    lastOutputintDigit = s;
                }
                if (followWithComma)
                {
                    mods.Add(insertMod(s, zeroStrip ? " " : ",", CellNumberStringMod.AFTER), null);
                    followWithComma = false;
                }
                digit++;
                --pos;
            }
            StringBuilder extraLeadingDigits = new StringBuilder();
            if (pos >= 0)
            {
                // We ran out of places to Put digits before we ran out of digits; Put this aside so we can add it later
                // pos was decremented at the end of the loop above when the iterator was at its end
                ++pos;
                extraLeadingDigits = new StringBuilder(result.ToString().Substring(0, pos));
                if (ShowCommas)
                {
                    while (pos > 0)
                    {
                        if (digit > 0 && digit % 3 == 0)
                            extraLeadingDigits.Insert(pos, ',');
                        digit++;
                        --pos;
                    }
                }
                mods.Add(insertMod(lastOutputintDigit, extraLeadingDigits.ToString(), CellNumberStringMod.BEFORE), null);
            }
        }

        private void WriteFractional(StringBuilder result, StringBuilder output)
        {
            int digit;
            int strip;
            if (fractionalSpecials.Count > 0)
            {
                digit = result.ToString().IndexOf(".") + 1;
                if (exponent != null)
                    strip = result.ToString().IndexOf("E") - 1;
                else
                    strip = result.Length - 1;
                while (strip > digit && result[strip] == '0')
                    strip--;
                
                foreach (Special s in fractionalSpecials)
                {
                    if (digit >= result.Length)
                        break;
                    char resultCh = result[digit];
                    if (resultCh != '0' || s.ch == '0' || digit < strip)
                        output[s.pos] = resultCh;
                    else if (s.ch == '?')
                    {
                        // This is when we're in trailing zeros, and the format is '?'.
                        // We still strip out remaining '#'s later
                        output[s.pos] = ' ';
                    }
                    digit++;
                }
            }
        }

        /**
         * {@inheritDoc}
         * <p/>
         * For a number, this is <tt>"#"</tt> for integer values, and <tt>"#.#"</tt>
         * for floating-point values.
         */
        public override void SimpleValue(StringBuilder toAppendTo, Object value)
        {
            SIMPLE_NUMBER.FormatValue(toAppendTo, value);
        }

        private static Special lastSpecial(List<Special> s)
        {
            return s[s.Count - 1];
        }
    }
}