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

    using NPOI.HSSF.Util;
    using System.Collections.Generic;
    using System.Drawing;
    using System.Collections;
    using System.Text.RegularExpressions;
    using System.Text;

    /**
     * Objects of this class represent a single part of a cell format expression.
     * Each cell can have up to four of these for positive, zero, negative, and text
     * values.
     * <p/>
     * Each format part can contain a color, a condition, and will always contain a
     * format specification.  For example <tt>"[Red][>=10]#"</tt> has a color
     * (<tt>[Red]</tt>), a condition (<tt>>=10</tt>) and a format specification
     * (<tt>#</tt>).
     * <p/>
     * This class also Contains patterns for matching the subparts of format
     * specification.  These are used internally, but are made public in case other
     * code has use for them.
     *
     * @author Ken Arnold, Industrious Media LLC
     */
    public class CellFormatPart
    {
        private Color color;
        private CellFormatCondition condition;
        private CellFormatter format;
        private CellFormatType type;
        private static Dictionary<String, Color> NAMED_COLORS;
        public static IEqualityComparer<String> CASE_INSENSITIVE_ORDER
                                             = new CaseInsensitiveComparator();
        private class CaseInsensitiveComparator : IEqualityComparer<String>
        {
            // use serialVersionUID from JDK 1.2.2 for interoperability
            //private const long serialVersionUID = 8575799808933029326L;



            #region IEqualityComparer<string>

            public bool Equals(string x, string y)
            {
                return x.Equals(y, StringComparison.InvariantCultureIgnoreCase);
            }

            public int GetHashCode(string obj)
            {
                return obj.GetHashCode();
            }

            #endregion
        }
        static CellFormatPart()
        {
            NAMED_COLORS = new Dictionary<String, Color>(CASE_INSENSITIVE_ORDER);

            var colors = HSSFColor.GetIndexHash();
            foreach (object v in colors.Values)
            {
                HSSFColor hc = (HSSFColor)v;
                Type type = hc.GetType();
                String name = type.Name;
                if (name.Equals(name.ToUpper()))
                {
                    byte[] rgb = hc.RGB;
                    Color c = Color.FromArgb(rgb[0], rgb[1], rgb[2]);
                    if (!NAMED_COLORS.ContainsKey(name))
                    {
                        NAMED_COLORS.Add(name, c);
                    }
                    if (name.IndexOf('_') > 0)
                    {
                        if (!NAMED_COLORS.ContainsKey(name.Replace('_', ' ')))
                        {
                            NAMED_COLORS.Add(name.Replace('_', ' '), c);
                        }
                    }
                    if (name.IndexOf("_PERCENT") > 0)
                    {
                        if (!NAMED_COLORS.ContainsKey(name.Replace("_PERCENT", "%").Replace('_', ' ')))
                        {
                            NAMED_COLORS.Add(name.Replace("_PERCENT", "%").Replace('_', ' '), c);
                        }
                    }
                }
            }
            // A condition specification
            String condition = "([<>=]=?|!=|<>)    # The operator\n" +
                    "  \\s*([0-9]+(?:\\.[0-9]*)?)\\s*  # The constant to test against\n";

            // A currency symbol / string, in a specific locale
            String currency = "(\\[\\$.{0,3}-[0-9a-f]{3}\\])";

            String color =
                    "\\[(black|blue|cyan|green|magenta|red|white|yellow|color [0-9]+)\\]";

            // A number specification
            // Note: careful that in something like ##, that the trailing comma is not caught up in the integer part

            // A part of a specification
            String part = "\\\\.                 # Quoted single character\n" +
                    "|\"([^\\\\\"]|\\\\.)*\"         # Quoted string of characters (handles escaped quotes like \\\") \n" +
                    "|" + currency + "               # Currency symbol in a given locale\n" +
                    "|_.                             # Space as wide as a given character\n" +
                    "|\\*.                           # Repeating fill character\n" +
                    "|@                              # Text: cell text\n" +
                    "|([0?\\#](?:[0?\\#,]*))         # Number: digit + other digits and commas\n" +
                    "|e[-+]                          # Number: Scientific: Exponent\n" +
                    "|m{1,5}                         # Date: month or minute spec\n" +
                    "|d{1,4}                         # Date: day/date spec\n" +
                    "|y{2,4}                         # Date: year spec\n" +
                    "|h{1,2}                         # Date: hour spec\n" +
                    "|s{1,2}                         # Date: second spec\n" +
                    "|am?/pm?                        # Date: am/pm spec\n" +
                    "|\\[h{1,2}\\]                   # Elapsed time: hour spec\n" +
                    "|\\[m{1,2}\\]                   # Elapsed time: minute spec\n" +
                    "|\\[s{1,2}\\]                   # Elapsed time: second spec\n" +
                    "|[^;]                           # A character\n" + "";

            String format = "(?:" + color + ")?                  # Text color\n" +
                    "(?:\\[" + condition + "\\])?                # Condition\n" +
                    // see https://msdn.microsoft.com/en-ca/goglobal/bb964664.aspx and https://bz.apache.org/ooo/show_bug.cgi?id=70003
                    // we ignore these for now though
                    "(?:\\[\\$-[0-9a-fA-F]+\\])?                # Optional locale id, ignored currently\n" +
                    "((?:" + part + ")+)                        # Format spec\n";

            RegexOptions flags = RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace | RegexOptions.Compiled;
            COLOR_PAT = new Regex(color, flags);
            CONDITION_PAT = new Regex(condition, flags);
            SPECIFICATION_PAT = new Regex(part, flags);
            CURRENCY_PAT = new Regex(currency, flags);
            FORMAT_PAT = new Regex(format, flags);

            // Calculate the group numbers of important groups.  (They shift around
            // when the pattern is Changed; this way we figure out the numbers by
            // experimentation.)

            COLOR_GROUP = FindGroup(FORMAT_PAT, "[Blue]@", "Blue");
            CONDITION_OPERATOR_GROUP = FindGroup(FORMAT_PAT, "[>=1]@", ">=");
            CONDITION_VALUE_GROUP = FindGroup(FORMAT_PAT, "[>=1]@", "1");
            SPECIFICATION_GROUP = FindGroup(FORMAT_PAT, "[Blue][>1]\\a ?", "\\a ?");
        }

        /** Pattern for the color part of a cell format part. */
        public static Regex COLOR_PAT;
        /** Pattern for the condition part of a cell format part. */
        public static Regex CONDITION_PAT;
        /** Pattern for the format specification part of a cell format part. */
        public static Regex SPECIFICATION_PAT;
        /** Pattern for the currency symbol part of a cell format part */
        public static Regex CURRENCY_PAT;
        /** Pattern for an entire cell single part. */
        public static Regex FORMAT_PAT;

        /** Within {@link #FORMAT_PAT}, the group number for the matched color. */
        public static int COLOR_GROUP;
        /**
         * Within {@link #FORMAT_PAT}, the group number for the operator in the
         * condition.
         */
        public static int CONDITION_OPERATOR_GROUP;
        /**
         * Within {@link #FORMAT_PAT}, the group number for the value in the
         * condition.
         */
        public static int CONDITION_VALUE_GROUP;
        /**
         * Within {@link #FORMAT_PAT}, the group number for the format
         * specification.
         */
        public static int SPECIFICATION_GROUP;


        public interface IPartHandler
        {
            String HandlePart(Match m, String part, CellFormatType type,
                    StringBuilder desc);
        }

        /**
         * Create an object to represent a format part.
         *
         * @param desc The string to Parse.
         */
        public CellFormatPart(String desc)
        {
            Match m = FORMAT_PAT.Match(desc);
            if (!m.Success)
            {
                throw new ArgumentException("Unrecognized format: " + "\"" + desc + "\"");
            }
            color = GetColor(m);
            condition = GetCondition(m);
            type = GetCellFormatType(m);
            format = GetFormatter(m);
        }

        /**
         * Returns <tt>true</tt> if this format part applies to the given value. If
         * the value is a number and this is part has a condition, returns
         * <tt>true</tt> only if the number passes the condition.  Otherwise, this
         * allways return <tt>true</tt>.
         *
         * @param valueObject The value to Evaluate.
         *
         * @return <tt>true</tt> if this format part applies to the given value.
         */
        public bool Applies(Object valueObject)
        {
            if (condition == null || !(valueObject.GetType().IsPrimitive))
            {
                if (valueObject == null)
                    throw new NullReferenceException("valueObject");
                return true;
            }
            else
            {
                double num = (double)valueObject;
                return condition.Pass(num);
            }
        }

        /**
         * Returns the number of the first group that is the same as the marker
         * string.  Starts from group 1.
         *
         * @param pat    The pattern to use.
         * @param str    The string to match against the pattern.
         * @param marker The marker value to find the group of.
         *
         * @return The matching group number.
         *
         * @throws ArgumentException No group matches the marker.
         */
        private static int FindGroup(Regex pat, String str, String marker)
        {
            Match m = pat.Match(str);
            if (!m.Success)
                throw new ArgumentException(
                        "Pattern \"" + pat.ToString() + "\" doesn't match \"" + str +
                                "\"");
            for (int i = 1; i <= m.Groups.Count; i++)
            {
                String grp = m.Groups[i].Value;
                if (grp != null && grp.Equals(marker))
                    return i;
            }
            throw new ArgumentException(
                    "\"" + marker + "\" not found in \"" + pat.ToString() + "\"");
        }

        /**
         * Returns the color specification from the matcher, or <tt>null</tt> if
         * there is none.
         *
         * @param m The matcher for the format part.
         *
         * @return The color specification or <tt>null</tt>.
         */
        private static Color GetColor(Match m)
        {
            String cdesc = m.Groups[(COLOR_GROUP)].Value.ToUpper();
            if (cdesc == null || cdesc.Length == 0)
                return Color.Empty;
            Color c = Color.Empty;
            if (NAMED_COLORS.ContainsKey(cdesc))
                c = NAMED_COLORS[(cdesc)];
            //if (c == null)
            //    logger.Warning("Unknown color: " + quote(cdesc));
            return c;
        }

        /**
         * Returns the condition specification from the matcher, or <tt>null</tt> if
         * there is none.
         *
         * @param m The matcher for the format part.
         *
         * @return The condition specification or <tt>null</tt>.
         */
        private CellFormatCondition GetCondition(Match m)
        {
            String mdesc = m.Groups[(CONDITION_OPERATOR_GROUP)].Value;
            if (mdesc == null || mdesc.Length == 0)
                return null;
            return CellFormatCondition.GetInstance(m.Groups[(
                    CONDITION_OPERATOR_GROUP)].Value, m.Groups[(CONDITION_VALUE_GROUP)].Value);
        }
        /**
         * Returns the CellFormatType object implied by the format specification for
         * the format part.
         *
         * @param matcher The matcher for the format part.
         *
         * @return The CellFormatType.
         */
        private CellFormatType GetCellFormatType(Match matcher)
        {
            String fdesc = matcher.Groups[SPECIFICATION_GROUP].Value;
            return formatType(fdesc);
        }
        /**
         * Returns the formatter object implied by the format specification for the
         * format part.
         *
         * @param matcher The matcher for the format part.
         *
         * @return The formatter.
         */
        private CellFormatter GetFormatter(Match matcher)
        {
            String fdesc = matcher.Groups[(SPECIFICATION_GROUP)].Value;
            // For now, we don't support localised currencies, so simplify if there
            Match currencyM = CURRENCY_PAT.Match(fdesc);
            if (currencyM.Success)
            {
                String currencyPart = currencyM.Groups[(1)].Value;
                String currencyRepl;
                if (currencyPart.StartsWith("[$-"))
                {
                    // Default $ in a different locale
                    currencyRepl = "$";
                }
                else
                {
                    currencyRepl = currencyPart.Substring(2, currencyPart.LastIndexOf('-') - 2);
                }
                fdesc = fdesc.Replace(currencyPart, currencyRepl);
            }

            // Build a formatter for this simplified string
            return type.Formatter(fdesc);
        }

        /**
         * Returns the type of format.
         *
         * @param fdesc The format specification
         *
         * @return The type of format.
         */
        private CellFormatType formatType(String fdesc)
        {
            fdesc = fdesc.Trim();
            if (fdesc.Equals("") || fdesc.Equals("General", StringComparison.InvariantCultureIgnoreCase))
                return CellFormatType.GENERAL;

            MatchCollection mc = SPECIFICATION_PAT.Matches(fdesc);
            bool couldBeDate = false;
            bool seenZero = false;
            foreach(Match m in mc)
            //while (m.Success)
            {
                String repl = m.Groups[(0)].Value;
                if (repl.Length > 0)
                {
                    //switch (repl[0])
                    char c1 = repl[0];
                    char c2 = '\0';
                    if (repl.Length > 1)
                        c2 = Char.ToLower(repl[(1)]);

                    switch (c1)
                    {
                        case '@':
                            return CellFormatType.TEXT;
                        case 'd':
                        case 'D':
                        case 'y':
                        case 'Y':
                            return CellFormatType.DATE;
                        case 'h':
                        case 'H':
                        case 'm':
                        case 'M':
                        case 's':
                        case 'S':
                            // These can be part of date, or elapsed
                            couldBeDate = true;
                            break;
                        case '0':
                            // This can be part of date, elapsed, or number
                            seenZero = true;
                            break;
                        case '[':
                            if (c2 == 'h' || c2 == 'm' || c2 == 's')
                            {
                                return CellFormatType.ELAPSED;
                            }
                            if (c2 == '$')
                            {
                                // Localised currency
                                return CellFormatType.NUMBER;
                            }
                            // Something else inside [] which isn't supported!
                            throw new ArgumentException("Unsupported [] format block '" +
                                                               repl + "' in '" + fdesc + "' with c2: " + c2);
                        case '#':
                        case '?':
                            return CellFormatType.NUMBER;
                    }
                }
            }

            // Nothing defInitive was found, so we figure out it deductively
            if (couldBeDate)
                return CellFormatType.DATE;
            if (seenZero)
                return CellFormatType.NUMBER;
            return CellFormatType.TEXT;
        }

        /**
         * Returns a version of the original string that has any special characters
         * quoted (or escaped) as appropriate for the cell format type.  The format
         * type object is queried to see what is special.
         *
         * @param repl The original string.
         * @param type The format type representation object.
         *
         * @return A version of the string with any special characters Replaced.
         *
         * @see CellFormatType#isSpecial(char)
         */
        static String QuoteSpecial(String repl, CellFormatType type)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < repl.Length; i++)
            {
                char ch = repl[i];
                if (ch == '\'' && type.IsSpecial('\''))
                {
                    sb.Append('\u0000');
                    continue;
                }

                bool special = type.IsSpecial(ch);
                if (special)
                    sb.Append("'");
                sb.Append(ch);
                if (special)
                    sb.Append("'");
            }
            return sb.ToString();
        }

        /**
         * Apply this format part to the given value.  This returns a {@link
         * CellFormatResult} object with the results.
         *
         * @param value The value to apply this format part to.
         *
         * @return A {@link CellFormatResult} object Containing the results of
         *         Applying the format to the value.
         */
        public CellFormatResult Apply(Object value)
        {
            bool applies = Applies(value);
            String text;
            Color textColor;
            if (applies)
            {
                text = format.Format(value);
                textColor = color;
            }
            else
            {
                text = format.SimpleFormat(value);
                textColor = Color.Empty;
            }
            return new CellFormatResult(applies, text, textColor);
        }

        /**
         * Returns the CellFormatType object implied by the format specification for
         * the format part.
         *
         * @return The CellFormatType.
         */
        internal CellFormatType CellFormatType
        {
            get
            {
                return type;
            }
        }

        /**
         * Returns <tt>true</tt> if this format part has a condition.
         *
         * @return <tt>true</tt> if this format part has a condition.
         */
        internal bool HasCondition
        {
            get
            {
                return condition != null;
            }
        }
        public static StringBuilder ParseFormat(String fdesc, CellFormatType type,
                IPartHandler partHandler)
        {

            // Quoting is very awkward.  In the Java classes, quoting is done
            // between ' chars, with '' meaning a single ' char. The problem is that
            // in Excel, it is legal to have two adjacent escaped strings.  For
            // example, consider the Excel format "\a\b#".  The naive (and easy)
            // translation into Java DecimalFormat is "'a''b'#".  For the number 17,
            // in Excel you would Get "ab17", but in Java it would be "a'b17" -- the
            // '' is in the middle of the quoted string in Java.  So the trick we
            // use is this: When we encounter a ' char in the Excel format, we
            // output a \u0000 char into the string.  Now we know that any '' in the
            // output is the result of two adjacent escaped strings.  So After the
            // main loop, we have to do two passes: One to eliminate any ''
            // sequences, to make "'a''b'" become "'ab'", and another to replace any
            // \u0000 with '' to mean a quote char.  Oy.
            //
            // For formats that don't use "'" we don't do any of this

            MatchCollection mc = SPECIFICATION_PAT.Matches(fdesc);
            StringBuilder fmt = new StringBuilder();
            Match lastMatch = null;
            //while (m.Find())
            foreach(Match m in mc)
            {
                String part = Group(m, 0);
                if (part.Length > 0)
                {
                    String repl = partHandler.HandlePart(m, part, type, fmt);
                    if (repl == null)
                    {
                        switch (part[0])
                        {
                            case '\"':
                                repl = QuoteSpecial(part.Substring(1,
                                        part.Length - 2), type);
                                break;
                            case '\\':
                                repl = QuoteSpecial(part.Substring(1), type);
                                break;
                            case '_':
                                repl = " ";
                                break;
                            case '*': //!! We don't do this for real, we just Put in 3 of them
                                repl = ExpandChar(part);
                                break;
                            default:
                                repl = part;
                                break;
                        }
                    }
                    //m.AppendReplacement(fmt, Match.QuoteReplacement(repl));
                    fmt.Append(part.Replace(m.Captures[0].Value, repl));
                    if (m.NextMatch().Index - (m.Index + part.Length) > 0)
                    {
                        fmt.Append(fdesc.Substring(m.Index + part.Length, m.NextMatch().Index - (m.Index + part.Length)));
                    }
                    lastMatch = m;
                }
            }
            if (lastMatch != null)
            {
                fmt.Append(fdesc.Substring(lastMatch.Index + lastMatch.Groups[0].Value.Length));
            }
            //m.AppendTail(fmt);

            if (type.IsSpecial('\''))
            {
                // Now the next pass for quoted characters: Remove '' chars, making "'a''b'" into "'ab'"
                int pos = 0;
                while ((pos = fmt.ToString().IndexOf("''", pos)) >= 0)
                {
                    //fmt.Delete(pos, pos + 2);
                    fmt.Remove(pos, 2);
                }

                // Now the pass for quoted chars: Replace any \u0000 with ''
                pos = 0;
                while ((pos = fmt.ToString().IndexOf('\u0000', pos)) >= 0)
                {
                    //fmt.Replace(pos, pos + 1, "''");
                    fmt.Remove(pos, 1);
                    fmt.Insert(pos, "''");
                }
            }

            return fmt;
        }
        public static String QuoteReplacement(String s)
        {
            if ((s.IndexOf('\\') == -1) && (s.IndexOf('$') == -1))
                return s;
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[(i)];
                if (c == '\\' || c == '$')
                {
                    sb.Append('\\');
                }
                sb.Append(c);
            }
            return sb.ToString();
        }
        /**
         * Expands a character. This is only partly done, because we don't have the
         * correct info.  In Excel, this would be expanded to fill the rest of the
         * cell, but we don't know, in general, what the "rest of the cell" is1.
         *
         * @param part The character to be repeated is the second character in this
         *             string.
         *
         * @return The character repeated three times.
         */
        internal static String ExpandChar(String part)
        {
            String repl;
            char ch = part[1];
            repl = "" + ch + ch + ch;
            return repl;
        }

        /**
         * Returns the string from the group, or <tt>""</tt> if the group is
         * <tt>null</tt>.
         *
         * @param m The matcher.
         * @param g The group number.
         *
         * @return The group or <tt>""</tt>.
         */
        public static String Group(Match m, int g)
        {
            String str = m.Groups[(g)].Value;
            return (str == null ? "" : str);
        }
        public override string ToString()
        {
            return format.ToString();
        }
    }
}