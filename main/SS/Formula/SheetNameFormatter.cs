/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */


namespace NPOI.SS.Formula
{
    using System;
    using System.Text;
    using System.Text.RegularExpressions;
    using NPOI.Util;
    using System.Globalization;

    /// <summary>
    /// Formats sheet names for use in formula expressions.
    /// </summary>
    /// @author Josh Micich
    public class SheetNameFormatter
    {

        private const string BIFF8_LAST_COLUMN = "IV";
        private const int BIFF8_LAST_COLUMN_TEXT_LEN = 2;
        private static readonly string BIFF8_LAST_ROW = (0x10000).ToString(CultureInfo.InvariantCulture);
        private static readonly int BIFF8_LAST_ROW_TEXT_LEN = BIFF8_LAST_ROW.Length;

        private const char DELIMITER = '\'';

        private const string CELL_REF_PATTERN = "^([A-Za-z]+)([0-9]+)$";

        private SheetNameFormatter()
        {
            // no instances of this class
        }
        /// <summary>
        /// Used to format sheet names as they would appear in cell formula expressions.
        /// </summary>
        /// <returns>the sheet name UnChanged if there is no need for delimiting.  Otherwise the sheet
        /// name is enclosed in single quotes (').  Any single quotes which were already present in the
        /// sheet name will be converted to double single quotes ('').
        /// </returns>
        public static String Format(String rawSheetName)
        {
            StringBuilder sb = new StringBuilder(rawSheetName.Length + 2);
            AppendFormat(sb, rawSheetName);
            return sb.ToString();
        }

        /// <summary>
        /// Convenience method for when a StringBuilder is already available
        /// </summary>
        /// <param name="out1">- sheet name will be Appended here possibly with delimiting quotes</param>
        public static void AppendFormat(StringBuilder out1, String rawSheetName)
        {
            bool needsQuotes = NeedsDelimiting(rawSheetName);
            if(needsQuotes)
            {
                out1.Append(DELIMITER);
                AppendAndEscape(out1, rawSheetName);
                out1.Append(DELIMITER);
            }
            else
            {
                out1.Append(rawSheetName);
            }
        }

        public static void AppendFormat(StringBuilder out1, String workbookName, String rawSheetName)
        {
            bool needsQuotes = NeedsDelimiting(workbookName) || NeedsDelimiting(rawSheetName);
            if(needsQuotes)
            {
                out1.Append(DELIMITER);
                out1.Append('[');
                AppendAndEscape(out1, workbookName.Replace('[', '(').Replace(']', ')'));
                out1.Append(']');
                AppendAndEscape(out1, rawSheetName);
                out1.Append(DELIMITER);
            }
            else
            {
                out1.Append('[');
                out1.Append(workbookName);
                out1.Append(']');
                out1.Append(rawSheetName);
            }
        }

        private static void AppendAndEscape(StringBuilder sb, String rawSheetName)
        {
            int len = rawSheetName.Length;
            for(int i = 0; i < len; i++)
            {
                char ch = rawSheetName[i];
                if(ch == DELIMITER)
                {
                    // single quotes (') are encoded as ('')
                    sb.Append(DELIMITER);
                }
                sb.Append(ch);
            }
        }

        private static bool NeedsDelimiting(String rawSheetName)
        {
            int len = rawSheetName.Length;
            if(len < 1)
            {
                throw new Exception("Zero Length string is an invalid sheet name");
            }
            if(Char.IsDigit(rawSheetName[0]))
            {
                // sheet name with digit in the first position always requires delimiting
                return true;
            }
            for(int i = 0; i < len; i++)
            {
                char ch = rawSheetName[i];
                if(IsSpecialChar(ch))
                {
                    return true;
                }
            }
            if(Char.IsLetter(rawSheetName[0])
                    && Char.IsDigit(rawSheetName[len - 1]))
            {
                // note - values like "A$1:$C$20" don't Get this far 
                if(NameLooksLikePlainCellReference(rawSheetName))
                {
                    return true;
                }
            }
            if(NameLooksLikeBooleanLiteral(rawSheetName))
            {
                return true;
            }
            return false;
        }
        private static bool NameLooksLikeBooleanLiteral(String rawSheetName)
        {
            switch(rawSheetName[0])
            {
                case 'T':
                case 't':
                    return "TRUE".Equals(rawSheetName, StringComparison.OrdinalIgnoreCase);
                case 'F':
                case 'f':
                    return "FALSE".Equals(rawSheetName, StringComparison.OrdinalIgnoreCase);
            }
            return false;
        }
        /// <summary>
        /// </summary>
        /// <returns><c>true</c> if the presence of the specified Char in a sheet name would
        /// require the sheet name to be delimited in formulas.  This includes every non-alphanumeric
        /// Char besides Underscore '_'.
        /// </returns>
        /* package */
        static bool IsSpecialChar(char ch)
        {
            // note - Char.IsJavaIdentifierPart() would allow dollars '$'
            if(Char.IsLetterOrDigit(ch))
            {
                return false;
            }
            switch(ch)
            {
                case '.': // dot is OK
                case '_': // Underscore is ok
                    return false;
                case '\n':
                case '\r':
                case '\t':
                    throw new Exception("Illegal Char (0x"
                            + StringUtil.ToHexString(ch) + ") found in sheet name");
            }
            return true;
        }


        /// <summary>
        /// <para>
        /// Used to decide whether sheet names like 'AB123' need delimiting due to the fact that they
        /// look like cell references.
        /// </para>
        /// <para>
        /// This code is currently being used for translating formulas represented with <c>Ptg</c>
        /// tokens into human readable text form.  In formula expressions, a sheet name always has a
        /// trailing '!' so there is little chance for ambiguity.  It doesn't matter too much what this
        /// method returns but it is worth noting the likely consumers of these formula text strings:
        /// <list type="number">
        /// <item><description>POI's own formula parser</description></item>
        /// <item><description>Visual reading by human</description></item>
        /// <item><description>VBA automation entry into Excel cell contents e.g.  ActiveCell.Formula = "=c64!A1"</description></item>
        /// <item><description>Manual entry into Excel cell contents</description></item>
        /// <item><description>Some third party formula parser</description></item>
        /// </list>
        /// </para>
        /// <para>
        /// At the time of writing, POI's formula parser tolerates cell-like sheet names in formulas
        /// with or without delimiters.  The same goes for Excel(2007), both manual and automated entry.
        /// </para>
        /// <para>
        /// For better or worse this implementation attempts to replicate Excel's formula renderer.
        /// Excel uses range checking on the apparent 'row' and 'column' components.  Note however that
        /// the maximum sheet size varies across versions.
        /// </para>
        /// </summary>
        /// <see cref="CellReference" />
        public static bool CellReferenceIsWithinRange(String lettersPrefix, String numbersSuffix)
        {
            return NPOI.SS.Util.CellReference.CellReferenceIsWithinRange(lettersPrefix, numbersSuffix, NPOI.SS.SpreadsheetVersion.EXCEL97);
        }

        /// <summary>
        /// <para>
        /// Note - this method assumes the specified rawSheetName has only letters and digits.  It
        /// cannot be used to match absolute or range references (using the dollar or colon char).
        /// </para>
        /// <para>
        /// Some notable cases:
        /// <list type="table">
        /// <listheader><term>Input </term><term>Result </term><description>Comments</description></listheader>
        /// <item><term>"A1" true </term><term>"a111" true </term>
        /// <term>"AA" false </term><term>"aa1" true </term><term>"A1A" false </term><term>"A1A1" false </term><term>"A$1:$C$20" falseNot a plain cell reference</term><description>"SALES20080101" true
        ///      		Still needs delimiting even though well out of range</description></item>
        /// </list>
        /// </para>
        /// </summary>
        /// <returns><c>true</c> if there is any possible ambiguity that the specified rawSheetName
        /// could be interpreted as a valid cell name.
        /// </returns>
        public static bool NameLooksLikePlainCellReference(String rawSheetName)
        {
            Regex matcher = new Regex(CELL_REF_PATTERN);
            if(!matcher.IsMatch(rawSheetName))
            {
                return false;
            }

            Match match = matcher.Matches(rawSheetName)[0];
            // rawSheetName == "Sheet1" Gets this far.
            String lettersPrefix = match.Groups[1].Value;
            String numbersSuffix = match.Groups[2].Value;
            return CellReferenceIsWithinRange(lettersPrefix, numbersSuffix);
        }

    }
}