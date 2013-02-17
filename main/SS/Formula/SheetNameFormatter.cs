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

    /**
     * Formats sheet names for use in formula expressions.
     * 
     * @author Josh Micich
     */
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
        /**
         * Used to format sheet names as they would appear in cell formula expressions.
         * @return the sheet name UnChanged if there is no need for delimiting.  Otherwise the sheet
         * name is enclosed in single quotes (').  Any single quotes which were already present in the 
         * sheet name will be converted to double single quotes ('').  
         */
        public static String Format(String rawSheetName)
        {
            StringBuilder sb = new StringBuilder(rawSheetName.Length + 2);
            AppendFormat(sb, rawSheetName);
            return sb.ToString();
        }

        /**
         * Convenience method for when a StringBuilder is already available
         * 
         * @param out - sheet name will be Appended here possibly with delimiting quotes 
         */
        public static void AppendFormat(StringBuilder out1, String rawSheetName)
        {
            bool needsQuotes = NeedsDelimiting(rawSheetName);
            if (needsQuotes)
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
            if (needsQuotes)
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
            for (int i = 0; i < len; i++)
            {
                char ch = rawSheetName[i];
                if (ch == DELIMITER)
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
            if (len < 1)
            {
                throw new Exception("Zero Length string is an invalid sheet name");
            }
            if (Char.IsDigit(rawSheetName[0]))
            {
                // sheet name with digit in the first position always requires delimiting
                return true;
            }
            for (int i = 0; i < len; i++)
            {
                char ch = rawSheetName[i];
                if (IsSpecialChar(ch))
                {
                    return true;
                }
            }
            if (Char.IsLetter(rawSheetName[0])
                    && Char.IsDigit(rawSheetName[len - 1]))
            {
                // note - values like "A$1:$C$20" don't Get this far 
                if (NameLooksLikePlainCellReference(rawSheetName))
                {
                    return true;
                }
            }
            if (NameLooksLikeBooleanLiteral(rawSheetName))
            {
                return true;
            }
            return false;
        }
        private static bool NameLooksLikeBooleanLiteral(String rawSheetName)
        {
            switch (rawSheetName[0])
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
        /**
         * @return <c>true</c> if the presence of the specified Char in a sheet name would 
         * require the sheet name to be delimited in formulas.  This includes every non-alphanumeric 
         * Char besides Underscore '_'.
         */
        /* package */
        static bool IsSpecialChar(char ch)
        {
            // note - Char.IsJavaIdentifierPart() would allow dollars '$'
            if (Char.IsLetterOrDigit(ch))
            {
                return false;
            }
            switch (ch)
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


        /**
         * Used to decide whether sheet names like 'AB123' need delimiting due to the fact that they 
         * look like cell references.
         * <p/>
         * This code is currently being used for translating formulas represented with <code>Ptg</code>
         * tokens into human readable text form.  In formula expressions, a sheet name always has a 
         * trailing '!' so there is little chance for ambiguity.  It doesn't matter too much what this 
         * method returns but it is worth noting the likely consumers of these formula text strings:
         * <ol>
         * <li>POI's own formula parser</li>
         * <li>Visual reading by human</li>
         * <li>VBA automation entry into Excel cell contents e.g.  ActiveCell.Formula = "=c64!A1"</li>
         * <li>Manual entry into Excel cell contents</li>
         * <li>Some third party formula parser</li>
         * </ol>
         * 
         * At the time of writing, POI's formula parser tolerates cell-like sheet names in formulas
         * with or without delimiters.  The same goes for Excel(2007), both manual and automated entry.  
         * <p/>
         * For better or worse this implementation attempts to replicate Excel's formula renderer.
         * Excel uses range checking on the apparent 'row' and 'column' components.  Note however that
         * the maximum sheet size varies across versions.
         * @see org.apache.poi.hssf.util.CellReference
         */
        public static bool CellReferenceIsWithinRange(String lettersPrefix, String numbersSuffix)
        {
            return NPOI.SS.Util.CellReference.CellReferenceIsWithinRange(lettersPrefix, numbersSuffix, NPOI.SS.SpreadsheetVersion.EXCEL97);
        }

        /**
         * Note - this method assumes the specified rawSheetName has only letters and digits.  It 
         * cannot be used to match absolute or range references (using the dollar or colon char).
         * 
         * Some notable cases:
         *    <blockquote><table border="0" cellpAdding="1" cellspacing="0" 
         *                 summary="Notable cases.">
         *      <tr><th>Input </th><th>Result </th><th>Comments</th></tr>
         *      <tr><td>"A1" </td><td>true</td><td> </td></tr>
         *      <tr><td>"a111" </td><td>true</td><td> </td></tr>
         *      <tr><td>"AA" </td><td>false</td><td> </td></tr>
         *      <tr><td>"aa1" </td><td>true</td><td> </td></tr>
         *      <tr><td>"A1A" </td><td>false</td><td> </td></tr>
         *      <tr><td>"A1A1" </td><td>false</td><td> </td></tr>
         *      <tr><td>"A$1:$C$20" </td><td>false</td><td>Not a plain cell reference</td></tr>
         *      <tr><td>"SALES20080101" </td><td>true</td>
         *      		<td>Still needs delimiting even though well out of range</td></tr>
         *    </table></blockquote>
         *  
         * @return <c>true</c> if there is any possible ambiguity that the specified rawSheetName
         * could be interpreted as a valid cell name.
         */
        public static bool NameLooksLikePlainCellReference(String rawSheetName)
        {
            Regex matcher = new Regex(CELL_REF_PATTERN);
            if (!matcher.IsMatch(rawSheetName))
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