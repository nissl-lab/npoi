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

    using NPOI.SS.UserModel;
    using System.Text.RegularExpressions;
    using System.Collections.Generic;
    using NPOI.Util;

    /**
     * Format a value according to the standard Excel behavior.  This "standard" is
     * not explicitly documented by Microsoft, so the behavior is determined by
     * experimentation; see the tests.
     * <p/>
     * An Excel format has up to four parts, separated by semicolons.  Each part
     * specifies what to do with particular kinds of values, depending on the number
     * of parts given: <dl> <dt>One part (example: <tt>[Green]#.##</tt>) <dd>If the
     * value is a number, display according to this one part (example: green text,
     * with up to two decimal points). If the value is text, display it as is.
     * <dt>Two parts (example: <tt>[Green]#.##;[Red]#.##</tt>) <dd>If the value is a
     * positive number or zero, display according to the first part (example: green
     * text, with up to two decimal points); if it is a negative number, display
     * according to the second part (example: red text, with up to two decimal
     * points). If the value is text, display it as is. <dt>Three parts (example:
     * <tt>[Green]#.##;[Black]#.##;[Red]#.##</tt>) <dd>If the value is a positive
     * number, display according to the first part (example: green text, with up to
     * two decimal points); if it is zero, display according to the second part
     * (example: black text, with up to two decimal points); if it is a negative
     * number, display according to the third part (example: red text, with up to
     * two decimal points). If the value is text, display it as is. <dt>Four parts
     * (example: <tt>[Green]#.##;[Black]#.##;[Red]#.##;[@]</tt>) <dd>If the value is
     * a positive number, display according to the first part (example: green text,
     * with up to two decimal points); if it is zero, display according to the
     * second part (example: black text, with up to two decimal points); if it is a
     * negative number, display according to the third part (example: red text, with
     * up to two decimal points). If the value is text, display according to the
     * fourth part (example: text in the cell's usual color, with the text value
     * surround by brackets). </dd></dt></dd></dt></dd></dt></dd></dt></dl>
     * <p/>
     * A given format part may specify a given Locale, by including something
     *  like <tt>[$$-409]</tt> or <tt>[$&pound;-809]</tt> or <tt>[$-40C]</tt>. These
     *  are (currently) largely ignored. You can use {@link DateFormatConverter}
     *  to look these up into Java Locales if desired.
     * <p/>
     * In addition to these, there is a general format that is used when no format
     * is specified.  This formatting is presented by the {@link #GENERAL_FORMAT}
     * object.
     * 
     * TODO Merge this with {@link DataFormatter} so we only have one set of
     *  code for formatting numbers.
     * TODO Re-use parts of this logic with {@link ConditionalFormatting} /
     *  {@link ConditionalFormattingRule} for reporting stylings which do/don't apply
     * TODO Support the full set of modifiers, including alternate calendars and
     *  native character numbers, as documented at https://help.libreoffice.org/Common/Number_Format_Codes
     */
    public class CellFormat
    {
        private String format;
        private CellFormatPart posNumFmt;
        private CellFormatPart zeroNumFmt;
        private CellFormatPart negNumFmt;
        private CellFormatPart textFmt;
        private int formatPartCount;

        private static readonly Regex ONE_PART = new Regex(CellFormatPart.FORMAT_PAT.ToString() + "(;|$)", RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace | RegexOptions.Compiled);

        private static readonly CellFormatPart DEFAULT_TEXT_FORMAT =
                new CellFormatPart("@");

        /*
         * Cells that cannot be formatted, e.g. cells that have a date or time
         * format and have an invalid date or time value, are displayed as 255
         * pound signs ("#").
         */
        private const string INVALID_VALUE_FOR_FORMAT =
                "###################################################" +
                "###################################################" +
                "###################################################" +
                "###################################################" +
                "###################################################";

        private const string QUOTE = "\"";

        private static readonly CellFormat GENERAL_FORMAT = new GeneralCellFormat();
        /**
         * Format a value as it would be were no format specified.  This is also
         * used when the format specified is <tt>General</tt>.
         */
        public class GeneralCellFormat : CellFormat
        {
            public GeneralCellFormat()
                : base("General")
            {
            }
            public override CellFormatResult Apply(Object value)
            {
                String text = (new CellGeneralFormatter()).Format(value);
                return new CellFormatResult(true, text, POIUtils.Color_Empty);
            }
        }

        /** Maps a format string to its Parsed version for efficiencies sake. */
        private static Dictionary<String, CellFormat> formatCache =
                new Dictionary<String, CellFormat>();

        /**
         * Returns a {@link CellFormat} that applies the given format.  Two calls
         * with the same format may or may not return the same object.
         *
         * @param format The format.
         *
         * @return A {@link CellFormat} that applies the given format.
         */
        public static CellFormat GetInstance(String format)
        {
            CellFormat fmt = null;
            if (formatCache.ContainsKey(format))
                fmt = formatCache[format];
            if (fmt == null)
            {
                if (format.Equals("General") || format.Equals("@"))
                    fmt = GENERAL_FORMAT;
                else
                    fmt = new CellFormat(format);
                formatCache.Add(format, fmt);
            }
            return fmt;
        }

        /**
         * Creates a new object.
         *
         * @param format The format.
         */
        private CellFormat(String format)
        {
            this.format = format;
            MatchCollection mc = ONE_PART.Matches(format);
            List<CellFormatPart> parts = new List<CellFormatPart>();

            //while (m.Success)
            foreach (Match m in mc)
            {
                try
                {
                    String valueDesc = m.Groups[0].Value;

                    // Strip out the semicolon if it's there
                    if (valueDesc.EndsWith(";"))
                        valueDesc = valueDesc.Substring(0, valueDesc.Length - 1);

                    parts.Add(new CellFormatPart(valueDesc));
                }
                catch (Exception)
                {
                    //CellFormatter.logger.Log(Level.WARNING,
                    //        "Invalid format: " + CellFormatter.Quote(m.Group()), e);
                    parts.Add(null);
                }
            }
            formatPartCount = parts.Count;
            switch (formatPartCount)
            {
                case 1:
                    posNumFmt = parts[(0)];
                    negNumFmt = null;
                    zeroNumFmt = null;
                    textFmt = DEFAULT_TEXT_FORMAT;
                    break;
                case 2:
                    posNumFmt = parts[0];
                    negNumFmt = parts[1];
                    zeroNumFmt = null;
                    textFmt = DEFAULT_TEXT_FORMAT;
                    break;
                case 3:
                    posNumFmt = parts[0];
                    negNumFmt = parts[1];
                    zeroNumFmt = parts[2];
                    textFmt = DEFAULT_TEXT_FORMAT;
                    break;
                case 4:
                default:
                    posNumFmt = parts[0];
                    negNumFmt = parts[1];
                    zeroNumFmt = parts[2];
                    textFmt = parts[3];
                    break;
            }
        }

        /**
         * Returns the result of Applying the format to the given value.  If the
         * value is a number (a type of {@link Number} object), the correct number
         * format type is chosen; otherwise it is considered a text object.
         *
         * @param value The value
         *
         * @return The result, in a {@link CellFormatResult}.
         */
        public virtual CellFormatResult Apply(Object value)
        {
            //if (value is Number) {
            if (NPOI.Util.Number.IsNumber(value))
            {
                double val ;
                double.TryParse(value.ToString(), out val);
                if (val < 0 &&
                    ((formatPartCount == 2
                            && !posNumFmt.HasCondition && !negNumFmt.HasCondition)
                    || (formatPartCount == 3 && !negNumFmt.HasCondition)
                    || (formatPartCount == 4 && !negNumFmt.HasCondition)))
                {
                    // The negative number format has the negative formatting required,
                    // e.g. minus sign or brackets, so pass a positive value so that
                    // the default leading minus sign is not also output
                    return negNumFmt.Apply(-val);
                }
                else
                {
                    return GetApplicableFormatPart(val).Apply(val);
                }
            }
            else if (value is DateTime)
            {
                // Don't know (and can't get) the workbook date windowing (1900 or 1904)
                // so assume 1900 date windowing
                Double numericValue = DateUtil.GetExcelDate((DateTime)value);
                if (DateUtil.IsValidExcelDate(numericValue))
                {
                    return GetApplicableFormatPart(numericValue).Apply(value);
                }
                else
                {
                    throw new ArgumentException("value " + numericValue + " of date " + value + " is not a valid Excel date");
                }
            }
            else
            {
                return textFmt.Apply(value);
            }
        }
        /**
         * Returns the result of applying the format to the given date.
         *
         * @param date         The date.
         * @param numericValue The numeric value for the date.
         *
         * @return The result, in a {@link CellFormatResult}.
         */
        private CellFormatResult Apply(DateTime date, double numericValue)
        {
            return GetApplicableFormatPart(numericValue).Apply(date);
        }
        /**
         * Fetches the appropriate value from the cell, and returns the result of
         * Applying it to the appropriate format.  For formula cells, the computed
         * value is what is used.
         *
         * @param c The cell.
         *
         * @return The result, in a {@link CellFormatResult}.
         */
        public CellFormatResult Apply(ICell c)
        {
            switch (UltimateType(c))
            {
                case CellType.Blank:
                    return Apply("");
                case CellType.Boolean:
                    return Apply(c.BooleanCellValue);
                case CellType.Numeric:
                    Double value = c.NumericCellValue;
                    if (GetApplicableFormatPart(value).CellFormatType == CellFormatType.DATE)
                    {
                        if (DateUtil.IsValidExcelDate(value))
                        {
                            return Apply(c.DateCellValue, value);
                        }
                        else
                        {
                            return Apply(INVALID_VALUE_FOR_FORMAT);
                        }
                    }
                    else
                    {
                        return Apply(value);
                    }
                case CellType.String:
                    return Apply(c.StringCellValue);
                default:
                    return Apply("?");
            }
        }

        /**
         * Returns the {@link CellFormatPart} that applies to the value.  Result
         * depends on how many parts the cell format has, the cell value and any
         * conditions.  The value must be a {@link Number}.
         * 
         * @param value The value.
         * @return The {@link CellFormatPart} that applies to the value.
         */
        private CellFormatPart GetApplicableFormatPart(Object value)
        {
            //if (value is Number) {
            if (NPOI.Util.Number.IsNumber(value))
            {
                double val;
                double.TryParse(value.ToString(), out val);

                if (formatPartCount == 1)
                {
                    if (!posNumFmt.HasCondition
                            || (posNumFmt.HasCondition && posNumFmt.Applies(val)))
                    {
                        return posNumFmt;
                    }
                    else
                    {
                        return new CellFormatPart("General");
                    }
                }
                else if (formatPartCount == 2)
                {
                    if ((!posNumFmt.HasCondition && val >= 0)
                            || (posNumFmt.HasCondition && posNumFmt.Applies(val)))
                    {
                        return posNumFmt;
                    }
                    else if (!negNumFmt.HasCondition
                            || (negNumFmt.HasCondition && negNumFmt.Applies(val)))
                    {
                        return negNumFmt;
                    }
                    else
                    {
                        // Return ###...### (255 #s) to match Excel 2007 behaviour
                        return new CellFormatPart(QUOTE + INVALID_VALUE_FOR_FORMAT + QUOTE);
                    }
                }
                else
                {
                    if ((!posNumFmt.HasCondition && val > 0)
                            || (posNumFmt.HasCondition && posNumFmt.Applies(val)))
                    {
                        return posNumFmt;
                    }
                    else if ((!negNumFmt.HasCondition && val < 0)
                          || (negNumFmt.HasCondition && negNumFmt.Applies(val)))
                    {
                        return negNumFmt;
                        // Only the first two format parts can have conditions
                    }
                    else
                    {
                        return zeroNumFmt;
                    }
                }
            }
            else
            {
                throw new ArgumentException("value must be a Number");
            }

        }
        /**
         * Returns the ultimate cell type, following the results of formulas.  If
         * the cell is a {@link Cell#CELL_TYPE_FORMULA}, this returns the result of
         * {@link Cell#getCachedFormulaResultType()}.  Otherwise this returns the
         * result of {@link Cell#getCellType()}.
         *
         * @param cell The cell.
         *
         * @return The ultimate type of this cell.
         */
        public static CellType UltimateType(ICell cell)
        {
            CellType type = cell.CellType;
            if (type == CellType.Formula)
                return cell.CachedFormulaResultType;
            else
                return type;
        }

        /**
         * Returns <tt>true</tt> if the other object is a {@link CellFormat} object
         * with the same format.
         *
         * @param obj The other object.
         *
         * @return <tt>true</tt> if the two objects are Equal.
         */

        public override bool Equals(Object obj)
        {
            if (this == obj)
                return true;
            if (obj is CellFormat)
            {
                CellFormat that = (CellFormat)obj;
                return format.Equals(that.format);
            }
            return false;
        }

        /**
         * Returns a hash code for the format.
         *
         * @return A hash code for the format.
         */

        public override int GetHashCode()
        {
            return format.GetHashCode();
        }

        public override string ToString()
        {
            return format;
        }
    }
}