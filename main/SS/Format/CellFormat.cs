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
    using System.Windows.Forms;
    using System.Drawing;








    /**
     * Format a value according to the standard Excel behavior.  This "standard" is
     * not explicitly documented by Microsoft, so the behavior is determined by
     * experimentation; see the tests.
     * 
     * An Excel format has up to four parts, Separated by semicolons.  Each part
     * specifies what to do with particular kinds of values, depending on the number
     * of parts given: 
     * 
     * - One part (example: <c>[Green]#.##</c>) 
     * If the value is a number, display according to this one part (example: green text,
     * with up to two decimal points). If the value is text, display it as is.
     * 
     * - Two parts (example: <c>[Green]#.##;[Red]#.##</c>) 
     * If the value is a positive number or zero, display according to the first part (example: green
     * text, with up to two decimal points); if it is a negative number, display
     * according to the second part (example: red text, with up to two decimal
     * points). If the value is text, display it as is. 
     * 
     * - Three parts (example: <c>[Green]#.##;[Black]#.##;[Red]#.##</c>) 
     * If the value is a positive number, display according to the first part (example: green text, with up to
     * two decimal points); if it is zero, display according to the second part
     * (example: black text, with up to two decimal points); if it is a negative
     * number, display according to the third part (example: red text, with up to
     * two decimal points). If the value is text, display it as is.
     * 
     * - Four parts (example: <c>[Green]#.##;[Black]#.##;[Red]#.##;[@]</c>)
     * If the value is a positive number, display according to the first part (example: green text,
     * with up to two decimal points); if it is zero, display according to the
     * second part (example: black text, with up to two decimal points); if it is a
     * negative number, display according to the third part (example: red text, with
     * up to two decimal points). If the value is text, display according to the
     * fourth part (example: text in the cell's usual color, with the text value
     * surround by brackets).
     * 
     * In Addition to these, there is a general format that is used when no format
     * is specified.  This formatting is presented by the {@link #GENERAL_FORMAT}
     * object.
     *
     * @author Ken Arnold, Industrious Media LLC
     */
    public class CellFormat
    {
        private String format;
        private CellFormatPart posNumFmt;
        private CellFormatPart zeroNumFmt;
        private CellFormatPart negNumFmt;
        private CellFormatPart textFmt;

        private static Regex ONE_PART = new Regex(CellFormatPart.FORMAT_PAT.ToString() + "(;|$)", RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace);

        private static CellFormatPart DEFAULT_TEXT_FORMAT =
                new CellFormatPart("@");

        private static CellFormat GENERAL_FORMAT = new GeneralCellFormat();
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
                String text;
                if (value == null)
                {
                    text = "";
                }
                else if (value.GetType().IsPrimitive/* is Number*/)
                {
                    throw new NotImplementedException();
                    //text = CellNumberFormatter.SIMPLE_NUMBER.Format(value);
                }
                else
                {
                    text = value.ToString();
                }
                return new CellFormatResult(true, text, Color.Empty);
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
                if (format.Equals("General"))
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
            foreach(Match m in mc)
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

            switch (parts.Count)
            {
                case 1:
                    posNumFmt = zeroNumFmt = negNumFmt = parts[(0)];
                    textFmt = DEFAULT_TEXT_FORMAT;
                    break;
                case 2:
                    posNumFmt = zeroNumFmt = parts[0];
                    negNumFmt = parts[1];
                    textFmt = DEFAULT_TEXT_FORMAT;
                    break;
                case 3:
                    posNumFmt = parts[0];
                    zeroNumFmt = parts[1];
                    negNumFmt = parts[2];
                    textFmt = DEFAULT_TEXT_FORMAT;
                    break;
                case 4:
                default:
                    posNumFmt = parts[0];
                    zeroNumFmt = parts[1];
                    negNumFmt = parts[2];
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
            if (value.GetType().IsPrimitive/* is Number*/)
            {
                double num = (double)value;
                double val = num;
                if (val > 0)
                    return posNumFmt.Apply(value);
                else if (val < 0)
                    return negNumFmt.Apply(-val);
                else
                    return zeroNumFmt.Apply(value);
            }
            else
            {
                return textFmt.Apply(value);
            }
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
                case CellType.BLANK:
                    return Apply("");
                case CellType.BOOLEAN:
                    return Apply(c.BooleanCellValue.ToString());
                case CellType.NUMERIC:
                    return Apply(c.NumericCellValue);
                case CellType.STRING:
                    return Apply(c.StringCellValue);
                default:
                    return Apply("?");
            }
        }

        /**
         * Uses the result of Applying this format to the value, Setting the text
         * and color of a label before returning the result.
         *
         * @param label The label to apply to.
         * @param value The value to Process.
         *
         * @return The result, in a {@link CellFormatResult}.
         */
        public CellFormatResult Apply(Label label, Object value)
        {
            CellFormatResult result = Apply(value);
            label.Text = (/*setter*/result.Text);
            if (result.TextColor != Color.Empty)
            {
                label.ForeColor = (/*setter*/result.TextColor);
            }
            return result;
        }

        /**
         * Fetches the appropriate value from the cell, and uses the result, Setting
         * the text and color of a label before returning the result.
         *
         * @param label The label to apply to.
         * @param c     The cell.
         *
         * @return The result, in a {@link CellFormatResult}.
         */
        public CellFormatResult Apply(Label label, ICell c)
        {
            switch (UltimateType(c))
            {
                case CellType.BLANK:
                    return Apply(label, "");
                case CellType.BOOLEAN:
                    return Apply(c.BooleanCellValue.ToString());
                case CellType.NUMERIC:
                    return Apply(label, c.NumericCellValue);
                case CellType.STRING:
                    return Apply(label, c.StringCellValue);
                default:
                    return Apply(label, "?");
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
            if (type == CellType.FORMULA)
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