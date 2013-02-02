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

namespace NPOI.HSSF.UserModel
{
    using NPOI.SS.UserModel;
    using System.Globalization;

    /**
     * HSSFDataFormatter contains methods for formatting the value stored in an
     * HSSFCell. This can be useful for reports and GUI presentations when you
     * need to display data exactly as it appears in Excel. Supported formats
     * include currency, SSN, percentages, decimals, dates, phone numbers, zip
     * codes, etc.
     * 
     * Internally, formats will be implemented using subclasses of <see cref="NPOI.SS.Util.FormatBase"/>
     * such as <see cref="NPOI.SS.Util.DecimalFormat"/> and <see cref="NPOI.SS.Util.SimpleDateFormat"/>. Therefore the
     * formats used by this class must obey the same pattern rules as these Format
     * subclasses. This means that only legal number pattern characters ("0", "#",
     * ".", "," etc.) may appear in number formats. Other characters can be
     * inserted <em>before</em> or <em>after</em> the number pattern to form a
     * prefix or suffix.
     * 
     * For example the Excel pattern <c>"$#,##0.00 "USD"_);($#,##0.00 "USD")"
     * </c> will be correctly formatted as "$1,000.00 USD" or "($1,000.00 USD)".
     * However the pattern <c>"00-00-00"</c> is incorrectly formatted by
     * DecimalFormat as "000000--". For Excel formats that are not compatible with
     * DecimalFormat, you can provide your own custom {@link Format} implementation
     * via <c>HSSFDataFormatter.AddFormat(String,Format)</c>. The following
     * custom formats are already provided by this class:
     * 
     * <pre>
     * <ul><li>SSN "000-00-0000"</li>
     *     <li>Phone Number "(###) ###-####"</li>
     *     <li>Zip plus 4 "00000-0000"</li>
     * </ul>
     * </pre>
     * 
     * If the Excel format pattern cannot be parsed successfully, then a default
     * format will be used. The default number format will mimic the Excel General
     * format: "#" for whole numbers and "#.##########" for decimal numbers. You
     * can override the default format pattern with <c>
     * HSSFDataFormatter.DefaultNumberFormat=(Format)</c>. <b>Note:</b> the
     * default format will only be used when a Format cannot be created from the
     * cell's data format string.
     *
     * @author James May (james dot may at fmr dot com)
     */
    public class HSSFDataFormatter : DataFormatter
    {

        /**
         * Creates a formatter using the given locale.
         */
        public HSSFDataFormatter(CultureInfo locale)
            : base(locale)
        {

        }

        /**
         * Creates a formatter using the {@link Locale#getDefault() default locale}.
         */
        public HSSFDataFormatter()
            : this(System.Globalization.CultureInfo.CurrentCulture)
        {
        }

    }

}