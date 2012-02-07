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

using NPOI.XSSF.UserModel.Helpers;
using System;
using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
namespace NPOI.XSSF.UserModel.Extensions
{

    /**
     * Parent class of all XSSF headers and footers.
     *
     * For a list of all the different fields that can be placed into a header or
     * footer, such as page number, bold, underline etc, see the follow formatting
     * syntax
     *
     *<b> Header/Footer Formatting Syntax</b>
     *<p>
     * There are a number of formatting codes that can be written inline with the
     * actual header / footer text, which affect the formatting in the header or
     * footer.
     *</p>
     *
     * This example Shows the text "Center Bold Header" on the first line (center
     * section), and the date on the second line (center section). &CCenter
     * &"-,Bold"Bold &"-,Regular"Header_x000A_&D
     *
     * <b>General Rules:</b> There is no required order in which these codes must
     * appear. The first occurrence of the following codes turns the formatting ON,
     * the second occurrence turns it OFF again:
     *
     * <dl>
     * <dt>&L</dt>
     * <dd>code for "left section" (there are three header / footer locations,
     * "left", "center", and "right"). When two or more occurrences of this section
     * marker exist, the contents from all markers are concatenated, in the order of
     * appearance, and placed into the left section.</dd>
     * <dt>&P</dt>
     * <dd>code for "current page #"</dd>
     * <dt>&N</dt>
     * <dd>code for "total pages"</dd>
     * <dt>&font size</dt>
     * <dd>code for "text font size", where font size is a font size in points.</dd>
     * <dt>&K</dt>
     * <dd>code for "text font color" RGB Color is specified as RRGGBB Theme Color
     * is specifed as TTSNN where TT is the theme color Id, S is either "+" or "-"
     * of the tint/shade value, NN is the tint/shade value.</dd>
     * <dt>&S</dt>
     * <dd>code for "text strikethrough" on / off</dd>
     * <dt>&X</dt>
     * <dd>code for "text super script" on / off</dd>
     * <dt>&Y</dt>
     * <dd>code for "text subscript" on / off</dd>
     * <dt>&C</dt>
     * <dd>code for "center section". When two or more occurrences of this section
     * marker exist, the contents from all markers are concatenated, in the order of
     * appearance, and placed into the center section. SpreadsheetML Reference
     * Material - Worksheets 1966</dd>
     * <dt>&D</dt>
     * <dd>code for "date"</dd>
     * <dt>&T</dt>
     * <dd>code for "time"</dd>
     * <dt>&G</dt>
     * <dd>code for "picture as background"</dd>
     * <dt>&U</dt>
     * <dd>code for "text single underline"</dd>
     * <dt>&E</dt>
     * <dd>code for "double underline"</dd>
     * <dt>&R</dt>
     * <dd>code for "right section". When two or more occurrences of this section
     * marker exist, the contents from all markers are concatenated, in the order of
     * appearance, and placed into the right section.</dd>
     * <dt>&Z</dt>
     * <dd>code for "this workbook's file path"</dd>
     * <dt>&F</dt>
     * <dd>code for "this workbook's file name"</dd>
     * <dt>&A</dt>
     * <dd>code for "sheet tab name"</dd>
     * <dt>&+</dt>
     * <dd>code for add to page #.</dd>
     * <dt>&-</dt>
     * <dd>code for subtract from page #.</dd>
     * <dt>&"font name,font type" - code for "text font name" and "text font type",
     * where font name and font type are strings specifying the name and type of the
     * font, Separated by a comma. When a hyphen appears in font name, it means
     * "none specified". Both of font name and font type can be localized
     * values.</dd>
     * <dt>&"-,Bold"</dt>
     * <dd>code for "bold font style"</dd>
     * <dt>&B</dt>
     * <dd>also means "bold font style"</dd>
     * <dt>&"-,Regular"</dt>
     * <dd>code for "regular font style"</dd>
     * <dt>&"-,Italic"</dt>
     * <dd>code for "italic font style"</dd>
     * <dt>&I</dt>
     * <dd>also means "italic font style"</dd>
     * <dt>&"-,Bold Italic"</dt>
     * <dd>code for "bold italic font style"</dd>
     * <dt>&O</dt>
     * <dd>code for "outline style"</dd>
     * <dt>&H</dt>
     * <dd>code for "shadow style"</dd>
     * </dl>
     *
     *
     */
    public abstract class XSSFHeaderFooter : IHeaderFooter
    {
        private HeaderFooterHelper helper;
        private CT_HeaderFooter headerFooter;

        private bool stripFields = false;

        /**
         * Create an instance of XSSFHeaderFooter from the supplied XML bean
         *
         * @param headerFooter
         */
        public XSSFHeaderFooter(CT_HeaderFooter headerFooter)
        {
            this.headerFooter = headerFooter;
            this.helper = new HeaderFooterHelper();
        }

        /**
         * Returns the underlying CTHeaderFooter xml bean
         *
         * @return the underlying CTHeaderFooter xml bean
         */

        public CT_HeaderFooter GetHeaderFooter()
        {
            return this.headerFooter;
        }

        public String GetValue()
        {
            String value = Text;
            if (value == null)
                return "";
            return value;
        }

        /**
         * Are fields currently being stripped from the text that this
         * {@link XSSFHeaderFooter} returns? Default is false, but can be Changed
         */
        public bool AreFieldsStripped()
        {
            return stripFields;
        }

        /**
         * Should fields (eg macros) be stripped from the text that this class
         * returns? Default is not to strip.
         *
         * @param StripFields
         */
        public void SetAreFieldsStripped(bool stripFields)
        {
            this.stripFields = stripFields;
        }

        public abstract String Text
        {
            get;
            set;
        }
        /**
 * Removes any fields (eg macros, page markers etc) from the string.
 * Normally used to make some text suitable for showing to humans, and the
 * resultant text should not normally be saved back into the document!
 */
        public static String StripFields(String text)
        {
            return HeaderFooter.StripFields(text);
        }
        /**
         * get the text representing the center part of this element
         */
        public String Center
        {
            get
            {
                String text = helper.GetCenterSection(Text);
                if (stripFields)
                    return StripFields(text);
                return text;
            }
            set 
            {
                this.Text = (helper.SetCenterSection(Text, value));
            }
        }

        /**
         * get the text representing the left part of this element
         */
        public String Left
        {
            get
            {
                String text = helper.GetLeftSection(Text);
                if (stripFields)
                    return StripFields(text);
                return text;
            }
            set 
            {
                this.Text = helper.SetLeftSection(Text, value);
            }
        }

        /**
         * get the text representing the right part of this element
         */
        public String Right
        {
            get
            {
                String text = helper.GetRightSection(Text);
                if (stripFields)
                    return StripFields(text);
                return text;
            }
            set 
            {
                this.Text = (helper.SetRightSection(Text, value));
            }
        }
    }
}


