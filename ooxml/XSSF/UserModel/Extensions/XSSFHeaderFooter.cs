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
    /// <summary>
    /// Parent class of all XSSF headers and footers.
    /// </summary>
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


