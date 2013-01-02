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


namespace NPOI.HSSF.UserModel
{
    using System;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Aggregates;
    using NPOI.SS.UserModel;

    /// <summary>
    /// Class to Read and manipulate the footer.
    /// The footer works by having a left, center, and right side.  The total cannot
    /// be more that 255 bytes long.  One uses this class by Getting the HSSFFooter
    /// from HSSFSheet and then Getting or Setting the left, center, and right side.
    /// For special things (such as page numbers and date), one can use a the methods
    /// that return the Chars used to represent these.  One can also Change the
    /// fonts by using similar methods.
    /// @author Shawn Laubach (slaubach at apache dot org)
    /// </summary>
    public class HSSFFooter : HeaderFooter,IFooter
    {
        private PageSettingsBlock _psb;

        /// <summary>
        /// Initializes a new instance of the <see cref="HSSFFooter"/> class.
        /// </summary>
        /// <param name="psb">Footer record to create the footer with</param>
        public HSSFFooter(PageSettingsBlock psb)
        {
            _psb = psb;
        }


        /// <summary>
        /// Gets the raw footer.
        /// </summary>
        /// <value>The raw footer.</value>
        public override String RawText
        {
            get
            {
                FooterRecord hf = _psb.Footer;
                if (hf == null)
                {
                    return string.Empty;
                }
                return hf.Text;
            }
        }
        protected override void SetHeaderFooterText(string text)
        {
            FooterRecord hfr = _psb.Footer;
            if (hfr == null)
            {
                hfr = new FooterRecord(text);
                _psb.Footer=(hfr);
            }
            else
            {
                hfr.Text=(text);
            }
        }
    }
}