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
    using NPOI.SS.UserModel;
    using NPOI.HSSF.Record.Aggregates;

    /// <summary>
    /// Class to Read and manipulate the header.
    /// The header works by having a left, center, and right side.  The total cannot
    /// be more that 255 bytes long.  One uses this class by Getting the HSSFHeader
    /// from HSSFSheet and then Getting or Setting the left, center, and right side.
    /// For special things (such as page numbers and date), one can use a the methods
    /// that return the Chars used to represent these.  One can also Change the
    /// fonts by using similar methods.
    /// @author Shawn Laubach (slaubach at apache dot org)
    /// </summary>
    public class HSSFHeader : HeaderFooter,IHeader
    {
        private PageSettingsBlock _psb;

        /// <summary>
        /// Initializes a new instance of the <see cref="HSSFFooter"/> class.
        /// </summary>
        /// <param name="psb">Footer record to Create the footer with</param>
        public HSSFHeader(PageSettingsBlock psb)
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
                HeaderRecord hf = _psb.Header;
                if (hf == null)
                {
                    return string.Empty;
                }
                return hf.Text;
            }
        }
        protected override void SetHeaderFooterText(string text)
        {
            HeaderRecord hfr = _psb.Header;
            if (hfr == null)
            {
                hfr = new HeaderRecord(text);
                _psb.Header = (hfr);
            }
            else
            {
                hfr.Text=(text);
            }
        }
    }
}