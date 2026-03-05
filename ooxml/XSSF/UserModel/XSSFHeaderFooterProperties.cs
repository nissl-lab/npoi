/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.XSSF.UserModel
{
    using NPOI.Util;

    using NPOI.OpenXmlFormats.Spreadsheet;


    /// <summary>
    /// All Header/Footer properties for a sheet are scoped to the sheet. This includes Different First Page,
    /// and Different Even/Odd. These properties can be Set or unset explicitly in this class. Note that while
    /// Scale With Document and Align With Margins default to unset, Different First, and Different Even/Odd
    /// are updated automatically as headers and footers are added and removed.
    /// </summary>
    public class XSSFHeaderFooterProperties
    {
        private CT_HeaderFooter headerFooter;

        /// <summary>
        /// Create an instance of XSSFHeaderFooterProperties from the supplied XML bean
        /// </summary>
        /// <param name="headerFooter"></param>
        public XSSFHeaderFooterProperties(CT_HeaderFooter headerFooter)
        {
            this.headerFooter = headerFooter;
        }

        /// <summary>
        /// Returns the underlying CTHeaderFooter xml bean
        /// </summary>
        /// <returns>the underlying CTHeaderFooter xml bean</returns>
        public CT_HeaderFooter HeaderFooter
        {
            get
            {
                return this.headerFooter;
            }
        }

        /// <summary>
        /// get or set alignWithMargins attribute
        /// </summary>
        public bool AlignWithMargins
        {
            get
            {
                return HeaderFooter.IsSetAlignWithMargins() ? HeaderFooter.alignWithMargins : true;
            }
            set
            {
                HeaderFooter.alignWithMargins = value;
            }
        }

        /// <summary>
        /// get or set differentFirst attribute
        /// </summary>
        public bool DifferentFirst
        {
            get
            {
                return HeaderFooter.IsSetDifferentFirst() ? HeaderFooter.differentFirst : false;
            }
            set
            {
                HeaderFooter.differentFirst = value;
            }
        }

        /// <summary>
        /// get or set differentOddEven attribute
        /// </summary>
        public bool DifferentOddEven
        {
            get
            {
                return HeaderFooter.IsSetDifferentOddEven() ? HeaderFooter.differentOddEven : false;
            }
            set
            {
                HeaderFooter.differentOddEven = value;
            }
        }

        /// <summary>
        /// get or set scaleWithDoc attribute
        /// </summary>
        public bool ScaleWithDoc
        {
            get
            {
                return HeaderFooter.IsSetScaleWithDoc() ? HeaderFooter.scaleWithDoc : true;
            }
            set
            {
                HeaderFooter.scaleWithDoc = value;
            }
        }

        /// <summary>
        /// remove alignWithMargins attribute
        /// </summary>
        public void RemoveAlignWithMargins()
        {
            if(HeaderFooter.IsSetAlignWithMargins())
            {
                HeaderFooter.UnsetAlignWithMargins();
            }
        }

        /// <summary>
        /// remove differentFirst attribute
        /// </summary>
        public void RemoveDifferentFirst()
        {
            if(HeaderFooter.IsSetDifferentFirst())
            {
                HeaderFooter.UnsetDifferentFirst();
            }
        }

        /// <summary>
        /// remove differentOddEven attribute
        /// </summary>
        public void RemoveDifferentOddEven()
        {
            if(HeaderFooter.IsSetDifferentOddEven())
            {
                HeaderFooter.UnsetDifferentOddEven();
            }
        }

        /// <summary>
        /// remove scaleWithDoc attribute
        /// </summary>
        public void RemoveScaleWithDoc()
        {
            if(HeaderFooter.IsSetScaleWithDoc())
            {
                HeaderFooter.UnsetScaleWithDoc();
            }
        }
    }
}


