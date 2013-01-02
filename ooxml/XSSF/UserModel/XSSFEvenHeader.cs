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

using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.XSSF.UserModel.Extensions;
using NPOI.SS.UserModel;
using System;
namespace NPOI.XSSF.UserModel
{

    /**
     * <p>
     * Even page header value. Corresponds to even printed pages.
     * Even page(s) in the sheet may not be printed, for example, if the print area is specified to be 
     * a range such that it falls outside an even page's scope.
     * If no even header is specified, then odd header value is assumed for even page headers.
     *</p>
     *
     */
    public class XSSFEvenHeader : XSSFHeaderFooter, IHeader
    {

        /**
         * Create an instance of XSSFEvenHeader from the supplied XML bean
         * @see XSSFSheet#GetEvenHeader()
         * @param headerFooter
         */
        public XSSFEvenHeader(CT_HeaderFooter headerFooter)
            : base(headerFooter)
        {

            headerFooter.differentOddEven = true;
        }

        /**
         * Get the content text representing this header
         * @return text
         */
        public override String Text
        {
            get
            {
                return GetHeaderFooter().evenHeader;
            }
            set
            {
                GetHeaderFooter().evenHeader = value;
            }
        }

    }
}

