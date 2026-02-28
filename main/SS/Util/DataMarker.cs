/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

using System;
using NPOI.SS.UserModel;
namespace NPOI.SS.Util
{
    /// <summary>
    /// Represents data marker used in charts.
    /// @author Roman Kashitsyn
    /// </summary>
    public class DataMarker
    {

        private ISheet sheet;
        private CellRangeAddress range;

        /// <summary>
        /// constructor
        /// </summary>
        /// <param name="sheet">the sheet where data located.</param>
        /// <param name="range">the range within that sheet.</param>
        public DataMarker(ISheet sheet, CellRangeAddress range)
        {
            this.sheet = sheet;
            this.range = range;
        }

        /// <summary>
        /// get or set the sheet marker points to.
        /// </summary>
        public ISheet Sheet
        {
            get
            {
                return sheet;
            }
            set
            {
                this.sheet = value;
            }
        }


        /// <summary>
        /// get or set range of the marker.
        /// </summary>
        public CellRangeAddress Range
        {
            get
            {
                return range;
            }
            set
            {
                range = value;
            }
        }

        /// <summary>
        /// Formats data marker using canonical format, for example
        /// 'SheetName!$A$1:$A$5'.
        /// </summary>
        /// <returns>formatted data marker</returns>
        public String FormatAsString()
        {
            String sheetName = (sheet == null) ? (null) : (sheet.SheetName);
            if (range == null)
            {
                return null;
            }
            else
            {
                return range.FormatAsString(sheetName, true);
            }
        }
    }
}