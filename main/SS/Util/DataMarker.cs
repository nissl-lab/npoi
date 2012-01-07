/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for Additional information regarding copyright ownership.
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

using NPOI.SS.UserModel;
using System;
namespace NPOI.SS.Util
{

    /**
     * Represents data marker used in charts.
     * @author Roman Kashitsyn
     */
    public class DataMarker
    {

        private ISheet sheet;
        private CellRangeAddress range;

        /**
         * @param sheet the sheet where data located.
         * @param range the range within that sheet.
         */
        public DataMarker(ISheet sheet, CellRangeAddress range)
        {
            this.sheet = sheet;
            this.range = range;
        }

        /**
         * Returns the sheet marker points to.
         * @return sheet marker points to.
         */
        public ISheet Sheet
        {
            get
            {
                return sheet;
            }
        }

        /**
         * Sets sheet marker points to.
         * @param sheet new sheet for the marker.
         */
        public void SetSheet(ISheet sheet)
        {
            this.sheet = sheet;
        }

        /**
         * Returns range of the marker.
         * @return range of cells marker points to.
         */
        public CellRangeAddress GetRange()
        {
            return range;
        }

        /**
         * Sets range of the marker.
         * @param range new range for the marker.
         */
        public void SetRange(CellRangeAddress range)
        {
            this.range = range;
        }

        /**
         * Formats data marker using canonical format, for example
         * `SheetName!$A$1:$A$5'.
         * @return formatted data marker.
         */
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

