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

namespace NPOI.XSSF.Binary
{
    using NPOI.Util;

    /// <summary>
    /// This class encapsulates what the spec calls a "Cell" object.
    /// I added "Header" to clarify that this does not contain the contents
    /// of the cell, only the column number, the style id and the phonetic bool
    /// </summary>
    /// @since 3.16-beta3
    public class XSSFBCellHeader
    {
        public static int Length = 8;

        /// <summary>
        /// </summary>
        /// <param name="data">raw data</param>
        /// <param name="offset">offset at which to start reading the record</param>
        /// <param name="currentRow">0-based current row count</param>
        /// <param name="cell">cell buffer to update</param>
        public static void Parse(byte[] data, int offset, int currentRow, XSSFBCellHeader cell)
        {
            int colNum = XSSFBUtils.CastToInt(LittleEndian.GetUInt(data, offset));
            offset += LittleEndian.INT_SIZE;
            int styleIdx = XSSFBUtils.Get24BitInt(data, offset);
            offset += 3;
            //TODO: range checking
            bool showPhonetic = false;//TODO: fill this out
            cell.Reset(currentRow, colNum, styleIdx, showPhonetic);
        }

        private int rowNum;
        private int colNum;
        private int styleIdx;
        private bool showPhonetic;

        public void Reset(int rowNum, int colNum, int styleIdx, bool showPhonetic)
        {
            this.rowNum = rowNum;
            this.colNum = colNum;
            this.styleIdx = styleIdx;
            this.showPhonetic = showPhonetic;
        }

        public int ColNum => colNum;

        public int StyleIdx => styleIdx;
    }
}

