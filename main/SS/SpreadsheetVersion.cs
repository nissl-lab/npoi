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
using NPOI.SS.Util;

namespace NPOI.SS
{
    /// <summary>
    /// This enum allows spReadsheets from multiple Excel versions to be handled by the common code.
    /// Properties of this enum correspond to attributes of the <i>spReadsheet</i> that are easily
    /// discernable to the user.  It is not intended to deal with low-level issues like file formats.
    /// </summary>
    public class SpreadsheetVersion
    {
        /// <summary>
        /// Excel97 format aka BIFF8
        /// <list type="bullet">
        /// <item><description>The total number of available columns is 256 (2^8)</li></description></item>
        /// <item><description>The total number of available rows is 64k (2^16)</li></description></item>
        /// <item><description>The maximum number of arguments to a function is 30</li></description></item>
        /// <item><description>Number of conditional format conditions on a cell is 3</li></description></item>
        /// <item><description>Length of text cell contents is unlimited </li></description></item>
        /// <item><description>Length of text cell contents is 32767</li></description></item>
        /// </list>
        /// </summary>
        public static SpreadsheetVersion EXCEL97 = new SpreadsheetVersion("xls", 0x10000, 0x0100, 30, 3, 4000, 32767, "EXCEL97");

        /// <summary>
        /// <para>
        /// Excel2007
        /// </para>
        /// <para>
        /// <list type="bullet">
        /// <item><description>The total number of available columns is 16K (2^14)</li></description></item>
        /// <item><description>The total number of available rows is 1M (2^20)</li></description></item>
        /// <item><description>The maximum number of arguments to a function is 255</li></description></item>
        /// <item><description>Number of conditional format conditions on a cell is unlimited
        /// (actually limited by available memory in Excel)</li></description></item>
        /// <item><description>Length of text cell contents is unlimited </li></description></item>
        /// </list>
        /// </para>
        /// </summary>
        public static SpreadsheetVersion EXCEL2007 = new SpreadsheetVersion("xlsx", 0x100000, 0x4000, 255, Int32.MaxValue, 64000, 32767, "EXCEL2007");

        private string _defaultExtension;
        private int _maxRows;
        private int _maxColumns;
        private int _maxFunctionArgs;
        private int _maxCondFormats;
        private int _maxCellStyles;
        private int _maxTextLength;
        private string _name;


        private SpreadsheetVersion(string defaultExtension, int maxRows, int maxColumns, int maxFunctionArgs, int maxCondFormats, int maxCellStyles, int maxText, string name)
        {
            _defaultExtension = defaultExtension;
            _maxRows = maxRows;
            _maxColumns = maxColumns;
            _maxFunctionArgs = maxFunctionArgs;
            _maxCondFormats = maxCondFormats;
            _maxCellStyles = maxCellStyles;
            _maxTextLength = maxText;
            _name = name;
        }

        public string Name
        {
            get
            {
                return _name;
            }
        }

        /// <summary>
        /// </summary>
        /// <return>the default file extension of spReadsheet</return>
        public string DefaultExtension
        {
            get
            {
                return _defaultExtension;
            }
        }

        /// <summary>
        /// </summary>
        /// <return>the maximum number of usable rows in each spReadsheet</return>
        public int MaxRows
        {
            get
            {
                return _maxRows;
            }
        }

        /// <summary>
        /// </summary>
        /// <return>the last (maximum) valid row index, equals to <c> GetMaxRows() - 1 </c></return>
        public int LastRowIndex
        {
            get
            {
                return _maxRows - 1;
            }
        }

        /// <summary>
        /// </summary>
        /// <return>the maximum number of usable columns in each spReadsheet</return>
        public int MaxColumns
        {
            get
            {
                return _maxColumns;
            }
        }

        /// <summary>
        /// </summary>
        /// <return>the last (maximum) valid column index, equals to <c> GetMaxColumns() - 1 </c></return>
        public int LastColumnIndex
        {
            get
            {
                return _maxColumns - 1;
            }
        }

        /// <summary>
        /// </summary>
        /// <return>the maximum number arguments that can be passed to a multi-arg function (e.g. COUNTIF)</return>
        public int MaxFunctionArgs
        {
            get
            {
                return _maxFunctionArgs;
            }
        }

        /// <summary>
        /// </summary>
        /// <return>the maximum number of conditional format conditions on a cell</return>
        public int MaxConditionalFormats
        {
            get
            {
                return _maxCondFormats;
            }
        }

        /// <summary>
        /// </summary>
        /// <return>the last valid column index in a ALPHA-26 representation
        /// (<c>IV</c> or <c>XFD</c>).
        /// </return>
        public String LastColumnName
        {
            get
            {
                return CellReference.ConvertNumToColString(LastColumnIndex);
            }
        }
        /// <summary>
        /// </summary>
        /// <return>the maximum number of cell styles per spreadsheet</return>
        public int MaxCellStyles
        {
            get
            {
                return _maxCellStyles;
            }
        }
        /// <summary>
        /// </summary>
        /// <return>the maximum length of a text cell</return>
        public int MaxTextLength
        {
            get
            {
                return _maxTextLength;
            }
        }

    }

}
