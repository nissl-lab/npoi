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
using System;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace NPOI.XSSF.Streaming
{
    public class SXSSFEvaluationCell : IEvaluationCell 
    {
        private readonly SXSSFEvaluationSheet _evalSheet;
        private readonly SXSSFCell _cell;

        public SXSSFEvaluationCell(SXSSFCell cell, SXSSFEvaluationSheet evaluationSheet)
        {
            _cell = cell;
            _evalSheet = evaluationSheet;
        }

        public SXSSFEvaluationCell(SXSSFCell cell)
            : this(cell, new SXSSFEvaluationSheet(cell.Sheet as SXSSFSheet))
        {
        }


        public Object IdentityKey
        {
            get
            {
                // save memory by just using the cell itself as the identity key
                // Note - this assumes SXSSFCell has not overridden hashCode and equals
                return _cell;
            }
        }

        public SXSSFCell GetSXSSFCell()
        {
            return _cell;
        }

        public bool BooleanCellValue
        {
            get
            {
                return _cell.BooleanCellValue;
            }
        }
        /**
         * Will return {@link CellType} in a future version of POI.
         * For forwards compatibility, do not hard-code cell type literals in your code.
         *
         * @return cell type
         */

        public CellType CellType
        {
            get
            {
                return _cell.CellType;
            }
        }
        /**
         * @since POI 3.15 beta 3
         * @deprecated POI 3.15 beta 3.
         * Will be deleted when we make the CellType enum transition. See bug 59791.
         */

        public CellType CellTypeEnum
        {
            get
            {
                return _cell.CellType;
            }
        }

        public int ColumnIndex
        {
            get
            {
                return _cell.ColumnIndex;
            }
            
        }

        public int ErrorCellValue
        {
            get
            {
                return _cell.ErrorCellValue;
            }
            
        }

        public double NumericCellValue
        {
            get
            {
                return _cell.NumericCellValue;
            }
        }

        public int RowIndex
        {
            get
            {
                return _cell.RowIndex;
            }
        }

        public IEvaluationSheet Sheet
        {
            get
            {
                return _evalSheet;
            }
        }

        public String StringCellValue
        {
            get
            {
                return _cell.RichStringCellValue.String;
            }
        }
        public bool IsPartOfArrayFormulaGroup
        {
            get
            {
                return _cell.IsPartOfArrayFormulaGroup;
            }
        }

        public CellRangeAddress ArrayFormulaRange
        {
            get 
            { 
                return _cell.ArrayFormulaRange; 
            }
        }
        /// <summary>
        /// Will return CellType in a future version of POI. For forwards compatibility, do not hard-code cell type literals in your code.
        /// </summary>
        public CellType CachedFormulaResultType
        {
            get
            {
                return _cell.CachedFormulaResultType;
            }
        }

        [Obsolete("Will be removed at NPOI 2.8, Use CachedFormulaResultType instead.")]
        public CellType GetCachedFormulaResultTypeEnum()
        {
            return _cell.GetCachedFormulaResultTypeEnum();
        }
    }
}
