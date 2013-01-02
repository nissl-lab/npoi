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

using NPOI.SS.Formula;
using NPOI.XSSF.UserModel;
using System;
using NPOI.SS.UserModel;
namespace NPOI.XSSF.UserModel
{

    /**
     * XSSF wrapper for a cell under Evaluation
     * 
     * @author Josh Micich
     */
    public class XSSFEvaluationCell : IEvaluationCell
    {

        private IEvaluationSheet _evalSheet;
        private XSSFCell _cell;

        public XSSFEvaluationCell(ICell cell, XSSFEvaluationSheet EvaluationSheet)
        {
            _cell = (XSSFCell)cell;
            _evalSheet = EvaluationSheet;
        }

        public XSSFEvaluationCell(ICell cell)
            : this(cell, new XSSFEvaluationSheet(cell.Sheet))
        {

        }

        public Object IdentityKey
        {
            get
            {
                // save memory by just using the cell itself as the identity key
                // Note - this assumes HSSFCell has not overridden hashCode and Equals
                return _cell;
            }
        }

        public XSSFCell GetXSSFCell()
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
        public CellType CellType
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

        #region IEvaluationCell ≥…‘±


        public CellType CachedFormulaResultType
        {
            get { return _cell.CachedFormulaResultType; }
        }

        #endregion
    }
}
