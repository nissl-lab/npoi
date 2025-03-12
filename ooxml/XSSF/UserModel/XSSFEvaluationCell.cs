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
using NPOI.SS.Util;

namespace NPOI.XSSF.UserModel
{

    /**
     * XSSF wrapper for a cell under Evaluation
     * 
     * @author Josh Micich
     */
    public class XSSFEvaluationCell : IEvaluationCell
    {

        private readonly IEvaluationSheet _evalSheet;
        private readonly XSSFCell _cell;

        public XSSFEvaluationCell(ICell cell, XSSFEvaluationSheet EvaluationSheet)
        {
            _cell = (XSSFCell)cell;
            _evalSheet = EvaluationSheet;
        }

        public XSSFEvaluationCell(ICell cell)
            : this(cell, new XSSFEvaluationSheet(cell.Sheet))
        {

        }

        public XSSFEvaluationCell()
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
        public virtual bool BooleanCellValue
        {
            get
            {
                return _cell.BooleanCellValue;
            }
        }
        public virtual CellType CellType
        {
            get
            {
                return _cell.CellType;
            }
        }
        public virtual int ColumnIndex
        {
            get
            {
                return _cell.ColumnIndex;
            }
        }
        public virtual int ErrorCellValue
        {
            get
            {
                return _cell.ErrorCellValue;
            }
        }
        public virtual double NumericCellValue
        {
            get
            {
                return _cell.NumericCellValue;
            }
        }
        public virtual int RowIndex
        {
            get
            {
                return _cell.RowIndex;
            }
        }
        public virtual IEvaluationSheet Sheet
        {
            get
            {
                return _evalSheet;
            }
        }
        public virtual String StringCellValue
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

            get { return _cell.ArrayFormulaRange; }
        }

        public virtual CellType CachedFormulaResultType
        {
            get { return _cell.CachedFormulaResultType; }
        }
    }
}
