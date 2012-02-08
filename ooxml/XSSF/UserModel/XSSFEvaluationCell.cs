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

namespace NPOI.xssf.usermodel;

using NPOI.ss.formula.EvaluationCell;
using NPOI.ss.formula.EvaluationSheet;

/**
 * XSSF wrapper for a cell under Evaluation
 * 
 * @author Josh Micich
 */
final class XSSFEvaluationCell : EvaluationCell {

	private EvaluationSheet _EvalSheet;
	private XSSFCell _cell;

	public XSSFEvaluationCell(XSSFCell cell, XSSFEvaluationSheet EvaluationSheet) {
		_cell = cell;
		_EvalSheet = EvaluationSheet;
	}

	public XSSFEvaluationCell(XSSFCell cell) {
		this(cell, new XSSFEvaluationSheet(cell.Sheet));
	}

	public Object GetIdentityKey() {
		// save memory by just using the cell itself as the identity key
		// Note - this assumes HSSFCell has not overridden hashCode and Equals
		return _cell;
	}

	public XSSFCell GetXSSFCell() {
		return _cell;
	}
	public bool GetBooleanCellValue() {
		return _cell.GetBooleanCellValue();
	}
	public int GetCellType() {
		return _cell.GetCellType();
	}
	public int ColumnIndex {
		return _cell.ColumnIndex;
	}
	public int GetErrorCellValue() {
		return _cell.GetErrorCellValue();
	}
	public double GetNumericCellValue() {
		return _cell.GetNumericCellValue();
	}
	public int RowIndex {
		return _cell.RowIndex;
	}
	public EvaluationSheet Sheet {
		return _EvalSheet;
	}
	public String GetStringCellValue() {
		return _cell.GetRichStringCellValue().GetString();
	}
}


