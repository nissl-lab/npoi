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

using NPOI.hssf.usermodel.HSSFFormulaEvaluator;
using NPOI.ss.formula.IStabilityClassifier;
using NPOI.ss.formula.WorkbookEvaluator;
using NPOI.ss.formula.Eval.BoolEval;
using NPOI.ss.formula.Eval.ErrorEval;
using NPOI.ss.formula.Eval.NumberEval;
using NPOI.ss.formula.Eval.StringEval;
using NPOI.ss.formula.Eval.ValueEval;
using NPOI.ss.formula.udf.UDFFinder;
using NPOI.ss.usermodel.Cell;
using NPOI.ss.usermodel.CellValue;
using NPOI.ss.usermodel.FormulaEvaluator;
using NPOI.ss.usermodel.Workbook;

/**
 * Evaluates formula cells.<p/>
 *
 * For performance reasons, this class keeps a cache of all previously calculated intermediate
 * cell values.  Be sure to call {@link #ClearAllCachedResultValues()} if any workbook cells are Changed between
 * calls to Evaluate~ methods on this class.
 *
 * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
 * @author Josh Micich
 */
public class XSSFFormulaEvaluator : FormulaEvaluator {

	private WorkbookEvaluator _bookEvaluator;
	private XSSFWorkbook _book;

	public XSSFFormulaEvaluator(XSSFWorkbook workbook) {
		this(workbook, null, null);
	}
	/**
	 * @param stabilityClassifier used to optimise caching performance. Pass <code>null</code>
	 * for the (conservative) assumption that any cell may have its defInition Changed After
	 * Evaluation begins.
	 * @deprecated (Sep 2009) (reduce overloading) use {@link #Create(XSSFWorkbook, NPOI.ss.formula.IStabilityClassifier, NPOI.ss.formula.udf.UDFFinder)}
	 */
    @Deprecated
    public XSSFFormulaEvaluator(XSSFWorkbook workbook, IStabilityClassifier stabilityClassifier) {
		_bookEvaluator = new WorkbookEvaluator(XSSFEvaluationWorkbook.Create(workbook), stabilityClassifier, null);
		_book = workbook;
	}
	private XSSFFormulaEvaluator(XSSFWorkbook workbook, IStabilityClassifier stabilityClassifier, UDFFinder udfFinder) {
		_bookEvaluator = new WorkbookEvaluator(XSSFEvaluationWorkbook.Create(workbook), stabilityClassifier, udfFinder);
      _book = workbook;
	}

	/**
	 * @param stabilityClassifier used to optimise caching performance. Pass <code>null</code>
	 * for the (conservative) assumption that any cell may have its defInition Changed After
	 * Evaluation begins.
	 * @param udfFinder pass <code>null</code> for default (AnalysisToolPak only)
	 */
	public static XSSFFormulaEvaluator Create(XSSFWorkbook workbook, IStabilityClassifier stabilityClassifier, UDFFinder udfFinder) {
		return new XSSFFormulaEvaluator(workbook, stabilityClassifier, udfFinder);
	}


	/**
	 * Should be called whenever there are major Changes (e.g. moving sheets) to input cells
	 * in the Evaluated workbook.
	 * Failure to call this method After changing cell values will cause incorrect behaviour
	 * of the Evaluate~ methods of this class
	 */
	public void ClearAllCachedResultValues() {
		_bookEvaluator.ClearAllCachedResultValues();
	}
	public void NotifySetFormula(Cell cell) {
		_bookEvaluator.NotifyUpdateCell(new XSSFEvaluationCell((XSSFCell)cell));
	}
	public void NotifyDeleteCell(Cell cell) {
		_bookEvaluator.NotifyDeleteCell(new XSSFEvaluationCell((XSSFCell)cell));
	}
    public void NotifyUpdateCell(Cell cell) {
        _bookEvaluator.NotifyUpdateCell(new XSSFEvaluationCell((XSSFCell)cell));
    }

	/**
	 * If cell Contains a formula, the formula is Evaluated and returned,
	 * else the CellValue simply copies the appropriate cell value from
	 * the cell and also its cell type. This method should be preferred over
	 * EvaluateInCell() when the call should not modify the contents of the
	 * original cell.
	 * @param cell
	 */
	public CellValue Evaluate(Cell cell) {
		if (cell == null) {
			return null;
		}

		switch (cell.GetCellType()) {
			case XSSFCell.CELL_TYPE_BOOLEAN:
				return CellValue.ValueOf(cell.GetBooleanCellValue());
			case XSSFCell.CELL_TYPE_ERROR:
				return CellValue.GetError(cell.GetErrorCellValue());
			case XSSFCell.CELL_TYPE_FORMULA:
				return EvaluateFormulaCellValue(cell);
			case XSSFCell.CELL_TYPE_NUMERIC:
				return new CellValue(cell.GetNumericCellValue());
			case XSSFCell.CELL_TYPE_STRING:
				return new CellValue(cell.GetRichStringCellValue().GetString());
            case XSSFCell.CELL_TYPE_BLANK:
                return null;
		}
		throw new InvalidOperationException("Bad cell type (" + cell.GetCellType() + ")");
	}


	/**
	 * If cell Contains formula, it Evaluates the formula,
	 *  and saves the result of the formula. The cell
	 *  remains as a formula cell.
	 * Else if cell does not contain formula, this method leaves
	 *  the cell unChanged.
	 * Note that the type of the formula result is returned,
	 *  so you know what kind of value is also stored with
	 *  the formula.
	 * <pre>
	 * int EvaluatedCellType = Evaluator.EvaluateFormulaCell(cell);
	 * </pre>
	 * Be aware that your cell will hold both the formula,
	 *  and the result. If you want the cell Replaced with
	 *  the result of the formula, use {@link #Evaluate(NPOI.ss.usermodel.Cell)} }
	 * @param cell The cell to Evaluate
	 * @return The type of the formula result (the cell's type remains as HSSFCell.CELL_TYPE_FORMULA however)
	 */
	public int EvaluateFormulaCell(Cell cell) {
		if (cell == null || cell.GetCellType() != XSSFCell.CELL_TYPE_FORMULA) {
			return -1;
		}
		CellValue cv = EvaluateFormulaCellValue(cell);
		// cell remains a formula cell, but the cached value is Changed
	 SetCellValue(cell, cv);
		return cv.GetCellType();
	}

	/**
	 * If cell Contains formula, it Evaluates the formula, and
	 *  Puts the formula result back into the cell, in place
	 *  of the old formula.
	 * Else if cell does not contain formula, this method leaves
	 *  the cell unChanged.
	 * Note that the same instance of HSSFCell is returned to
	 * allow chained calls like:
	 * <pre>
	 * int EvaluatedCellType = Evaluator.EvaluateInCell(cell).GetCellType();
	 * </pre>
	 * Be aware that your cell value will be Changed to hold the
	 *  result of the formula. If you simply want the formula
	 *  value computed for you, use {@link #EvaluateFormulaCell(NPOI.ss.usermodel.Cell)} }
	 * @param cell
	 */
	public XSSFCell EvaluateInCell(Cell cell) {
		if (cell == null) {
			return null;
		}
		XSSFCell result = (XSSFCell) cell;
		if (cell.GetCellType() == XSSFCell.CELL_TYPE_FORMULA) {
			CellValue cv = EvaluateFormulaCellValue(cell);
		 SetCellType(cell, cv); // cell will no longer be a formula cell
		 SetCellValue(cell, cv);
		}
		return result;
	}
	private static void SetCellType(Cell cell, CellValue cv) {
		int cellType = cv.GetCellType();
		switch (cellType) {
			case XSSFCell.CELL_TYPE_BOOLEAN:
			case XSSFCell.CELL_TYPE_ERROR:
			case XSSFCell.CELL_TYPE_NUMERIC:
			case XSSFCell.CELL_TYPE_STRING:
				cell.SetCellType(cellType);
				return;
			case XSSFCell.CELL_TYPE_BLANK:
				// never happens - blanks eventually Get translated to zero
			case XSSFCell.CELL_TYPE_FORMULA:
				// this will never happen, we have already Evaluated the formula
		}
		throw new InvalidOperationException("Unexpected cell value type (" + cellType + ")");
	}

	private static void SetCellValue(Cell cell, CellValue cv) {
		int cellType = cv.GetCellType();
		switch (cellType) {
			case XSSFCell.CELL_TYPE_BOOLEAN:
				cell.SetCellValue(cv.GetBooleanValue());
				break;
			case XSSFCell.CELL_TYPE_ERROR:
				cell.SetCellErrorValue(cv.GetErrorValue());
				break;
			case XSSFCell.CELL_TYPE_NUMERIC:
				cell.SetCellValue(cv.GetNumberValue());
				break;
			case XSSFCell.CELL_TYPE_STRING:
				cell.SetCellValue(new XSSFRichTextString(cv.StringValue));
				break;
			case XSSFCell.CELL_TYPE_BLANK:
				// never happens - blanks eventually Get translated to zero
			case XSSFCell.CELL_TYPE_FORMULA:
				// this will never happen, we have already Evaluated the formula
			default:
				throw new InvalidOperationException("Unexpected cell value type (" + cellType + ")");
		}
	}

	/**
	 * Loops over all cells in all sheets of the supplied
	 *  workbook.
	 * For cells that contain formulas, their formulas are
	 *  Evaluated, and the results are saved. These cells
	 *  remain as formula cells.
	 * For cells that do not contain formulas, no Changes
	 *  are made.
	 * This is a helpful wrapper around looping over all
	 *  cells, and calling EvaluateFormulaCell on each one.
	 */
	public static void EvaluateAllFormulaCells(XSSFWorkbook wb) {
	   HSSFFormulaEvaluator.EvaluateAllFormulaCells((Workbook)wb);
	}
   /**
    * Loops over all cells in all sheets of the supplied
    *  workbook.
    * For cells that contain formulas, their formulas are
    *  Evaluated, and the results are saved. These cells
    *  remain as formula cells.
    * For cells that do not contain formulas, no Changes
    *  are made.
    * This is a helpful wrapper around looping over all
    *  cells, and calling EvaluateFormulaCell on each one.
    */
   public void EvaluateAll() {
      HSSFFormulaEvaluator.EvaluateAllFormulaCells(_book);
   }

	/**
	 * Returns a CellValue wrapper around the supplied ValueEval instance.
	 */
	private CellValue EvaluateFormulaCellValue(Cell cell) {
        if(!(cell is XSSFCell)){
            throw new ArgumentException("Unexpected type of cell: " + cell.GetType() + "." +
                    " Only XSSFCells can be Evaluated.");
        }

		ValueEval eval = _bookEvaluator.Evaluate(new XSSFEvaluationCell((XSSFCell) cell));
		if (eval is NumberEval) {
			NumberEval ne = (NumberEval) Eval;
			return new CellValue(ne.GetNumberValue());
		}
		if (eval is BoolEval) {
			BoolEval be = (BoolEval) Eval;
			return CellValue.ValueOf(be.GetBooleanValue());
		}
		if (eval is StringEval) {
			StringEval ne = (StringEval) Eval;
			return new CellValue(ne.StringValue);
		}
		if (eval is ErrorEval) {
			return CellValue.GetError(((ErrorEval)Eval).GetErrorCode());
		}
		throw new RuntimeException("Unexpected eval class (" + Eval.GetType().GetName() + ")");
	}
}


