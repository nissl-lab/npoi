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

namespace NPOI.SS.Formula.Eval;

using junit.framework.AssertionFailedError;
using junit.framework.TestCase;

using NPOI.SS.Formula.PTG.AreaI;
using NPOI.SS.Formula.PTG.AreaI.OffsetArea;
using NPOI.hssf.UserModel.HSSFCell;
using NPOI.hssf.UserModel.HSSFFormulaEvaluator;
using NPOI.hssf.UserModel.HSSFRow;
using NPOI.hssf.UserModel.HSSFWorkbook;
using NPOI.hssf.Util.AreaReference;
using NPOI.hssf.Util.CellReference;
using NPOI.SS.Formula.TwoDEval;
using NPOI.SS.UserModel.CellValue;

/**
 * Test for unary plus operator Evaluator.
 *
 * @author Josh Micich
 */
public class TestRangeEval  {

	public void TestPermutations() {

		Confirm("B3", "D7", "B3:D7");
		Confirm("B1", "B1", "B1:B1");

		Confirm("B7", "D3", "B3:D7");
		Confirm("D3", "B7", "B3:D7");
		Confirm("D7", "B3", "B3:D7");
	}

	private static void Confirm(String refA, String refB, String expectedAreaRef) {

		ValueEval[] args = {
			CreateRefEval(refA),
			CreateRefEval(refB),
		};
		AreaReference ar = new AreaReference(expectedAreaRef);
		ValueEval result = EvalInstances.Range.Evaluate(args, 0, (short)0);
		Assert.IsTrue(result is AreaEval);
		AreaEval ae = (AreaEval) result;
		Assert.AreEqual(ar.GetFirstCell().Row, ae.FirstRow);
		Assert.AreEqual(ar.GetLastCell().Row, ae.LastRow);
		Assert.AreEqual(ar.GetFirstCell().Col, ae.FirstColumn);
		Assert.AreEqual(ar.GetLastCell().Col, ae.LastColumn);
	}

	private static ValueEval CreateRefEval(String refStr) {
		CellReference cr = new CellReference(refStr);
		return new MockRefEval(cr.Row, cr.Col);

	}

	private static class MockRefEval : RefEvalBase {

		public MockRefEval(int rowIndex, int columnIndex) {
			base(rowIndex, columnIndex);
		}
		public ValueEval InnerValueEval {
			throw new RuntimeException("not expected to be called during this Test");
		}
		public AreaEval offset(int relFirstRowIx, int relLastRowIx, int relFirstColIx,
				int relLastColIx) {
			AreaI area = new OffsetArea(getRow(), Column,
					relFirstRowIx, relLastRowIx, relFirstColIx, relLastColIx);
			return new MockAreaEval(area);
		}
	}

	private static class MockAreaEval : AreaEvalBase {

		public MockAreaEval(AreaI ptg) {
			base(ptg);
		}
		private MockAreaEval(int firstRow, int firstColumn, int lastRow, int lastColumn) {
			base(firstRow, firstColumn, lastRow, lastColumn);
		}
		public ValueEval GetRelativeValue(int relativeRowIndex, int relativeColumnIndex) {
			throw new RuntimeException("not expected to be called during this Test");
		}
		public AreaEval offset(int relFirstRowIx, int relLastRowIx, int relFirstColIx,
				int relLastColIx) {
			AreaI area = new OffsetArea(getFirstRow(), FirstColumn,
					relFirstRowIx, relLastRowIx, relFirstColIx, relLastColIx);

			return new MockAreaEval(area);
		}
		public TwoDEval GetRow(int rowIndex) {
			if (rowIndex >= Height) {
				throw new ArgumentException("Invalid rowIndex " + rowIndex
						+ ".  Allowable range is (0.." + Height + ").");
			}
			return new MockAreaEval(rowIndex, FirstColumn, rowIndex, LastColumn);
		}
		public TwoDEval GetColumn(int columnIndex) {
			if (columnIndex >= Width) {
				throw new ArgumentException("Invalid columnIndex " + columnIndex
						+ ".  Allowable range is (0.." + Width + ").");
			}
			return new MockAreaEval(getFirstRow(), columnIndex, LastRow, columnIndex);
		}
	}

	public void TestRangeUsingOffsetFunc_bug46948() {
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFRow row = wb.CreateSheet("Sheet1").createRow(0);
		HSSFCell cellA1 = row.CreateCell(0);
		HSSFCell cellB1 = row.CreateCell(1);
		row.CreateCell(2).SetCellValue(5.0); // C1
		row.CreateCell(3).SetCellValue(7.0); // D1
		row.CreateCell(4).SetCellValue(9.0); // E1


		cellA1.SetCellFormula("SUM(C1:OFFSET(C1,0,B1))");

		cellB1.SetCellValue(1.0); // range will be C1:D1

		HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
		CellValue cv;
		try {
			cv = fe.Evaluate(cellA1);
		} catch (ArgumentException e) {
			if (e.GetMessage().Equals("Unexpected ref arg class (NPOI.SS.Formula.LazyAreaEval)")) {
				throw new AssertionFailedError("Identified bug 46948");
			}
			throw e;
		}

		Assert.AreEqual(12.0, cv.GetNumberValue(), 0.0);

		cellB1.SetCellValue(2.0); // range will be C1:E1
		fe.NotifyUpdateCell(cellB1);
		cv = fe.Evaluate(cellA1);
		Assert.AreEqual(21.0, cv.GetNumberValue(), 0.0);

		cellB1.SetCellValue(0.0); // range will be C1:C1
		fe.NotifyUpdateCell(cellB1);
		cv = fe.Evaluate(cellA1);
		Assert.AreEqual(5.0, cv.GetNumberValue(), 0.0);
	}
}

