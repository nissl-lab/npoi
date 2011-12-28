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

package org.apache.poi.ss.formula.eval.forked;

import junit.framework.AssertionFailedError;
import junit.framework.TestCase;

import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.IStabilityClassifier;

/**
 * @author Josh Micich
 */
public final class TestForkedEvaluator extends TestCase {
	/**
	 * set up a calculation workbook with input cells nicely segregated on a
	 * sheet called "Inputs"
	 */
	private static HSSFWorkbook createWorkbook() {
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet1 = wb.createSheet("Inputs");
		HSSFSheet sheet2 = wb.createSheet("Calculations");
		HSSFRow row;
		row = sheet2.createRow(0);
		row.createCell(0).setCellFormula("B1*Inputs!A1-Inputs!B1");
		row.createCell(1).setCellValue(5.0); // Calculations!B1

		// some default input values
		row = sheet1.createRow(0);
		row.createCell(0).setCellValue(2.0); // Inputs!A1
		row.createCell(1).setCellValue(3.0); // Inputs!B1
		return wb;
	}

	/**
	 * Shows a basic use-case for {@link ForkedEvaluator}
	 */
	public void testBasic() {
		HSSFWorkbook wb = createWorkbook();

		// The stability classifier is useful to reduce memory consumption of caching logic
		IStabilityClassifier stabilityClassifier = new IStabilityClassifier() {
			public boolean isCellFinal(int sheetIndex, int rowIndex, int columnIndex) {
				return sheetIndex == 1;
			}
		};

		ForkedEvaluator fe1 = ForkedEvaluator.create(wb, stabilityClassifier, null);
		ForkedEvaluator fe2 = ForkedEvaluator.create(wb, stabilityClassifier, null);

		// fe1 and fe2 can be used concurrently on separate threads

		fe1.updateCell("Inputs", 0, 0, new NumberEval(4.0));
		fe1.updateCell("Inputs", 0, 1, new NumberEval(1.1));

		fe2.updateCell("Inputs", 0, 0, new NumberEval(1.2));
		fe2.updateCell("Inputs", 0, 1, new NumberEval(2.0));

		assertEquals(18.9, ((NumberEval) fe1.evaluate("Calculations", 0, 0)).getNumberValue(), 0.0);
		assertEquals(4.0, ((NumberEval) fe2.evaluate("Calculations", 0, 0)).getNumberValue(), 0.0);
		fe1.updateCell("Inputs", 0, 0, new NumberEval(3.0));
		assertEquals(13.9, ((NumberEval) fe1.evaluate("Calculations", 0, 0)).getNumberValue(), 0.0);
	}

	/**
	 * As of Sep 2009, the Forked evaluator can update values from existing cells (this is because
	 * the underlying 'master' cell is used as a key into the calculation cache.  Prior to the fix
	 * for this bug, an attempt to update a missing cell would result in NPE.  This junit tests for
	 * a more meaningful error message.<br/>
	 *
	 * An alternate solution might involve allowing empty cells to be created as necessary.  That
	 * was considered less desirable because so far, the underlying 'master' workbook is strictly
	 * <i>read-only</i> with respect to the ForkedEvaluator.
	 */
	public void testMissingInputCell() {
		HSSFWorkbook wb = createWorkbook();

		ForkedEvaluator fe = ForkedEvaluator.create(wb, null, null);

		// attempt update input at cell A2 (which is missing)
		try {
			fe.updateCell("Inputs", 1, 0, new NumberEval(4.0));
			throw new AssertionFailedError(
					"Expected exception to be thrown due to missing input cell");
		} catch (NullPointerException e) {
			StackTraceElement[] stes = e.getStackTrace();
			if (stes[0].getMethodName().equals("getIdentityKey")) {
				throw new AssertionFailedError("Identified bug with update of missing input cell");
			}
			throw e;
		} catch (UnsupportedOperationException e) {
			if (e.getMessage().equals(
					"Underlying cell 'A2' is missing in master sheet.")) {
				// expected during successful test
			} else {
				throw e;
			}
		}
	}
}
