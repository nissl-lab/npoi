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

using NPOI.hssf.HSSFTestDataSamples;
using NPOI.hssf.UserModel.HSSFCell;
using NPOI.hssf.UserModel.HSSFFormulaEvaluator;
using NPOI.hssf.UserModel.HSSFRow;
using NPOI.hssf.UserModel.HSSFSheet;
using NPOI.hssf.UserModel.HSSFWorkbook;
using NPOI.SS.UserModel.CellValue;

/**
 * Miscellaneous Tests for bugzilla entries.<p/> The Test name Contains the
 * bugzilla bug id.
 * 
 * 
 * @author Josh Micich
 */
public class TestFormulaBugs  {

	/**
	 * Bug 27349 - VLOOKUP with reference to another sheet.<p/> This Test was
	 * Added <em>long</em> After the relevant functionality was fixed.
	 */
	public void Test27349() {
		// 27349-vLookupAcrossSheets.xls is bugzilla/attachment.cgi?id=10622
		InputStream is = HSSFTestDataSamples.OpenSampleFileStream("27349-vLookupAcrossSheets.xls");
		HSSFWorkbook wb;
		try {
			// original bug may have thrown exception here, or output warning to
			// stderr
			wb = new HSSFWorkbook(is);
		} catch (IOException e) {
			throw new RuntimeException(e);
		}

		HSSFSheet sheet = wb.GetSheetAt(0);
		HSSFRow row = sheet.GetRow(1);
		HSSFCell cell = row.GetCell(0);

		// this defInitely would have failed due to 27349
		Assert.AreEqual("VLOOKUP(1,'DATA TABLE'!$A$8:'DATA TABLE'!$B$10,2)", cell
				.GetCellFormula());

		// We might as well Evaluate the formula
		HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
		CellValue cv = fe.Evaluate(cell);

		Assert.AreEqual(HSSFCell.CELL_TYPE_NUMERIC, cv.GetCellType());
		Assert.AreEqual(3.0, cv.GetNumberValue(), 0.0);
	}

	/**
	 * Bug 27405 - isnumber() formula always Evaluates to false in if statement<p/>
	 * 
	 * seems to be a duplicate of 24925
	 */
	public void Test27405() {

		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.CreateSheet("input");
		// input row 0
		HSSFRow row = sheet.CreateRow(0);
		HSSFCell cell = row.CreateCell(0);
		cell = row.CreateCell(1);
		cell.SetCellValue(1); // B1
		// input row 1
		row = sheet.CreateRow(1);
		cell = row.CreateCell(1);
		cell.SetCellValue(999); // B2

		int rno = 4;
		row = sheet.CreateRow(rno);
		cell = row.CreateCell(1); // B5
		cell.SetCellFormula("isnumber(b1)");
		cell = row.CreateCell(3); // D5
		cell.SetCellFormula("IF(ISNUMBER(b1),b1,b2)");

		if (false) { // Set true to check excel file manually
			// bug report mentions 'Editing the formula in excel "fixes" the problem.'
			try {
				FileOutputStream fileOut = new FileOutputStream("27405output.xls");
				wb.Write(fileOut);
				fileOut.Close();
			} catch (IOException e) {
				throw new RuntimeException(e);
			}
		}
		
		// use POI's Evaluator as an extra sanity check
		HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
		CellValue cv;
		cv = fe.Evaluate(cell);
		Assert.AreEqual(HSSFCell.CELL_TYPE_NUMERIC, cv.GetCellType());
		Assert.AreEqual(1.0, cv.GetNumberValue(), 0.0);
		
		cv = fe.Evaluate(row.GetCell(1));
		Assert.AreEqual(HSSFCell.CELL_TYPE_BOOLEAN, cv.GetCellType());
		Assert.AreEqual(true, cv.GetBooleanValue());
	}

	/**
	 * Bug 42448 - Can't parse SUMPRODUCT(A!C7:A!C67, B8:B68) / B69 <p/>
	 */
	public void Test42448() {
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet1 = wb.CreateSheet("Sheet1");

		HSSFRow row = sheet1.CreateRow(0);
		HSSFCell cell = row.CreateCell(0);

		// it's important to create the referenced sheet first
		HSSFSheet sheet2 = wb.CreateSheet("A"); // note name 'A'
		// TODO - POI crashes if the formula is Added before this sheet
		// RuntimeException("Zero length string is an invalid sheet name")
		// Excel doesn't crash but the formula doesn't work until it is
		// re-entered

		String inputFormula = "SUMPRODUCT(A!C7:A!C67, B8:B68) / B69"; // as per bug report
		try {
			cell.SetCellFormula(inputFormula); 
		} catch (StringIndexOutOfBoundsException e) {
			throw new AssertionFailedError("Identified bug 42448");
		}

		Assert.AreEqual("SUMPRODUCT(A!C7:A!C67,B8:B68)/B69", cell.GetCellFormula());

		// might as well Evaluate the sucker...

		AddCell(sheet2, 5, 2, 3.0); // A!C6
		AddCell(sheet2, 6, 2, 4.0); // A!C7
		AddCell(sheet2, 66, 2, 5.0); // A!C67
		AddCell(sheet2, 67, 2, 6.0); // A!C68

		AddCell(sheet1, 6, 1, 7.0); // B7
		AddCell(sheet1, 7, 1, 8.0); // B8
		AddCell(sheet1, 67, 1, 9.0); // B68
		AddCell(sheet1, 68, 1, 10.0); // B69

		double expectedResult = (4.0 * 8.0 + 5.0 * 9.0) / 10.0;

		HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
		CellValue cv = fe.Evaluate(cell);

		Assert.AreEqual(HSSFCell.CELL_TYPE_NUMERIC, cv.GetCellType());
		Assert.AreEqual(expectedResult, cv.GetNumberValue(), 0.0);
	}

	private static void AddCell(HSSFSheet sheet, int rowIx, int colIx,
			double value) {
		sheet.CreateRow(rowIx).createCell(colIx).SetCellValue(value);
	}
}

