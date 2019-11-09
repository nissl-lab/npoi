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

namespace TestCases.HSSF.Record
{
    using System;
    using NUnit.Framework;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.UserModel;
    using NPOI.SS;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using TestCases.Exceptions;
    using TestCases.HSSF;
    using TestCases.HSSF.UserModel;

    /**
     * @author Josh Micich
     */
    [TestFixture]
    public class TestSharedFormulaRecord
    {

        /**
         * A sample spreadsheet known to have one sheet with 4 shared formula ranges
         */
        private static String SHARED_FORMULA_TEST_XLS = "SharedFormulaTest.xls";
        /**
         * Binary data for an encoded formula.  Taken from attachment 22062 (bugzilla 45123/45421).
         * The shared formula is in Sheet1!C6:C21, with text "SUMPRODUCT(--(End_Acct=$C6),--(End_Bal))"
         * This data is found at offset 0x1A4A (within the shared formula record).
         * The critical thing about this formula is that it Contains shared formula tokens (tRefN*,
         * tAreaN*) with operand class 'array'.
         */
        private static byte[] SHARED_FORMULA_WITH_REF_ARRAYS_DATA = {
		0x1A, 0x00,
		0x63, 0x02, 0x00, 0x00, 0x00,
		0x6C, 0x00, 0x00, 0x02, (byte)0x80,  // tRefNA
		0x0B,
		0x15,
		0x13,
		0x13,
		0x63, 0x03, 0x00, 0x00, 0x00,
		0x15,
		0x13,
		0x13,
		0x42, 0x02, (byte)0xE4, 0x00,
	};

        /**
         * The method <tt>SharedFormulaRecord.ConvertSharedFormulas()</tt> Converts formulas from
         * 'shared formula' to 'single cell formula' format.  It is important that token operand
         * classes are preserved during this transformation, because Excel may not tolerate the
         * incorrect encoding.  The formula here is one such example (Excel displays #VALUE!).
         */
        [Test]
        public void TestConvertSharedFormulasOperandClasses_bug45123()
        {

            ILittleEndianInput in1 = TestcaseRecordInputStream.CreateLittleEndian(SHARED_FORMULA_WITH_REF_ARRAYS_DATA);
            int encodedLen = in1.ReadUShort();
            Ptg[] sharedFormula = Ptg.ReadTokens(encodedLen, in1);

            SharedFormula sf = new SharedFormula(SpreadsheetVersion.EXCEL97);
            Ptg[] ConvertedFormula = sf.ConvertSharedFormulas(sharedFormula, 100, 200);

            RefPtg refPtg = (RefPtg)ConvertedFormula[1];
            Assert.AreEqual("$C101", refPtg.ToFormulaString());
            if (refPtg.PtgClass == Ptg.CLASS_REF)
            {
                throw new AssertionException("Identified bug 45123");
            }

            ConfirmOperandClasses(sharedFormula, ConvertedFormula);
        }

        private static void ConfirmOperandClasses(Ptg[] originalPtgs, Ptg[] convertedPtg)
        {
            Assert.AreEqual(originalPtgs.Length, convertedPtg.Length);
            for (int i = 0; i < convertedPtg.Length; i++)
            {
                Ptg originalPtg = originalPtgs[i];
                Ptg ConvertedPtg = convertedPtg[i];
                if (originalPtg.PtgClass != ConvertedPtg.PtgClass)
                {
                    throw new ComparisonFailure("Different operand class for token[" + i + "]",
                            originalPtg.PtgClass.ToString(), ConvertedPtg.PtgClass.ToString());
                }
            }
        }
        [Test]
        public void TestConvertSharedFormulas()
        {
            IWorkbook wb = new HSSFWorkbook();
            HSSFEvaluationWorkbook fpb = HSSFEvaluationWorkbook.Create(wb);
            Ptg[] sharedFormula, ConvertedFormula;

            SharedFormula sf = new SharedFormula(SpreadsheetVersion.EXCEL97);

            sharedFormula = FormulaParser.Parse("A2", fpb, FormulaType.Cell, -1, -1);
            ConvertedFormula = sf.ConvertSharedFormulas(sharedFormula, 0, 0);
            ConfirmOperandClasses(sharedFormula, ConvertedFormula);
            //conversion relative to [0,0] should return the original formula
            Assert.AreEqual(FormulaRenderer.ToFormulaString(fpb, ConvertedFormula), "A2");

            ConvertedFormula = sf.ConvertSharedFormulas(sharedFormula, 1, 0);
            ConfirmOperandClasses(sharedFormula, ConvertedFormula);
            //one row down
            Assert.AreEqual(FormulaRenderer.ToFormulaString(fpb, ConvertedFormula), "A3");

            ConvertedFormula = sf.ConvertSharedFormulas(sharedFormula, 1, 1);
            ConfirmOperandClasses(sharedFormula, ConvertedFormula);
            //one row down and one cell right
            Assert.AreEqual(FormulaRenderer.ToFormulaString(fpb, ConvertedFormula), "B3");

            sharedFormula = FormulaParser.Parse("SUM(A1:C1)", fpb, FormulaType.Cell, -1, -1);
            ConvertedFormula = sf.ConvertSharedFormulas(sharedFormula, 0, 0);
            ConfirmOperandClasses(sharedFormula, ConvertedFormula);
            Assert.AreEqual(FormulaRenderer.ToFormulaString(fpb, ConvertedFormula), "SUM(A1:C1)");

            ConvertedFormula = sf.ConvertSharedFormulas(sharedFormula, 1, 0);
            ConfirmOperandClasses(sharedFormula, ConvertedFormula);
            Assert.AreEqual(FormulaRenderer.ToFormulaString(fpb, ConvertedFormula), "SUM(A2:C2)");

            ConvertedFormula = sf.ConvertSharedFormulas(sharedFormula, 1, 1);
            ConfirmOperandClasses(sharedFormula, ConvertedFormula);
            Assert.AreEqual(FormulaRenderer.ToFormulaString(fpb, ConvertedFormula), "SUM(B2:D2)");
        }

        /**
         * Make sure that POI preserves {@link SharedFormulaRecord}s
         */
        [Test]
        public void TestPreserveOnReSerialize()
        {
            HSSFWorkbook wb;
            ISheet sheet;
            ICell cellB32769;
            ICell cellC32769;

            // Reading directly from XLS file
            wb = HSSFTestDataSamples.OpenSampleWorkbook(SHARED_FORMULA_TEST_XLS);
            sheet = wb.GetSheetAt(0);
            cellB32769 = sheet.GetRow(32768).GetCell(1);
            cellC32769 = sheet.GetRow(32768).GetCell(2);
            // check Reading of formulas which are shared (two cells from a 1R x 8C range)
            Assert.AreEqual("B32770*2", cellB32769.CellFormula);
            Assert.AreEqual("C32770*2", cellC32769.CellFormula);
            ConfirmCellEvaluation(wb, cellB32769, 4);
            ConfirmCellEvaluation(wb, cellC32769, 6);
            // Confirm this example really does have SharedFormulas.
            // there are 3 others besides the one at A32769:H32769
            Assert.AreEqual(4, countSharedFormulas(sheet));


            // Re-serialize and check again
            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0);
            cellB32769 = sheet.GetRow(32768).GetCell(1);
            cellC32769 = sheet.GetRow(32768).GetCell(2);
            Assert.AreEqual("B32770*2", cellB32769.CellFormula);
            ConfirmCellEvaluation(wb, cellB32769, 4);
            Assert.AreEqual(4, countSharedFormulas(sheet));
        }
        [Test]
        public void TestUnshareFormulaDueToChangeFormula()
        {
            HSSFWorkbook wb;
            ISheet sheet;
            ICell cellB32769;
            ICell cellC32769;

            wb = HSSFTestDataSamples.OpenSampleWorkbook(SHARED_FORMULA_TEST_XLS);
            sheet = wb.GetSheetAt(0);
            cellB32769 = sheet.GetRow(32768).GetCell(1);
            cellC32769 = sheet.GetRow(32768).GetCell(2);

            // Updating cell formula, causing it to become unshared
            cellB32769.CellFormula = (/*setter*/"1+1");
            ConfirmCellEvaluation(wb, cellB32769, 2);
            // currently (Oct 2008) POI handles this by exploding the whole shared formula group
            Assert.AreEqual(3, countSharedFormulas(sheet)); // one less now
            // check that nearby cell of the same group still has the same formula
            Assert.AreEqual("C32770*2", cellC32769.CellFormula);
            ConfirmCellEvaluation(wb, cellC32769, 6);
        }
        [Test]
        public void TestUnshareFormulaDueToDelete()
        {
            HSSFWorkbook wb;
            ISheet sheet;
            ICell cell;
            int ROW_IX = 2;

            // changing shared formula cell to blank
            wb = HSSFTestDataSamples.OpenSampleWorkbook(SHARED_FORMULA_TEST_XLS);
            sheet = wb.GetSheetAt(0);

            Assert.AreEqual("A$1*2", sheet.GetRow(ROW_IX).GetCell(1).CellFormula);
            cell = sheet.GetRow(ROW_IX).GetCell(1);
            cell.SetCellType(CellType.Blank);
            Assert.AreEqual(3, countSharedFormulas(sheet));

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0);
            Assert.AreEqual("A$1*2", sheet.GetRow(ROW_IX + 1).GetCell(1).CellFormula);

            // deleting shared formula cell
            wb = HSSFTestDataSamples.OpenSampleWorkbook(SHARED_FORMULA_TEST_XLS);
            sheet = wb.GetSheetAt(0);

            Assert.AreEqual("A$1*2", sheet.GetRow(ROW_IX).GetCell(1).CellFormula);
            cell = sheet.GetRow(ROW_IX).GetCell(1);
            sheet.GetRow(ROW_IX).RemoveCell(cell);
            Assert.AreEqual(3, countSharedFormulas(sheet));

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0);
            Assert.AreEqual("A$1*2", sheet.GetRow(ROW_IX + 1).GetCell(1).CellFormula);
        }

        private static void ConfirmCellEvaluation(IWorkbook wb, ICell cell, double expectedValue)
        {
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            CellValue cv = fe.Evaluate(cell);
            Assert.AreEqual(CellType.Numeric, cv.CellType);
            Assert.AreEqual(expectedValue, cv.NumberValue, 0.0);
        }

        /**
         * @return the number of {@link SharedFormulaRecord}s encoded for the specified sheet
         */
        private static int countSharedFormulas(ISheet sheet)
        {
            NPOI.HSSF.Record.Record[] records = RecordInspector.GetRecords(sheet, 0);
            int count = 0;
            for (int i = 0; i < records.Length; i++)
            {
                NPOI.HSSF.Record.Record rec = records[i];
                if (rec is SharedFormulaRecord)
                {
                    count++;
                }
            }
            return count;
        }
    }

}