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

namespace TestCases.SS.Formula
{

    using NPOI.SS.Formula.PTG;
    using NUnit.Framework;
    using NPOI.SS.Formula;
    using NPOI.SS;
    using NPOI.SS.Util;
    using System;


    /**
     * Tests for {@link FormulaShifter}.
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestFormulaShifter
    {
        // Note - the expected result row coordinates here were determined/verified
        // in Excel 2007 by manually Testing.

        /**
         * Tests what happens to area refs when a range of rows from inside, or overlapping are
         * Moved
         */
        [Test]
        public void TestShiftAreasSourceRows()
        {

            // all these operations are on an area ref spanning rows 10 to 20
            AreaPtg aptg = CreateAreaPtg(10, 20);

            ConfirmAreaShift(aptg, 9, 21, 20, 30, 40);
            ConfirmAreaShift(aptg, 10, 21, 20, 30, 40);
            ConfirmAreaShift(aptg, 9, 20, 20, 30, 40);

            ConfirmAreaShift(aptg, 8, 11, -3, 7, 20); // simple expansion of top
            // rows Containing area top being Shifted down:
            ConfirmAreaShift(aptg, 8, 11, 3, 13, 20);
            ConfirmAreaShift(aptg, 8, 11, 7, 17, 20);
            ConfirmAreaShift(aptg, 8, 11, 8, 18, 20);
            ConfirmAreaShift(aptg, 8, 11, 9, 12, 20); // note behaviour Changes here
            ConfirmAreaShift(aptg, 8, 11, 10, 12, 21);
            ConfirmAreaShift(aptg, 8, 11, 12, 12, 23);
            ConfirmAreaShift(aptg, 8, 11, 13, 10, 20);  // ignored

            // rows from within being Moved:
            ConfirmAreaShift(aptg, 12, 16, 3, 10, 20);  // stay within - no change
            ConfirmAreaShift(aptg, 11, 19, 20, 10, 20);  // Move completely out - no change
            ConfirmAreaShift(aptg, 16, 17, -6, 10, 20);  // Moved exactly to top - no change
            ConfirmAreaShift(aptg, 16, 17, -7, 11, 20);  // tRuncation at top
            ConfirmAreaShift(aptg, 12, 16, 4, 10, 20);  // Moved exactly to bottom - no change
            ConfirmAreaShift(aptg, 12, 16, 6, 10, 17);  // tRuncation at bottom

            // rows Containing area bottom being Shifted up:
            ConfirmAreaShift(aptg, 18, 22, -1, 10, 19); // simple contraction at bottom
            ConfirmAreaShift(aptg, 18, 22, -7, 10, 13); // simple contraction at bottom
            ConfirmAreaShift(aptg, 18, 22, -8, 10, 17); // top calculated differently here
            ConfirmAreaShift(aptg, 18, 22, -9, 9, 17);
            ConfirmAreaShift(aptg, 18, 22, -15, 10, 20); // no change because range would be turned inside out
            ConfirmAreaShift(aptg, 15, 19, -7, 13, 20); // dest tRuncates top (even though src is from inside range)
            ConfirmAreaShift(aptg, 19, 23, -12, 7, 18); // complex: src encloses bottom, dest encloses top

            ConfirmAreaShift(aptg, 18, 22, 5, 10, 25); // simple expansion at bottom
        }
        [Test]
        public void TestCopyAreasSourceRowsRelRel()
        {

            // all these operations are on an area ref spanning rows 10 to 20
            AreaPtg aptg = CreateAreaPtg(10, 20, true, true);

            ConfirmAreaCopy(aptg, 0, 30, 20, 30, 40, true);
            ConfirmAreaCopy(aptg, 15, 25, -15, -1, -1, true); //DeletedRef
        }
        [Test]
        public void TestCopyAreasSourceRowsRelAbs()
        {

            // all these operations are on an area ref spanning rows 10 to 20
            AreaPtg aptg = CreateAreaPtg(10, 20, true, false);

            // Only first row should move
            ConfirmAreaCopy(aptg, 0, 30, 20, 20, 30, true);
            ConfirmAreaCopy(aptg, 15, 25, -15, -1, -1, true); //DeletedRef
        }
        [Test]
        public void TestCopyAreasSourceRowsAbsRel()
        {
            // aptg is part of a formula in a cell that was just copied to another row
            // aptg row references should be updated by the difference in rows that the cell was copied
            // No other references besides the cells that were involved in the copy need to be updated
            // this makes the row copy significantly different from the row shift, where all references
            // in the workbook need to track the row shift

            // all these operations are on an area ref spanning rows 10 to 20
            AreaPtg aptg = CreateAreaPtg(10, 20, false, true);

            // Only last row should move
            ConfirmAreaCopy(aptg, 0, 30, 20, 10, 40, true);
            ConfirmAreaCopy(aptg, 15, 25, -15, 5, 10, true); //sortTopLeftToBottomRight swapped firstRow and lastRow because firstRow is absolute
        }
        [Test]
        public void TestCopyAreasSourceRowsAbsAbs()
        {
            // aptg is part of a formula in a cell that was just copied to another row
            // aptg row references should be updated by the difference in rows that the cell was copied
            // No other references besides the cells that were involved in the copy need to be updated
            // this makes the row copy significantly different from the row shift, where all references
            // in the workbook need to track the row shift

            // all these operations are on an area ref spanning rows 10 to 20
            AreaPtg aptg = CreateAreaPtg(10, 20, false, false);

            //AbsFirstRow AbsLastRow references should't change when copied to a different row
            ConfirmAreaCopy(aptg, 0, 30, 20, 10, 20, false);
            ConfirmAreaCopy(aptg, 15, 25, -15, 10, 20, false);
        }

        /**
         * Tests what happens to an area ref when some outside rows are Moved to overlap
         * that area ref
         */
        [Test]
        public void TestShiftAreasDestRows()
        {
            // all these operations are on an area ref spanning rows 20 to 25
            AreaPtg aptg = CreateAreaPtg(20, 25);

            // no change because no overlap:
            ConfirmAreaShift(aptg, 5, 10, 9, 20, 25);
            ConfirmAreaShift(aptg, 5, 10, 21, 20, 25);

            ConfirmAreaShift(aptg, 11, 14, 10, 20, 25);

            ConfirmAreaShift(aptg, 7, 17, 10, -1, -1); // Converted to DeletedAreaRef
            ConfirmAreaShift(aptg, 5, 15, 7, 23, 25); // tRuncation at top
            ConfirmAreaShift(aptg, 13, 16, 10, 20, 22); // tRuncation at bottom
        }

        private static void ConfirmAreaShift(AreaPtg aptg,
                int firstRowMoved, int lastRowMoved, int numberRowsMoved,
                int expectedAreaFirstRow, int expectedAreaLastRow)
        {

            FormulaShifter fs = FormulaShifter.CreateForRowShift(0, "", firstRowMoved, lastRowMoved, numberRowsMoved, SpreadsheetVersion.EXCEL2007);
            bool expectedChanged = aptg.FirstRow != expectedAreaFirstRow || aptg.LastRow != expectedAreaLastRow;

            AreaPtg copyPtg = (AreaPtg)aptg.Copy(); // clone so we can re-use aptg in calling method
            Ptg[] ptgs = { copyPtg, };
            bool actualChanged = fs.AdjustFormula(ptgs, 0);
            if (expectedAreaFirstRow < 0)
            {
                Assert.AreEqual(typeof(AreaErrPtg), ptgs[0].GetType());
                return;
            }
            Assert.AreEqual(expectedChanged, actualChanged);
            Assert.AreEqual(copyPtg, ptgs[0]);  // expected to change in place (although this is not a strict requirement)
            Assert.AreEqual(expectedAreaFirstRow, copyPtg.FirstRow);
            Assert.AreEqual(expectedAreaLastRow, copyPtg.LastRow);

        }

        private static void ConfirmAreaCopy(AreaPtg aptg,
            int firstRowCopied, int lastRowCopied, int rowOffset,
            int expectedFirstRow, int expectedLastRow, bool expectedChanged)
        {

            AreaPtg copyPtg = (AreaPtg)aptg.Copy(); // clone so we can re-use aptg in calling method
            Ptg[] ptgs = { copyPtg, };
            FormulaShifter fs = FormulaShifter.CreateForRowCopy(0, null, firstRowCopied, lastRowCopied, rowOffset, SpreadsheetVersion.EXCEL2007);
            bool actualChanged = fs.AdjustFormula(ptgs, 0);

            // DeletedAreaRef
            if (expectedFirstRow < 0 || expectedLastRow < 0)
            {
                Assert.AreEqual(typeof(AreaErrPtg), ptgs[0].GetType(),
                    "Reference should have shifted off worksheet, producing #REF! error: " + ptgs[0]);
                return;
            }

            Assert.AreEqual(expectedChanged, actualChanged, "Should this AreaPtg change due to row copy?");
            Assert.AreEqual(copyPtg, ptgs[0], "AreaPtgs should be modified in-place when a row containing the AreaPtg is copied");  // expected to change in place (although this is not a strict requirement)
            Assert.AreEqual(expectedFirstRow, copyPtg.FirstRow, "AreaPtg first row");
            Assert.AreEqual(expectedLastRow, copyPtg.LastRow, "AreaPtg last row");

        }

        private static AreaPtg CreateAreaPtg(int InitialAreaFirstRow, int InitialAreaLastRow)
        {
            return CreateAreaPtg(InitialAreaFirstRow, InitialAreaLastRow, false, false);
        }

        private static AreaPtg CreateAreaPtg(int initialAreaFirstRow, int initialAreaLastRow, bool firstRowRelative, bool lastRowRelative)
        {
            return new AreaPtg(initialAreaFirstRow, initialAreaLastRow, 2, 5, firstRowRelative, lastRowRelative, false, false);
        }

        [Test]
        public void TestShiftSheet()
        {
            // 4 sheets, move a sheet from pos 2 to pos 0, i.e. current 0 becomes 1, current 1 becomes pos 2 
            FormulaShifter shifter = FormulaShifter.CreateForSheetShift(2, 0);

            Ptg[] ptgs = new Ptg[] {
          new Ref3DPtg(new CellReference("first", 0, 0, true, true), 0),
          new Ref3DPtg(new CellReference("second", 0, 0, true, true), 1),
          new Ref3DPtg(new CellReference("third", 0, 0, true, true), 2),
          new Ref3DPtg(new CellReference("fourth", 0, 0, true, true), 3),
        };

            shifter.AdjustFormula(ptgs, -1);

            Assert.AreEqual(1, ((Ref3DPtg)ptgs[0]).ExternSheetIndex,
                "formula previously pointing to sheet 0 should now point to sheet 1");
            Assert.AreEqual(2, ((Ref3DPtg)ptgs[1]).ExternSheetIndex,
                "formula previously pointing to sheet 1 should now point to sheet 2");
            Assert.AreEqual(0, ((Ref3DPtg)ptgs[2]).ExternSheetIndex,
                "formula previously pointing to sheet 2 should now point to sheet 0");
            Assert.AreEqual(3, ((Ref3DPtg)ptgs[3]).ExternSheetIndex,
                "formula previously pointing to sheet 3 should be unchanged");
        }

        [Test]
        public void TestShiftSheet2()
        {
            // 4 sheets, move a sheet from pos 1 to pos 2, i.e. current 2 becomes 1, current 1 becomes pos 2 
            FormulaShifter shifter = FormulaShifter.CreateForSheetShift(1, 2);

            Ptg[] ptgs = new Ptg[] {
          new Ref3DPtg(new CellReference("first", 0, 0, true, true), 0),
          new Ref3DPtg(new CellReference("second", 0, 0, true, true), 1),
          new Ref3DPtg(new CellReference("third", 0, 0, true, true), 2),
          new Ref3DPtg(new CellReference("fourth", 0, 0, true, true), 3),
        };
            shifter.AdjustFormula(ptgs, -1);

            Assert.AreEqual(0, ((Ref3DPtg)ptgs[0]).ExternSheetIndex,
                "formula previously pointing to sheet 0 should be unchanged");
            Assert.AreEqual(2, ((Ref3DPtg)ptgs[1]).ExternSheetIndex,
                "formula previously pointing to sheet 1 should now point to sheet 2");
            Assert.AreEqual(1, ((Ref3DPtg)ptgs[2]).ExternSheetIndex,
                "formula previously pointing to sheet 2 should now point to sheet 1");
            Assert.AreEqual(3, ((Ref3DPtg)ptgs[3]).ExternSheetIndex,
                "formula previously pointing to sheet 3 should be unchanged");
        }

        [Test]
        public void TestInvalidArgument()
        {
            try
            {
                FormulaShifter.CreateForRowShift(1, "name", 1, 2, 0, SpreadsheetVersion.EXCEL97);
                Assert.Fail("Should catch exception here");
            }
            catch (ArgumentException)
            {
                // expected here
            }
            try
            {
                FormulaShifter.CreateForRowShift(1, "name", 2, 1, 2, SpreadsheetVersion.EXCEL97);
                Assert.Fail("Should catch exception here");
            }
            catch (ArgumentException)
            {
                // expected here
            }
        }
        [Test]
        public void TestConstructor()
        {
            Assert.IsNotNull(FormulaShifter.CreateForRowShift(1, "name", 1, 2, 2, SpreadsheetVersion.EXCEL2007));
        }
        [Test]
        public void TestToString()
        {
            FormulaShifter shifter = FormulaShifter.CreateForRowShift(0, "sheet", 123, 456, 789,
                    SpreadsheetVersion.EXCEL2007);
            Assert.IsNotNull(shifter);
            Assert.IsNotNull(shifter.ToString());
            Assert.IsTrue(shifter.ToString().Contains("123"));
            Assert.IsTrue(shifter.ToString().Contains("456"));
            Assert.IsTrue(shifter.ToString().Contains("789"));
        }
    }

}