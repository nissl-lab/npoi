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

            FormulaShifter fs = FormulaShifter.CreateForRowShift(0,"", firstRowMoved, lastRowMoved, numberRowsMoved);
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
        private static AreaPtg CreateAreaPtg(int InitialAreaFirstRow, int InitialAreaLastRow)
        {
            return new AreaPtg(InitialAreaFirstRow, InitialAreaLastRow, 2, 5, false, false, false, false);
        }
    }

}