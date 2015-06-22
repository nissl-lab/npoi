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

namespace TestCases.HSSF.Record.CF
{
    using System;
    using NUnit.Framework;
    using NPOI.HSSF.Record.CF;
    using NPOI.HSSF.Util;
    using NPOI.SS.Util;
    using NPOI.Util;

    /**
     * Tests CellRange operations.
     */
    [TestFixture]
    public class TestCellRange
    {
        private static CellRangeAddress biggest = CreateCR(0, -1, 0, -1);
        private static CellRangeAddress tenthColumn = CreateCR(0, -1, 10, 10);
        private static CellRangeAddress tenthRow = CreateCR(10, 10, 0, -1);
        private static CellRangeAddress box10x10 = CreateCR(0, 10, 0, 10);
        private static CellRangeAddress box9x9 = CreateCR(0, 9, 0, 9);
        private static CellRangeAddress box10to20c = CreateCR(0, 10, 10, 20);
        private static CellRangeAddress oneCell = CreateCR(10, 10, 10, 10);

        private static CellRangeAddress[] sampleRanges = {
            biggest, tenthColumn, tenthRow, box10x10, box9x9, box10to20c, oneCell,
        };

        /** cross-reference of <c>contains()</c> operations for sampleRanges against itself */
        private static bool[,] containsExpectedResults =
        {
        //               biggest, tenthColumn, tenthRow, box10x10, box9x9, box10to20c, oneCell
        /*biggest    */ {true,       true ,    true ,    true ,    true ,      true ,  true},	
        /*tenthColumn*/ {false,      true ,    false,    false,    false,      false,  true},	
        /*tenthRow   */ {false,      false,    true ,    false,    false,      false,  true},	
        /*box10x10   */ {false,      false,    false,    true ,    true ,      false,  true},	
        /*box9x9     */ {false,      false,    false,    false,    true ,      false, false},	
        /*box10to20c */ {false,      false,    false,    false,    false,      true ,  true},	
        /*oneCell    */ {false,      false,    false,    false,    false,      false,  true},	
        };

        /**
         * @param lastRow pass -1 for max row index 
         * @param lastCol pass -1 for max col index
         */
        private static CellRangeAddress CreateCR(int firstRow, int lastRow, int firstCol, int lastCol)
        {
            // max row & max col limit as per BIFF8
            return new CellRangeAddress(
                    firstRow,
                    lastRow == -1 ? 0xFFFF : lastRow,
                    firstCol,
                    lastCol == -1 ? 0x00FF : lastCol);
        }

        [Test]
        public void TestContainsMethod()
        {
            CellRangeAddress[] ranges = sampleRanges;
            for (int i = 0; i != ranges.Length; i++)
            {
                for (int j = 0; j != ranges.Length; j++)
                {
                    bool expectedResult = containsExpectedResults[i, j];
                    Assert.AreEqual(expectedResult, CellRangeUtil.Contains(ranges[i], ranges[j]), "(" + i + "," + j + "): ");
                }
            }
        }

        private static CellRangeAddress col1 = CreateCR(0, -1, 1, 1);
        private static CellRangeAddress col2 = CreateCR(0, -1, 2, 2);
        private static CellRangeAddress row1 = CreateCR(1, 1, 0, -1);
        private static CellRangeAddress row2 = CreateCR(2, 2, 0, -1);

        private static CellRangeAddress box0 = CreateCR(0, 2, 0, 2);
        private static CellRangeAddress box1 = CreateCR(0, 1, 0, 1);
        private static CellRangeAddress box2 = CreateCR(0, 1, 2, 3);
        private static CellRangeAddress box3 = CreateCR(2, 3, 0, 1);
        private static CellRangeAddress box4 = CreateCR(2, 3, 2, 3);
        private static CellRangeAddress box5 = CreateCR(1, 3, 1, 3);

        [Test]
        public void TestHasSharedBorderMethod()
        {
            Assert.IsFalse(CellRangeUtil.HasExactSharedBorder(col1, col1));
            Assert.IsFalse(CellRangeUtil.HasExactSharedBorder(col2, col2));
            Assert.IsTrue(CellRangeUtil.HasExactSharedBorder(col1, col2));
            Assert.IsTrue(CellRangeUtil.HasExactSharedBorder(col2, col1));

            Assert.IsFalse(CellRangeUtil.HasExactSharedBorder(row1, row1));
            Assert.IsFalse(CellRangeUtil.HasExactSharedBorder(row2, row2));
            Assert.IsTrue(CellRangeUtil.HasExactSharedBorder(row1, row2));
            Assert.IsTrue(CellRangeUtil.HasExactSharedBorder(row2, row1));

            Assert.IsFalse(CellRangeUtil.HasExactSharedBorder(row1, col1));
            Assert.IsFalse(CellRangeUtil.HasExactSharedBorder(row1, col2));
            Assert.IsFalse(CellRangeUtil.HasExactSharedBorder(col1, row1));
            Assert.IsFalse(CellRangeUtil.HasExactSharedBorder(col2, row1));
            Assert.IsFalse(CellRangeUtil.HasExactSharedBorder(row2, col1));
            Assert.IsFalse(CellRangeUtil.HasExactSharedBorder(row2, col2));
            Assert.IsFalse(CellRangeUtil.HasExactSharedBorder(col1, row2));
            Assert.IsFalse(CellRangeUtil.HasExactSharedBorder(col2, row2));
            Assert.IsTrue(CellRangeUtil.HasExactSharedBorder(col2, col1));

            Assert.IsFalse(CellRangeUtil.HasExactSharedBorder(box1, box1));
            Assert.IsTrue(CellRangeUtil.HasExactSharedBorder(box1, box2));
            Assert.IsTrue(CellRangeUtil.HasExactSharedBorder(box1, box3));
            Assert.IsFalse(CellRangeUtil.HasExactSharedBorder(box1, box4));

            Assert.IsTrue(CellRangeUtil.HasExactSharedBorder(box2, box1));
            Assert.IsFalse(CellRangeUtil.HasExactSharedBorder(box2, box2));
            Assert.IsFalse(CellRangeUtil.HasExactSharedBorder(box2, box3));
            Assert.IsTrue(CellRangeUtil.HasExactSharedBorder(box2, box4));

            Assert.IsTrue(CellRangeUtil.HasExactSharedBorder(box3, box1));
            Assert.IsFalse(CellRangeUtil.HasExactSharedBorder(box3, box2));
            Assert.IsFalse(CellRangeUtil.HasExactSharedBorder(box3, box3));
            Assert.IsTrue(CellRangeUtil.HasExactSharedBorder(box3, box4));

            Assert.IsFalse(CellRangeUtil.HasExactSharedBorder(box4, box1));
            Assert.IsTrue(CellRangeUtil.HasExactSharedBorder(box4, box2));
            Assert.IsTrue(CellRangeUtil.HasExactSharedBorder(box4, box3));
            Assert.IsFalse(CellRangeUtil.HasExactSharedBorder(box4, box4));
        }

        [Test]
        public void TestIntersectMethod()
        {
            Assert.AreEqual(CellRangeUtil.OVERLAP, CellRangeUtil.Intersect(box0, box5));
            Assert.AreEqual(CellRangeUtil.OVERLAP, CellRangeUtil.Intersect(box5, box0));
            Assert.AreEqual(CellRangeUtil.NO_INTERSECTION, CellRangeUtil.Intersect(box1, box4));
            Assert.AreEqual(CellRangeUtil.NO_INTERSECTION, CellRangeUtil.Intersect(box4, box1));
            Assert.AreEqual(CellRangeUtil.NO_INTERSECTION, CellRangeUtil.Intersect(box2, box3));
            Assert.AreEqual(CellRangeUtil.NO_INTERSECTION, CellRangeUtil.Intersect(box3, box2));
            Assert.AreEqual(CellRangeUtil.INSIDE, CellRangeUtil.Intersect(box0, box1));
            Assert.AreEqual(CellRangeUtil.INSIDE, CellRangeUtil.Intersect(box0, box0));
            Assert.AreEqual(CellRangeUtil.ENCLOSES, CellRangeUtil.Intersect(box1, box0));
            Assert.AreEqual(CellRangeUtil.INSIDE, CellRangeUtil.Intersect(tenthColumn, oneCell));
            Assert.AreEqual(CellRangeUtil.ENCLOSES, CellRangeUtil.Intersect(oneCell, tenthColumn));
            Assert.AreEqual(CellRangeUtil.OVERLAP, CellRangeUtil.Intersect(tenthColumn, tenthRow));
            Assert.AreEqual(CellRangeUtil.OVERLAP, CellRangeUtil.Intersect(tenthRow, tenthColumn));
            Assert.AreEqual(CellRangeUtil.INSIDE, CellRangeUtil.Intersect(tenthColumn, tenthColumn));
            Assert.AreEqual(CellRangeUtil.INSIDE, CellRangeUtil.Intersect(tenthRow, tenthRow));
        }

        /**
         * Cell ranges like the following are valid
         * =$C:$IV,$B$1:$B$8,$B$10:$B$65536,$A:$A
         */
        [Test]
        public void TestCreate()
        {
            CellRangeAddress cr;

            cr = CreateCR(0, -1, 2, 255); // $C:$IV
            ConfirmRange(cr, false, true);
            cr = CreateCR(0, 7, 1, 1); // $B$1:$B$8

            try
            {
                cr = CreateCR(9, -1, 1, 1); // $B$65536
            }
            catch (ArgumentException e)
            {
                if (e.Message.StartsWith("invalid cell range"))
                {
                    throw new AssertionException("Identified bug 44739");
                }
                throw e;
            }
            cr = CreateCR(0, -1, 0, 0); // $A:$A
        }

        private static void ConfirmRange(CellRangeAddress cr, bool isFullRow, bool isFullColumn)
        {
            Assert.AreEqual(isFullRow, cr.IsFullRowRange, "isFullRowRange");
            Assert.AreEqual(isFullColumn, cr.IsFullColumnRange, "isFullColumnRange");
        }

        [Test]
        public void TestNumberOfCells()
        {
            Assert.AreEqual(1, oneCell.NumberOfCells);
            Assert.AreEqual(100, box9x9.NumberOfCells);
            Assert.AreEqual(121, box10to20c.NumberOfCells);
        }

        [Test]
        public void TestMergeCellRanges()
        {
            // no result on empty
            CellRangeTest(new String[] { });

            // various cases with two ranges
            CellRangeTest(new String[] { "A1:B1", "A2:B2" }, "A1:B2");
            CellRangeTest(new String[] { "A1:B1" }, "A1:B1");
            CellRangeTest(new String[] { "A1:B2", "A2:B2" }, "A1:B2");
            CellRangeTest(new String[] { "A1:B3", "A2:B2" }, "A1:B3");
            CellRangeTest(new String[] { "A1:C1", "A2:B2" }, new String[] { "A1:C1", "A2:B2" });

            // cases with three ranges
            CellRangeTest(new String[] { "A1:A1", "A2:B2", "A1:C1" }, new String[] { "A1:C1", "A2:B2" });
            CellRangeTest(new String[] { "A1:C1", "A2:B2", "A1:A1" }, new String[] { "A1:C1", "A2:B2" });

            // "standard" cases
            // enclose
            CellRangeTest(new String[] { "A1:D4", "B2:C3" }, new String[] { "A1:D4" });
            // inside
            CellRangeTest(new String[] { "B2:C3", "A1:D4" }, new String[] { "A1:D4" });
            CellRangeTest(new String[] { "B2:C3", "A1:D4" }, new String[] { "A1:D4" });
            // disjunct
            CellRangeTest(new String[] { "A1:B2", "C3:D4" }, new String[] { "A1:B2", "C3:D4" });
            CellRangeTest(new String[] { "A1:B2", "A3:D4" }, new String[] { "A1:B2", "A3:D4" });
            // overlap that cannot be merged
            CellRangeTest(new String[] { "C1:D2", "C2:C3" }, new String[] { "C1:D2", "C2:C3" });
            // overlap which could theoretically be merged, but isn't because the implementation was buggy and therefore was removed
            CellRangeTest(new String[] { "A1:C3", "B1:D3" }, new String[] { "A1:C3", "B1:D3" }); // could be one region "A1:D3"
            CellRangeTest(new String[] { "A1:C3", "B1:D1" }, new String[] { "A1:C3", "B1:D1" }); // could be one region "A1:D3"
        }
        [Test]
        public void TestMergeCellRanges55380()
        {
            CellRangeTest(new String[] { "C1:D2", "C2:C3" }, new String[] { "C1:D2", "C2:C3" });
            CellRangeTest(new String[] { "A1:C3", "B2:D2" }, new String[] { "A1:C3", "B2:D2" });
            CellRangeTest(new String[] { "C9:D30", "C7:C31" }, new String[] { "C9:D30", "C7:C31" });
        }
    
    //    public void testResolveRangeOverlap() {
    //        resolveRangeOverlapTest("C1:D2", "C2:C3");
    //    }

        private void CellRangeTest(String[] input, params string[] expectedOutput)
        {
            CellRangeAddress[] inputArr = new CellRangeAddress[input.Length];
            for (int i = 0; i < input.Length; i++)
            {
                inputArr[i] = CellRangeAddress.ValueOf(input[i]);
            }
            CellRangeAddress[] result = CellRangeUtil.MergeCellRanges(inputArr);
            VerifyExpectedResult(result, expectedOutput);
        }

//    private void resolveRangeOverlapTest(String a, String b, String...expectedOutput) {
//        CellRangeAddress rangeA = CellRangeAddress.valueOf(a);
//        CellRangeAddress rangeB = CellRangeAddress.valueOf(b);
//        CellRangeAddress[] result = CellRangeUtil.resolveRangeOverlap(rangeA, rangeB);
//        verifyExpectedResult(result, expectedOutput);
//    }

        private void VerifyExpectedResult(CellRangeAddress[] result, params string[] expectedOutput)
        {
            Assert.AreEqual(expectedOutput.Length, result.Length,
                "\nExpected: " + Arrays.ToString(expectedOutput) + "\nHad: " + Arrays.ToString(result));
            for (int i = 0; i < expectedOutput.Length; i++)
            {
                Assert.AreEqual(expectedOutput[i], result[i].FormatAsString(),
                    "\nExpected: " + Arrays.ToString(expectedOutput) + "\nHad: " + Arrays.ToString(result));
            }
        }

        [Test]
        public void TestValueOf()
        {
            CellRangeAddress cr1 = CellRangeAddress.ValueOf("A1:B1");
            Assert.AreEqual(0, cr1.FirstColumn);
            Assert.AreEqual(0, cr1.FirstRow);
            Assert.AreEqual(1, cr1.LastColumn);
            Assert.AreEqual(0, cr1.LastRow);

            CellRangeAddress cr2 = CellRangeAddress.ValueOf("B1");
            Assert.AreEqual(1, cr2.FirstColumn);
            Assert.AreEqual(0, cr2.FirstRow);
            Assert.AreEqual(1, cr2.LastColumn);
            Assert.AreEqual(0, cr2.LastRow);
        }
    }
}