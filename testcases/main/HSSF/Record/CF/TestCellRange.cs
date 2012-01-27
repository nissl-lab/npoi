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
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.HSSF.Record.CF;
    using NPOI.HSSF.Util;
    using NPOI.SS.Util;

    /**
     * Tests CellRange operations.
     */
    [TestClass]
    public class TestCellRange
    {
	    private static CellRangeAddress biggest     = CreateCR( 0, -1, 0,-1);
	    private static CellRangeAddress tenthColumn = CreateCR( 0, -1,10,10);
	    private static CellRangeAddress tenthRow    = CreateCR(10, 10, 0,-1);
	    private static CellRangeAddress box10x10    = CreateCR( 0, 10, 0,10);
	    private static CellRangeAddress box9x9      = CreateCR( 0,  9, 0, 9);
	    private static CellRangeAddress box10to20c  = CreateCR( 0, 10,10,20);
	    private static CellRangeAddress oneCell     = CreateCR(10, 10,10,10);

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
	    private static CellRangeAddress CreateCR(int firstRow, int lastRow, int firstCol, int lastCol) {
		    // max row & max col limit as per BIFF8
		    return new CellRangeAddress(
				    firstRow, 
				    lastRow == -1 ? 0xFFFF : lastRow, 
				    firstCol,
				    lastCol == -1 ? 0x00FF : lastCol);
	    }

        [TestMethod]
	    public void TestContainsMethod()
	    {
		    CellRangeAddress[] ranges = sampleRanges;
		    for(int i = 0; i != ranges.Length; i++)
		    {
			    for(int j = 0; j != ranges.Length; j++)
			    {
				    bool expectedResult = containsExpectedResults[i,j];
                    Assert.AreEqual<bool>(expectedResult, CellRangeUtil.Contains(ranges[i], ranges[j]), "("+i+","+j+"): ");
			    }
		    }
	    }

	    private static CellRangeAddress col1     = CreateCR( 0, -1, 1,1);
	    private static CellRangeAddress col2     = CreateCR( 0, -1, 2,2);
	    private static CellRangeAddress row1     = CreateCR( 1,  1, 0,-1);
	    private static CellRangeAddress row2     = CreateCR( 2,  2, 0,-1);

	    private static CellRangeAddress box0     = CreateCR( 0, 2, 0,2);
	    private static CellRangeAddress box1     = CreateCR( 0, 1, 0,1);
	    private static CellRangeAddress box2     = CreateCR( 0, 1, 2,3);
	    private static CellRangeAddress box3     = CreateCR( 2, 3, 0,1);
	    private static CellRangeAddress box4     = CreateCR( 2, 3, 2,3);
	    private static CellRangeAddress box5     = CreateCR( 1, 3, 1,3);

        [TestMethod]
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

        [TestMethod]
	    public void TestIntersectMethod()
	    {
		    Assert.AreEqual<int>(CellRangeUtil.OVERLAP, CellRangeUtil.Intersect(box0, box5));
		    Assert.AreEqual<int>(CellRangeUtil.OVERLAP, CellRangeUtil.Intersect(box5, box0));
		    Assert.AreEqual<int>(CellRangeUtil.NO_INTERSECTION, CellRangeUtil.Intersect(box1, box4));
		    Assert.AreEqual<int>(CellRangeUtil.NO_INTERSECTION, CellRangeUtil.Intersect(box4, box1));
		    Assert.AreEqual<int>(CellRangeUtil.NO_INTERSECTION, CellRangeUtil.Intersect(box2, box3));
		    Assert.AreEqual<int>(CellRangeUtil.NO_INTERSECTION, CellRangeUtil.Intersect(box3, box2));
		    Assert.AreEqual<int>(CellRangeUtil.INSIDE, CellRangeUtil.Intersect(box0, box1));
		    Assert.AreEqual<int>(CellRangeUtil.INSIDE, CellRangeUtil.Intersect(box0, box0));
		    Assert.AreEqual<int>(CellRangeUtil.ENCLOSES, CellRangeUtil.Intersect(box1, box0));
		    Assert.AreEqual<int>(CellRangeUtil.INSIDE, CellRangeUtil.Intersect(tenthColumn, oneCell));
		    Assert.AreEqual<int>(CellRangeUtil.ENCLOSES, CellRangeUtil.Intersect(oneCell, tenthColumn));
		    Assert.AreEqual<int>(CellRangeUtil.OVERLAP, CellRangeUtil.Intersect(tenthColumn, tenthRow));
		    Assert.AreEqual<int>(CellRangeUtil.OVERLAP, CellRangeUtil.Intersect(tenthRow, tenthColumn));
		    Assert.AreEqual<int>(CellRangeUtil.INSIDE, CellRangeUtil.Intersect(tenthColumn, tenthColumn));
		    Assert.AreEqual<int>(CellRangeUtil.INSIDE, CellRangeUtil.Intersect(tenthRow, tenthRow));
	    }
    	
	    /**
	     * Cell ranges like the following are valid
	     * =$C:$IV,$B$1:$B$8,$B$10:$B$65536,$A:$A
	     */
        [TestMethod]
	    public void TestCreate() {
		    CellRangeAddress cr;
    		
		    cr = CreateCR(0, -1, 2, 255); // $C:$IV
		    ConfirmRange(cr, false, true);
		    cr = CreateCR(0, 7, 1, 1); // $B$1:$B$8
    		
		    try {
			    cr = CreateCR(9, -1, 1, 1); // $B$65536
		    } catch (ArgumentException e) {
			    if(e.Message.StartsWith("invalid cell range")) {
				    throw new AssertFailedException("Identified bug 44739");
			    }
			    throw e;
		    }
		    cr = CreateCR(0, -1, 0, 0); // $A:$A
	    }

	    private static void ConfirmRange(CellRangeAddress cr, bool isFullRow, bool isFullColumn) {
		    Assert.AreEqual<bool>(isFullRow, cr.IsFullRowRange, "isFullRowRange");
		    Assert.AreEqual<bool>(isFullColumn, cr.IsFullColumnRange, "isFullColumnRange");
	    }
    }
}