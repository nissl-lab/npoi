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
namespace TestCases.SS.Util.CellWalk
{

    using NPOI.SS.UserModel;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Util;
    using System;
    using NUnit.Framework;
    using NPOI.SS.Util.CellWalk;
    [TestFixture]
    public class TestCellWalk
    {

        private static Object[][] TestData = new Object[][] {
	new object[] {   1,          2,  null},
	new object[] {null, new DateTime(),  null},
	new object[] {null,       null, "str"}
    };

        private CountCellHandler countCellHandler = new CountCellHandler();
        [Test]
        public void TestNotTraverseEmptyCells()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = new SheetBuilder(wb, TestData).Build();
            CellRangeAddress range = CellRangeAddress.ValueOf("A1:C3");

            CellWalk cellWalk = new CellWalk(sheet, range);
            countCellHandler.reset();
            cellWalk.Traverse(countCellHandler);

            Assert.AreEqual(4, countCellHandler.GetVisitedCellsNumber());
            /* 1 + 2 + 5 + 9 */
            Assert.AreEqual(17L, countCellHandler.GetOrdinalNumberSum());
        }


        private class CountCellHandler : ICellHandler
        {

            private int cellsVisited = 0;
            private long ordinalNumberSum = 0L;

            public void OnCell(ICell cell, ICellWalkContext ctx)
            {
                ++cellsVisited;
                ordinalNumberSum += ctx.OrdinalNumber;
            }

            public int GetVisitedCellsNumber()
            {
                return cellsVisited;
            }

            public long GetOrdinalNumberSum()
            {
                return ordinalNumberSum;
            }

            public void reset()
            {
                cellsVisited = 0;
                ordinalNumberSum = 0L;
            }
        }
    }
}