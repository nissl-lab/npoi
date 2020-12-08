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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;

namespace TestCases.SS.Formula.Atp
{
    [TestFixture]
    public class TestIfs
    {
        /// <summary>
        /// =IFS(A1="A", "Value for A" , A1="B", "Value for B")
        /// </summary>
        [Test]
        public void TestEvaluate()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sh = wb.CreateSheet();
            IRow row1 = sh.CreateRow(0);

	        // Create cells
	        row1.CreateCell(0, CellType.String);

	        // Create references
	        CellReference a1Ref = new CellReference("A1");

	        // Set values
	        ICell cellA1 = sh.GetRow(a1Ref.Row).GetCell(a1Ref.Col);

	        ICell cell1 = row1.CreateCell(1);
	        cell1.CellFormula = "IFS(A1=\"A\", \"Value for A\", A1=\"B\",\"Value for B\")";

	        IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

	        cellA1.SetCellValue("A");
	        Assert.AreEqual(CellType.String, evaluator.Evaluate(cell1).CellType, "Checks that the cell is numeric");
	        Assert.AreEqual("Value for A", evaluator.Evaluate(cell1).StringValue, "IFS should return 'Value for B'");

	        cellA1.SetCellValue("B");
	        evaluator.ClearAllCachedResultValues();

            Assert.AreEqual(CellType.String, evaluator.Evaluate(cell1).CellType, "Checks that the cell is numeric");
	        Assert.AreEqual("Value for B", evaluator.Evaluate(cell1).StringValue, "IFS should return 'Value for B'");
        }
    }
}
