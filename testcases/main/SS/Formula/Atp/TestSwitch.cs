using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCases.SS.Formula.Atp
{
    /// <summary>
    /// Testcase for 'Analysis Toolpak' function SWITCH() 
    /// </summary>
    ///<remarks>@author Pieter Degraeuwe</remarks>
    [TestFixture]
    public class TestSwitch
    {
        [Test]
        public void TestEvaluate()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sh = wb.CreateSheet();
            IRow row1 = sh.CreateRow(0);

            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            // Create cells
            row1.CreateCell(0, CellType.String);

            // Create references
            CellReference a1Ref = new CellReference("A1");

            // Set values
            ICell cellA1 = sh.GetRow(a1Ref.Row).GetCell(a1Ref.Col);


            ICell cell1 = row1.CreateCell(1);
            cell1.SetCellFormula("SWITCH(A1, \"A\",\"Value for A\", \"B\",\"Value for B\", \"Something else\")");


            cellA1.SetCellValue("A");
            Assert.AreEqual(CellType.String, evaluator.Evaluate(cell1).CellType);
            Assert.AreEqual("Value for A", evaluator.Evaluate(cell1).StringValue,
                    "SWITCH should return 'Value for A'");

            cellA1.SetCellValue("B");
            evaluator.ClearAllCachedResultValues();
            Assert.AreEqual(CellType.String, evaluator.Evaluate(cell1).CellType);
            Assert.AreEqual("Value for B", evaluator.Evaluate(cell1).StringValue,
                    "SWITCH should return 'Value for B'");

            cellA1.SetCellValue("");
            evaluator.ClearAllCachedResultValues();
            Assert.AreEqual(CellType.String, evaluator.Evaluate(cell1).CellType);
            Assert.AreEqual("Something else", evaluator.Evaluate(cell1).StringValue,
                    "SWITCH should return 'Something else'");

        }
    }
}
