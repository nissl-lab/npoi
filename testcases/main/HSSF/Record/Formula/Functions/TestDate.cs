/*
 * Created on Sep 11, 2007
 * 
 * The Copyright statements and Licenses for the commons application may be
 * found in the file LICENSE.txt
 */

namespace TestCases.HSSF.Record.Formula.Functions
{
    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.HSSF.UserModel;

    using NPOI.HSSF.Record.Formula.Functions;


    /**
     * @author Pavel Krupets (pkrupets at palmtreebusiness dot com)
     */
    [TestClass]
    public class TestDate
    {
        private NPOI.SS.UserModel.ICell cell11, cell12, cell13, cell14, cell15,cell16;
        private HSSFFormulaEvaluator evaluator;

        [TestInitialize]
        public void SetUp()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet("new sheet");
            NPOI.SS.UserModel.IRow row1 = sheet.CreateRow(0);

            this.cell11 = row1.CreateCell(0);
            this.cell12 = row1.CreateCell(1);
            this.cell13 = row1.CreateCell(2);
            this.cell14 = row1.CreateCell(3);
            this.cell15 = row1.CreateCell(4);
            this.cell16 = row1.CreateCell(5);

            this.evaluator = new HSSFFormulaEvaluator(sheet, wb);
            //this.evaluator.SetCurrentRow(row1);
        }

        /**
         * Test disabled pending a fix in the formula parser
         */
        
        public void DISABLEDtestSomeArgumentsMissing()
        {
            this.cell11.CellFormula=("DATE(, 1, 0)");
            Assert.AreEqual(0.0, this.evaluator.Evaluate(this.cell11).NumberValue, 0);

            this.cell12.CellFormula =("DATE(, 1, 1)");
            Assert.AreEqual(1.0, this.evaluator.Evaluate(this.cell12).NumberValue, 0);
        }
        [TestMethod]
        public void TestValid()
        {
            this.cell11.SetCellType(NPOI.SS.UserModel.CellType.FORMULA);

            this.cell11.CellFormula =("DATE(1900, 1, 1)");
            Assert.AreEqual(1, Convert.ToInt32(this.evaluator.Evaluate(this.cell11).NumberValue));

            this.cell12.CellFormula =("DATE(1900, 1, 32)");
            Assert.AreEqual(32, Convert.ToInt32(this.evaluator.Evaluate(this.cell12).NumberValue));

            this.cell13.CellFormula =("DATE(1900, 222, 1)");
            Assert.AreEqual(6727, Convert.ToInt32(this.evaluator.Evaluate(this.cell13).NumberValue));

            this.cell14.CellFormula =("DATE(1900, 2, 0)");
            Assert.AreEqual(31, Convert.ToInt32(this.evaluator.Evaluate(this.cell14).NumberValue));

            this.cell15.CellFormula =("DATE(2000, 1, 222)");
            Assert.AreEqual(36747.00, this.evaluator.Evaluate(this.cell15).NumberValue, 0);

            this.cell16.CellFormula =("DATE(2007, 1, 1)");
            Assert.AreEqual(39083, Convert.ToInt32(this.evaluator.Evaluate(this.cell16).NumberValue));
        }
        [TestMethod]
        public void TestBugDate()
        {
            this.cell11.CellFormula =("DATE(1900, 2, 29)");
            Assert.AreEqual(60, Convert.ToInt32(this.evaluator.Evaluate(this.cell11).NumberValue));

            this.cell12.CellFormula =("DATE(1900, 2, 30)");
            Assert.AreEqual(61, Convert.ToInt32(this.evaluator.Evaluate(this.cell12).NumberValue));

            this.cell13.CellFormula =("DATE(1900, 1, 222)");
            Assert.AreEqual(222, Convert.ToInt32(this.evaluator.Evaluate(this.cell13).NumberValue));

            this.cell14.CellFormula =("DATE(1900, 1, 2222)");
            Assert.AreEqual(2222, Convert.ToInt32(this.evaluator.Evaluate(this.cell14).NumberValue));

            this.cell15.CellFormula =("DATE(1900, 1, 22222)");
            Assert.AreEqual(22222, Convert.ToInt32(this.evaluator.Evaluate(this.cell15).NumberValue));
        }
        [TestMethod]
        public void TestPartYears()
        {
            this.cell11.CellFormula =("DATE(4, 1, 1)");
            Assert.AreEqual(1462.00, this.evaluator.Evaluate(this.cell11).NumberValue);

            this.cell12.CellFormula =("DATE(14, 1, 1)");
            Assert.AreEqual(5115.00, this.evaluator.Evaluate(this.cell12).NumberValue);

            this.cell13.CellFormula =("DATE(104, 1, 1)");
            Assert.AreEqual(37987.00, this.evaluator.Evaluate(this.cell13).NumberValue);

            this.cell14.CellFormula =("DATE(1004, 1, 1)");
            Assert.AreEqual(366705.00, this.evaluator.Evaluate(this.cell14).NumberValue);
        }


    }
}
