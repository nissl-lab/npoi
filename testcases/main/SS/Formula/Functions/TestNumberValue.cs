using NPOI.HSSF.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCases.SS.Formula.Functions
{
    [TestFixture]
    public class TestNumberValue
    {
        [Test]
        public void TestMicrosoftExample1()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            HSSFWorkbook wb = new HSSFWorkbook();
            var sheet = wb.CreateSheet();
            var row = sheet.CreateRow(0);
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            var cell = row.CreateCell(0);
            SS.Util.Utils.AssertDouble(fe, cell, "NUMBERVALUE(\"2.500,27\",\",\",\".\")", 2500.27, 0.000000000001);
        }

        [Test]
        public void TestMicrosoftExample2()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            HSSFWorkbook wb = new HSSFWorkbook();
            var sheet = wb.CreateSheet();
            var row = sheet.CreateRow(0);
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            var cell = row.CreateCell(0);
            SS.Util.Utils.AssertDouble(fe, cell, "NUMBERVALUE(\"3.5%\")", 0.035, 0.000000000001);
        }
    }
}
