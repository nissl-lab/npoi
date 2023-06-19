using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCases.SS.Formula.Functions
{
    [TestFixture]
    public class TestNumericFunction
    {
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestINT()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ICell cell = wb.CreateSheet().CreateRow(0).CreateCell(0);
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            SS.Util.Utils.AssertDouble(fe, cell, "INT(880000000.0001)", 880000000.0);
            //the following INT(-880000000.0001) resulting in -880000001.0 has been observed in excel
            //see also https://support.microsoft.com/en-us/office/int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef
            SS.Util.Utils.AssertDouble(fe, cell, "INT(-880000000.0001)", -880000001.0);
            SS.Util.Utils.AssertDouble(fe, cell, "880000000*0.00849", 7471200.0);
            SS.Util.Utils.AssertDouble(fe, cell, "880000000*0.00849/3", 2490400.0);
            SS.Util.Utils.AssertDouble(fe, cell, "INT(880000000*0.00849/3)", 2490400.0);
        }

        [Test]
        public void TestSIGN()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ICell cell = wb.CreateSheet().CreateRow(0).CreateCell(0);
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            //https://support.microsoft.com/en-us/office/sign-function-109c932d-fcdc-4023-91f1-2dd0e916a1d8
            SS.Util.Utils.AssertDouble(fe, cell, "SIGN(10)", 1.0);
            SS.Util.Utils.AssertDouble(fe, cell, "SIGN(4-4)", 0.0);
            SS.Util.Utils.AssertDouble(fe, cell, "SIGN(-0.00001)", -1.0);
        }

        [Test]
        public void TestDOLLAR()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ICell cell = wb.CreateSheet().CreateRow(0).CreateCell(0);
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            //https://support.microsoft.com/en-us/office/dollar-function-a6cd05d9-9740-4ad3-a469-8109d18ff611
            SS.Util.Utils.AssertString(fe, cell, "DOLLAR(1234.567,2)", "$1,234.57");
            SS.Util.Utils.AssertString(fe, cell, "DOLLAR(-1234.567,0)", "($1,235)");
            SS.Util.Utils.AssertString(fe, cell, "DOLLAR(-1234.567,-2)", "($1,200)");
            SS.Util.Utils.AssertString(fe, cell, "DOLLAR(-0.123,4)", "($0.1230)");
            SS.Util.Utils.AssertString(fe, cell, "DOLLAR(99.888)", "$99.89");
            SS.Util.Utils.AssertString(fe, cell, "DOLLAR(123456789.567,2)", "$123,456,789.57");
        }
    }
}
