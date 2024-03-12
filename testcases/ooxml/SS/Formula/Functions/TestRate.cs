using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NUnit.Framework;
using System.Globalization;
using TestCases.SS.Util;

namespace TestCases.SS.Formula.Functions
{
    /// <summary>
    /// Testing Rate
    /// </summary>
    [TestFixture]
    public class TestRate
    {
        private IWorkbook _workbook;
        private IFormulaEvaluator _formulaEvaluator;
        private ISheet _sheet;

        [OneTimeSetUp]
        public void Setup()
        {
            _workbook = new HSSFWorkbook();
            _sheet = _workbook.CreateSheet();
            _formulaEvaluator = new HSSFFormulaEvaluator(_workbook);
        }

        [OneTimeTearDown]
        public void Dispose()
        {
            _workbook?.Dispose();
            _sheet = null;
            _formulaEvaluator = null;
        }

        [Test]
        public void TestMicrosoftExample1()
        {
            Utils.AddRow(_sheet, 0, "Data", "Description");
            Utils.AddRow(_sheet, 1, 4, "Years of loan");
            Utils.AddRow(_sheet, 2, -200, "Monthly payment");
            Utils.AddRow(_sheet, 3, 8000, "Amount of the loan");
            ICell cell = _sheet.GetRow(0).CreateCell(100);
            
            Utils.AssertDouble(_formulaEvaluator, cell, "RATE(A2*12, A3, A4)", 0.007701472, 0.000001);
            Utils.AssertDouble(_formulaEvaluator, cell, "RATE(A2*12, A3, A4)*12", 0.09241767, 0.000001);
        }

        [Test]
        public void TestLibreOfficeExample1()
        {
            IRow row = _sheet.CreateRow(4);
            ICell cell = row.CreateCell(0);
            Utils.AssertDouble(_formulaEvaluator, cell, "RATE(3,-10,900,1,0,0.5)", -0.7634, 0.0001);
        }

        [Test]
        public void TestLibreOfficeExample2()
        {
            IRow row = _sheet.CreateRow(5);
            ICell cell = row.CreateCell(0);
            Utils.AssertDouble(_formulaEvaluator, cell, "RATE(3,-10,900)", -0.7563, 0.0001);
        }

        // https://wiki.documentfoundation.org/Documentation/Calc_Functions/RATE
        [Ignore("LibreOffice and Excel return #!NUM, but POI and NumPy return a result")]
        [Test]
        public void TestLibreOfficeInfeasibleSolution()
        {
            IRow row = _sheet.CreateRow(6);
            ICell cell = row.CreateCell(0);
            Utils.AssertDouble(_formulaEvaluator, cell, "RATE(3,10,900,1,0,0.5)", -0.7634, 0.0001);
        }


        /**
         * See https://github.com/numpy/numpy-financial/blob/d02edfb65dcdf23bd571c2cded7fcd4a0528c6af/numpy_financial/tests/test_financial.py#L126
        */
        [Test]
        public void TestNumPyExample1()
        {
            double[] expected = {-0.39920185, -0.02305873, -0.41818459, 0.26513414};
            double nper = 2;
            double pmt = 0;
            double[] pv = {-593.06, -4725.38, -662.05, -428.78};
            double[] fv = {214.07, 4509.97, 224.11, 686.29};

            IRow row = _sheet.CreateRow(7);
            CultureInfo culture = CultureInfo.InvariantCulture;

            for(int i = 0; i < pv.Length; i++)
            {
                string fmla = string.Format(
                    "RATE({0}, {1}, {2}, {3}, 0, 0.1)", 
                    nper.ToString("0.00", culture), 
                    pmt.ToString("0.00", culture), 
                    pv[i].ToString("0.00", culture), 
                    fv[i].ToString("0.00", culture));

                ICell cell = row.CreateCell(i);
                Utils.AssertDouble(_formulaEvaluator, cell, fmla, expected[i], 1e-8);
            }
        }

        [Test]
        public void TestNumPyExample2() 
        {
            IRow row = _sheet.CreateRow(8);
            ICell cell = row.CreateCell(0);
            Utils.AssertDouble(_formulaEvaluator, cell, "RATE(10, 0, -3500, 10000.0, 0, 0.1)", 0.1106908537142689284704528100, 1E-6);
        }

        /**
         * RATE will return NaN, if the Newton Raphson method cannot find a
         * feasible rate within the required tolerance or number of iterations.
         * This can occur if both `pmt` and `pv` have the same sign, as it is
         * impossible to repay a loan by making further withdrawls.
         *
         * See https://github.com/numpy/numpy-financial/blob/d02edfb65dcdf23bd571c2cded7fcd4a0528c6af/numpy_financial/tests/test_financial.py#L113
         */
        [Test]
        public void TestNumPyInfeasibleSolution1()
        {
            IRow row = _sheet.CreateRow(9);
            ICell cell = row.CreateCell(0);
            Utils.AssertError(_formulaEvaluator, cell, "RATE(12, 400, 10000, 5000.0, 0, 0.1)", FormulaError.NUM);
        }

        // https://github.com/numpy/numpy-financial/blob/d02edfb65dcdf23bd571c2cded7fcd4a0528c6af/numpy_financial/tests/test_financial.py#L126
        [Test]
        public void TestNumPyInfeasibleSolution2()
        {
            IRow row = _sheet.CreateRow(10);
            ICell cell = row.CreateCell(0);
            Utils.AssertError(_formulaEvaluator, cell, "RATE(2, 0, -13.65, -329.67, 0, 0.1)", FormulaError.NUM);
        }

        [Test]
        public void TestBug65988()
        {
            IRow row = _sheet.CreateRow(11);
            ICell cell = row.CreateCell(0);
            Utils.AssertDouble(_formulaEvaluator, cell, "RATE(360.0,6.56,-2000.0)", 0.0009480170844060, 0.000001);
        }
    }
}
