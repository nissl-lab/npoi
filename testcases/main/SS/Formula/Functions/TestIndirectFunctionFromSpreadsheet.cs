using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System;
using NUnit.Framework;
namespace TestCases.SS.Formula.Functions
{
    [TestFixture]
    public class TestIndirectFunctionFromSpreadsheet:BaseTestFunctionsFromSpreadsheet
    {
        protected override string Filename
        {
            get { return "IndirectFunctionTestCaseData.xls"; }
        }
    }
}
