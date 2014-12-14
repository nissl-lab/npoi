using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestCases.SS.Formula.Functions
{
    public class TestIndirectFunctionFromSpreadsheet:BaseTestFunctionsFromSpreadsheet
    {
        protected override string Filename
        {
            get { return "IndirectFunctionTestCaseData.xls"; }
        }
    }
}
