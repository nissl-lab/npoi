using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestCases.SS.Formula.Functions
{
    public class TestMatchFunctionsFromSpreadsheet:BaseTestFunctionsFromSpreadsheet
    {
        protected override string Filename
        {
            get { return "MatchFunctionTestCaseData.xls"; }
        }
    }
}
