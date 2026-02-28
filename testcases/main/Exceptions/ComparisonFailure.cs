using System;
using System.Collections.Generic;
using System.Text;

namespace TestCases.Exceptions
{
    //simulate java class ComparisonFailure
    public class ComparisonFailure: ApplicationException
    {
        public ComparisonFailure(string message, string expected, string actual):
            base(string.Format("{0} excepted:{1} actual:{2}", message, expected, actual))
        {

        }
    }
}
