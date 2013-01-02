using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace JavaToCSharp.Rules
{
    public class DeleteRule : Rule
    {
        public override string RuleName
        {
            get { return "Delete Keyword"; }
        }

        const string pattern = @"\sfinal\s|\sthrows\s[a-zA-Z]+Exception|\.doubleValue\(\)";

        public override bool Execute(string strOrigin, out string strOutput, int iRowNumber)
        {
            Regex regex = new Regex(pattern);
            string result = strOrigin;
            bool changedFlag = false;
            if (regex.IsMatch(strOrigin))
            {
                result = regex.Replace(strOrigin," ");
                changedFlag = true;
            }
            strOutput = result;
            return changedFlag;
        }
    }
}
