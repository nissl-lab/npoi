using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace JavaToCSharp.Rules
{
    public class skiplineRule:Rule
    {
        public override bool Execute(string strOrigin, out string strOutput, int iRowNumber)
        {
            Regex regex = new Regex(@"(import|using)\s(java\.)", RegexOptions.IgnoreCase);

            if (regex.IsMatch(strOrigin))
            {
                strOutput = string.Empty;
                return true;
            }
            strOutput = strOrigin;
            return false;
        }
        public override string RuleName
        {
            get { return "Skip this line"; }
        }
    }
}
