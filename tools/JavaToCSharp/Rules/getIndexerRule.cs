using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace JavaToCSharp.Rules
{
    public class getIndexerRule : Rule
    {
        string ReplaceString(Match matchString)
        {
            string strTemp = matchString.Groups[1].Value;
            return string.Format("[{0}]",strTemp);
        }

        public override bool Execute(string strOrigin, out string strOutput, int iRowNumber)
        {
            strOutput = strOrigin;
            Regex regex=new Regex(@"\.get\((\s+\w+\s+)\)",RegexOptions.IgnoreCase);

            bool changedFlag = false;
            if (regex.IsMatch(strOrigin))
            {
                strOutput = regex.Replace(strOrigin, new MatchEvaluator(ReplaceString));
                changedFlag = true;
            }
            return changedFlag;
        }
        public override string RuleName
        {
            get { return ".get(i)"; }
        }
    }
}
