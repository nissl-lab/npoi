using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;


namespace JavaToCSharp.Rules
{
    public class CapitalizeRule : Rule
    {
        public override string RuleName
        {
            get { return "Capitalize Rule"; }
        }

        const string pattern = @"(remove|add|apply|process|make|rename|change|run|equal|confirm|write|compare|separate|notify|contain|clone|ungroup|trim|isNaN|substring|classify|init|show|shift|replace|check|single|create|eval|append|convert|clear|fill|serialize|find|indexOf|lastIndexOf|resolve|parse|toString|coerce|lookup|valueOf|hook|unhook)\S+";

        string CapitalizeString(Match matchString)
        {
            string strTemp = matchString.ToString();
            strTemp = char.ToUpper(strTemp[0]) + strTemp.Substring(1, strTemp.Length - 1);
            return strTemp;
        }
 

        public override bool Execute(string strOrigin, out string strOutput, int iRowNumber)
        {
            Regex regex = new Regex(pattern);
            string result = strOrigin;
            bool changedFlag = false;
            if (regex.IsMatch(strOrigin))
            {
                result = regex.Replace(strOrigin, new MatchEvaluator(CapitalizeString));
                changedFlag = true;
            }
            strOutput = result;
            return changedFlag;
        }
    }
}
