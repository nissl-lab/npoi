using System;
using System.Collections.Generic;
using System.Text;

namespace JavaToCSharp.Rules
{
    public class foreachRule : Rule
    {
        public override bool Execute(string strOrigin, out string strOutput, int iRowNumber)
        {
            string trimmedStr= strOrigin.Trim();
            strOutput = strOrigin;
            if (!trimmedStr.StartsWith("for(") && !trimmedStr.StartsWith("for ("))
            {
                return false;
            }
            if (strOutput.IndexOf(":", 4) < 0)
                return false;
            strOutput = strOutput.Replace("for(", "foreach(");
            strOutput = strOutput.Replace("for (", "foreach (");
            strOutput = strOutput.Replace(" : ", " in ");
            return true;
        }
        public override string RuleName
        {
            get { return "for(Type x : y)"; }
        }
    }
}
