using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace JavaToCSharp.Rules
{
    public class EquivalentRule : Rule
    {
        public EquivalentRule()
        { 
            
        }
        public EquivalentRule(string name)
        {
            this._name = name;
        }

        string _name;
        public override string RuleName
        {
            get { return _name; }
        }

        public virtual string Replacement
        {
            get;
            set;
        }
        
        public virtual string Pattern
        {
            get;
            set;
        }

        protected virtual string ReplaceString(Match match)
        {
            return this.Replacement;
        }

        public override sealed bool Execute(string strOrigin, out string strOutput, int iRowNumber)
        {
            Regex regex = new Regex(this.Pattern);
            string result=strOrigin;
            bool changedFlag = false;

            if (regex.IsMatch(strOrigin))
            {
                result= regex.Replace(strOrigin, new MatchEvaluator(ReplaceString));
                changedFlag = true;
            }
            strOutput = result;
            return changedFlag;
        }
        public override string ToString()
        {
            return string.Format("Equivalent Rule '{0}'", this.RuleName);
        }
    }
}
