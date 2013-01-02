using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace JavaToCSharp.Rules
{
    class NPOINamespaceRule:EquivalentRule
    {
        public override string RuleName
        {
            get
            {
                return "namespace";
            }
        }

        public override string Pattern
        {
            get { return @"org\.apache\.poi\.([a-zA-Z]+)\."; }
        }

        public override string Replacement
        {
            get { return "namespace"; }
        }
        protected override string ReplaceString(System.Text.RegularExpressions.Match match)
        {
            return "NPOI."+match.Groups[1].Value.ToUpper()+".";
        }
    }

    public class EquivalentRule1 : EquivalentRule
    {
        public override string RuleName
        {
            get
            {
                return "new Double";
            }
        }

        public override string Replacement
        {
            get { return "new Double"; }
        }

        public override string Pattern
        {
            get { return @"new Double\((.+)\)"; }
        }

        protected override string ReplaceString(Match match)
        {
            return match.Groups[1].Value;
        }
    }
    public class EquivalentRule2 : EquivalentRule
    {
        public override string Replacement
        {
            get { return ".charAt(i)"; }
        }

        public override string Pattern
        {
            get { return @"\.charAt\((\w+)\)"; }
        }

        protected override string ReplaceString(Match match)
        {
            return string.Format("[{0}]", match.Groups[1].Value);
        }
    }


    public class EquivalentRule3 : EquivalentRule
    {
        public override string Replacement
        {
            get { return "read"; }
        }

        public override string Pattern
        {
            get { return @"(\s|\.)(read\S)"; }
        }

        protected override string ReplaceString(Match match)
        {
            string strTemp = match.Groups[2].Value;
            return match.Groups[1].Value + char.ToUpper(strTemp[0]) + strTemp.Substring(1, strTemp.Length - 1);
        }


    }
}
