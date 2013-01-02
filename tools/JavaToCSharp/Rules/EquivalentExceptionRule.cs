using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace JavaToCSharp.Rules
{
    public abstract class EquivalentExceptionRule : EquivalentRule
    {
        protected override sealed string Replacement
        {
            get { return this.Pattern; }
        }
    }


    public class EquivalentRuleException1 : EquivalentExceptionRule
    {

        protected override string Pattern
        {
            get { return @"RuntimeException"; }
        }

        public override string ReplaceString(Match match)
        {
            return "Exception";
        }
    }

    public class EquivalentRuleException2 : EquivalentExceptionRule
    {

        protected override string Pattern
        {
            get { return @"ClassCastException"; }
        }

        public override string ReplaceString(Match match)
        {
            return "InvalidCastException";
        }
    }

    public class EquivalentRuleException3 : EquivalentExceptionRule
    {

        protected override string Pattern
        {
            get { return @"UnsupportedEncodingException"; }
        }

        public override string ReplaceString(Match match)
        {
            return "EncoderFallbackException";
        }
    }

    public class EquivalentRuleException4 : EquivalentExceptionRule
    {

        protected override string Pattern
        {
            get { return @"AssertionFailedError"; }
        }

        public override string ReplaceString(Match match)
        {
            return "AssertFailedException";
        }
    }

    public class EquivalentRuleException5 : EquivalentExceptionRule
    {

        protected override string Pattern
        {
            get { return @"IllegalArgumentException"; }
        }

        public override string ReplaceString(Match match)
        {
            return "ArgumentException";
        }
    }
    public class EquivalentRuleException6 : EquivalentExceptionRule
    {

        protected override string Pattern
        {
            get { return @"NumberFormatException"; }
        }

        public override string ReplaceString(Match match)
        {
            return "FormatException";
        }
    }
    public class EquivalentRuleException7 : EquivalentExceptionRule
    {

        protected override string Pattern
        {
            get { return @"IllegalAccessError|IllegalStateException"; }
        }

        public override string ReplaceString(Match match)
        {
            return "InvalidOperationException";
        }
    }
    
}
