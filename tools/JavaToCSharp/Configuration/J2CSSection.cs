using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;

namespace JavaToCSharp.Configuration
{
    public class J2CSSection : ConfigurationSection
    {
        [ConfigurationProperty("Rules", IsDefaultCollection = true)]
        [ConfigurationCollection(typeof(RuleSection))]
        public RuleSectionCollection Rules
        {
            get { return this["Rules"] as RuleSectionCollection; }
        }
    }
}
