using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;

namespace JavaToCSharp.Configuration
{
    public class RuleSection : ConfigurationElement
    {
        public RuleSection() { }
        public RuleSection(string name, string pattern, string replacement)
        {
            this["name"] = name;
            this["pattern"] = name;
            this["replacement"] = name;
        }

        [ConfigurationProperty("name", IsRequired = true)]
        public string Name {
            get { return (string)this["name"]; }
            set { this["name"] = value; } 
        }
        [ConfigurationProperty("pattern", IsRequired = true)]
        public string Pattern 
        {
            get { return (string)this["pattern"]; }
            set { this["pattern"] = value; }
        }
        [ConfigurationProperty("replacement", IsRequired = true)]
        public string Replacement 
        {
            get { return (string)this["replacement"]; }
            set { this["replacement"] = value; }
        }
        [ConfigurationProperty("type", IsRequired = false)]
        public string Type
        {
            get { return (string)this["type"]; }
            set { this["type"] = value; }
        }
    }
}
