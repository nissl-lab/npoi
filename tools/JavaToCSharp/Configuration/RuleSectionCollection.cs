using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;

namespace JavaToCSharp.Configuration
{
    public class RuleSectionCollection : ConfigurationElementCollection
    {
        public void Add(RuleSection element)
        {
            BaseAdd(element);
        }

        public void Remove(string name)
        {
            BaseRemove(name);
        }

        public override ConfigurationElementCollectionType CollectionType
        {
            get 
            {
                return ConfigurationElementCollectionType.AddRemoveClearMap;
            }
        }
        protected override ConfigurationElement CreateNewElement()
        {
            return new RuleSection();
        }
        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((RuleSection)element).Name;
        }
        public new RuleSection this[string name]
        {
            get { return (RuleSection)BaseGet(name); }
        }
    }
}
