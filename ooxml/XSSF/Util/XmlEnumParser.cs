using NPOI.OpenXmlFormats.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Serialization;

namespace NPOI.XSSF.Util
{
    [Obsolete]
    public class XmlEnumParser<TReturn>
    {
        private static readonly Dictionary<string, TReturn> values;
        static XmlEnumParser()
        {
            Type type = typeof(TReturn);
            MemberInfo[] members = type.GetMembers(BindingFlags.Public | BindingFlags.Static);
            string[] names = Enum.GetNames(type);
            values = new Dictionary<string, TReturn>();
            Array array = type.GetEnumValues();
            foreach (var member in members)
            {
                object[] cas = member.GetCustomAttributes(typeof(XmlEnumAttribute), false);
                if (cas.Length > 0)
                {
                    XmlEnumAttribute attribute = (XmlEnumAttribute)cas[0];
                    values.Add(attribute.Name, (TReturn)Enum.Parse(type, member.Name));
                }
            }
        }

        public static TReturn ForName(string name, TReturn defaultValue)
        {
            if (values.TryGetValue(name, out TReturn forName))
                return forName;
            return defaultValue;
        }
    }
}
