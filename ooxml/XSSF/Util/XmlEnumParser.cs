using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Serialization;

// REMOVE-REFLECTION: Reflection used is unremovable but the class is obsolete and not used elsewhere.
// AOT should automatically trim this class if the end user is not using it.

namespace NPOI.XSSF.Util
{
    [Obsolete]
    public class XmlEnumParser<
#if NET6_0_OR_GREATER
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicFields)]
#endif
        TReturn> where TReturn : struct, Enum
    {
        private static Dictionary<string, TReturn> values;
        static XmlEnumParser()
        {
            Type type = typeof(TReturn);
            MemberInfo[] members = type.GetFields(BindingFlags.Public | BindingFlags.Static);
            values = new Dictionary<string, TReturn>();
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
            if (values.ContainsKey(name))
                return values[name];
            return defaultValue;
        }
    }
}
