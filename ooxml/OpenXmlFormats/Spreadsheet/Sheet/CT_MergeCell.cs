using System;
using System.Collections.Generic;

using System.Text;
using System.Xml.Serialization;
using System.Xml;
using NPOI.OpenXml4Net.Util;
using System.IO;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_MergeCell
    {
        public static CT_MergeCell Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_MergeCell ctObj = new CT_MergeCell();
            ctObj.@ref = XmlHelper.ReadString(node.Attributes["ref"]);
            return ctObj;
        }



    internal void Write(StreamWriter sw, string nodeName)
    {
        sw.Write(string.Format("<{0}", nodeName));
        XmlHelper.WriteAttribute(sw, "ref", this.@ref);
        sw.Write("/>");
    }


        private string refField;
        [XmlAttribute]
        public string @ref
        {
            get
            {
                return this.refField;
            }
            set
            {
                this.refField = value;
            }
        }
    }
}
