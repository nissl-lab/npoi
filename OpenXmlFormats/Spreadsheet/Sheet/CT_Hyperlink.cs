using System;

using System.Xml.Serialization;
using System.Xml;
using NPOI.OpenXml4Net.Util;
using System.IO;
using NPOI.OpenXml4Net.OPC;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Hyperlink
    {
        [XmlAttribute(nameof(@ref))]
        public string @ref { get; set; } = null;

        [XmlAttribute(Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id { get; set; } = null;

        [XmlAttribute]
        public string location { get; set; } = null;

        [XmlAttribute]
        public string tooltip { get; set; } = null;

        [XmlAttribute]
        public string display { get; set; } = null;

        public static CT_Hyperlink Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Hyperlink ctObj = new CT_Hyperlink();
            ctObj.@ref = XmlHelper.ReadString(node.Attributes[nameof(@ref)]);
            ctObj.id = XmlHelper.ReadString(node.Attributes[nameof(id), PackageNamespaces.SCHEMA_RELATIONSHIPS]);
            ctObj.location = XmlHelper.ReadString(node.Attributes[nameof(location)]);
            ctObj.tooltip = XmlHelper.ReadString(node.Attributes[nameof(tooltip)]);
            ctObj.display = XmlHelper.ReadString(node.Attributes[nameof(display)]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");
            XmlHelper.WriteAttribute(sw, nameof(@ref), this.@ref);
            XmlHelper.WriteAttribute(sw, $"r:{nameof(id)}", this.id);
            XmlHelper.WriteAttribute(sw, nameof(location), this.location);
            XmlHelper.WriteAttribute(sw, nameof(tooltip), this.tooltip);
            XmlHelper.WriteAttribute(sw, nameof(display), this.display);
            sw.Write(">");
            sw.Write($"</{nodeName}>");
        }







    }
}
