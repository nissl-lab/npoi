using NPOI.OpenXml4Net.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ElementName = "comment")]
    public class CT_Comment
    {
        //public CT_Comment()
        //{
        //    this.textField = new CT_Rst();
        //}
        public static CT_Comment Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Comment ctObj = new CT_Comment();
            ctObj.@ref = XmlHelper.ReadString(node.Attributes[nameof(@ref)]);
            if (node.Attributes[nameof(authorId)] != null)
                ctObj.authorId = XmlHelper.ReadUInt(node.Attributes[nameof(authorId)]);
            ctObj.guid = XmlHelper.ReadString(node.Attributes[nameof(guid)]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(text))
                    ctObj.text = CT_Rst.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }
        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");
            XmlHelper.WriteAttribute(sw, nameof(@ref), this.@ref);
            XmlHelper.WriteAttribute(sw, nameof(authorId), this.authorId, true);
            XmlHelper.WriteAttribute(sw, nameof(guid), this.guid);
            sw.Write(">");
            if (this.text != null)
                this.text.Write(sw, nameof(text));
            sw.Write($"</{nodeName}>");
        }
        [XmlElement(nameof(text))]
        public CT_Rst text { get; set; } = new CT_Rst();

        [XmlAttribute(nameof(@ref))]
        public string @ref { get; set; } = string.Empty;

        [XmlAttribute(nameof(authorId))]
        public uint authorId { get; set; } = 0;

        [XmlAttribute(nameof(guid))] // 0..1 TODO: Type is ST_Guid
        public string guid { get; set; } = null;
    }
}
