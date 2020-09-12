using NPOI.OpenXml4Net.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    //[Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(ElementName = "fonts", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public class CT_Fonts
    {

        private List<CT_Font> fontField;

        private uint countField;

        private bool countFieldSpecified;
        private uint knownFontsField; 

        public CT_Fonts()
        {
            //this.fontField = new List<CT_Font>();
        }
        public static CT_Fonts Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Fonts ctObj = new CT_Fonts();
            ctObj.count = XmlHelper.ReadUInt(node.Attributes["count"]);
            ctObj.knownFontsField = XmlHelper.ReadUInt(node.Attributes["x14ac:knownFonts"]);
            ctObj.font = new List<CT_Font>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "font")
                    ctObj.font.Add(CT_Font.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "count", this.count);
            XmlHelper.WriteAttribute(sw, "x14ac:knownFonts", this.knownFontsField, false);
            sw.Write(">");
            if (this.font != null)
            {
                foreach (CT_Font x in this.font)
                {
                    x.Write(sw, "font");
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public void SetFontArray(List<CT_Font> array)
        {
             fontField = array;
        }
        [XmlElement]
        public List<CT_Font> font
        {
            get
            {
                return this.fontField;
            }
            set
            {
                this.fontField = value;
            }
        }
        [XmlAttribute]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        [XmlIgnore]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
        //internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Fonts));
        //internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
        //    new XmlQualifiedName("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main") });
        //public override string ToString()
        //{
        //    StringWriter stringWriter = new StringWriter();
        //    serializer.Serialize(stringWriter, this, namespaces);
        //    return stringWriter.ToString();
        //}
    }
}
