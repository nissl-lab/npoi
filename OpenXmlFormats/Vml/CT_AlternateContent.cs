using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;

namespace NPOI.OpenXmlFormats.Vml
{
    public class CT_AlternateContent
    {
        public string InnerXml { get; set; }
        public CT_AlternateContent()
        {
        }
        public static CT_AlternateContent Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;

            var ac = new CT_AlternateContent();
            if (string.IsNullOrEmpty(node.InnerXml))
            {
                return ac;
            }
            ac.InnerXml = node.InnerXml;
            return ac;
        }
        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<mc:{0} xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"", nodeName));
            if (this.InnerXml == null)
            {
                sw.Write(string.Format("/>", nodeName));
            }
            else
            {
                sw.Write(">");
                sw.Write(this.InnerXml);
                sw.Write(string.Format("</mc:{0}>", nodeName));
            }
        }
    }
}
