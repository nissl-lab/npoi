using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
// add name space
using System.Xml;
using System.IO;
using System.Windows.Forms;

namespace NPOI.StyleViewer
{
    internal class StyleDocumentInitializer
    {
        private const string _ProjectDir = "Xml";
        private const string _StylesXml = "styles.xml";

        internal static bool IsReady = false;
        internal static XmlDocument m_StyleXmlDocument = null;

        internal static List<WordStyle> StyleSet = new List<WordStyle>();

        private StyleDocumentInitializer()
        {
            Load();
        }

        internal static bool Load()
        {
            try
            {
                m_StyleXmlDocument = new XmlDocument();

                m_StyleXmlDocument.Load(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                                     _ProjectDir,
                                                     _StylesXml));
                LoadStyleSet();

                return IsReady = true;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private static void LoadStyleSet()
        {
            XmlNamespaceManager m_XmlNamespaceManager = new XmlNamespaceManager(m_StyleXmlDocument.NameTable);
            m_XmlNamespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            XmlElement m_XmlElement = m_StyleXmlDocument.DocumentElement;
            XmlNodeList StyleNodeList = m_XmlElement.SelectNodes("w:style", m_XmlNamespaceManager);


            foreach (XmlNode styleNode in StyleNodeList)
            {
                string TypeName = styleNode.Attributes["w:type"].InnerText.Trim();
                string StyleID = styleNode.Attributes["w:styleId"].InnerText.Trim();
                string StyleXml = styleNode.OuterXml;

                WordStyle ExistWordStyle = StyleSet.FirstOrDefault(p => p.TypeName == TypeName);

                if (ExistWordStyle == null)
                {
                    WordStyle WordStyle = new WordStyle();

                    WordStyle.TypeName = TypeName;
                    WordStyle.Styles.Add(StyleID, StyleXml);
                    StyleSet.Add(WordStyle);
                }
                else
                {
                    ExistWordStyle.Styles.Add(StyleID, StyleXml);

                }
            }
        }

        internal static string GetXmlByStyleID(string typeName, string styleID)
        {
            if (!string.IsNullOrEmpty(typeName) && !String.IsNullOrEmpty(styleID))
            {
                return Xml2XmlFormat(StyleSet.FirstOrDefault(p => p.TypeName == typeName).Styles.FirstOrDefault(p => p.Key == styleID).Value);
            }
            return string.Empty;
        }

        private static string Xml2XmlFormat(string xml)
        {
            xml = ("<w:styles>" + xml + "</w:styles>").Replace("w:", string.Empty).Replace("xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"", string.Empty);

            XmlDocument m_XmlDocument = new XmlDocument();
            m_XmlDocument.LoadXml(xml);

            return ChildNode2XmlStringFormat(string.Empty, m_XmlDocument.ChildNodes);
        }

        private static string ChildNode2XmlStringFormat(string space, XmlNodeList Node)
        {
            StringBuilder m_XmlStringBuilder = new StringBuilder();

            foreach (XmlNode node in Node)
            {
                if (node.HasChildNodes)
                {
                    m_XmlStringBuilder.AppendLine(space + "<" + node.Name + GetXmlAttribute(node) + ">");
                    m_XmlStringBuilder.Append(ChildNode2XmlStringFormat(space + "   ", node.ChildNodes));
                }
                else
                {
                    m_XmlStringBuilder.AppendLine(space + "<" + node.Name + GetXmlAttribute(node) + ">");
                }
            }
            return m_XmlStringBuilder.ToString();
        }

        private static string GetXmlAttribute(XmlNode node)
        {
            if (node.Attributes.Count == 0) return string.Empty;

            StringBuilder m_AttributeStringBuilder = new StringBuilder();

            foreach (XmlAttribute attribute in node.Attributes)
            {
                m_AttributeStringBuilder.Append(" " + attribute.Name + "=\"" + attribute.Value + "\"");
            }

            return m_AttributeStringBuilder.ToString();
        }
    }
}
