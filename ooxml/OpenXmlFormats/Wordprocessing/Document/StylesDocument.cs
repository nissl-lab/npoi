using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Wordprocessing
{
    public class StylesDocument
    {
        CT_Styles styles = null;
        public StylesDocument()
        {
            styles = new CT_Styles();
        }
        public static StylesDocument Parse(XmlDocument doc, XmlNamespaceManager namespaceMgr)
        {
            CT_Styles obj = CT_Styles.Parse(doc.DocumentElement, namespaceMgr);
            return new StylesDocument(obj);
        }
        public StylesDocument(CT_Styles styles)
        {
            this.styles = styles;
        }
        public CT_Styles Styles
        {
            get
            {
                return this.styles;
            }
        }
        public void Save(Stream stream)
        {
            using (StreamWriter sw = new StreamWriter(stream))
            {
                styles.Write(sw);
            }
        }
    }
}
