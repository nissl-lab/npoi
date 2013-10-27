using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Wordprocessing
{
    public class StylesDocument
    {
        //TODO: add namespace according the documnet file.
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Styles));
        //internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
        //    new XmlQualifiedName("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"), 
        //    new XmlQualifiedName("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
        //    new XmlQualifiedName("m", "http://schemas.openxmlformats.org/officeDocument/2006/math"),
        //    new XmlQualifiedName("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
        //    new XmlQualifiedName("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships") });

        CT_Styles styles = null;
        public StylesDocument()
        {
            styles = new CT_Styles();
        }
        public static StylesDocument Parse(Stream stream)
        {
            CT_Styles obj = (CT_Styles)serializer.Deserialize(stream);

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
        public void Save(Stream stream, XmlSerializerNamespaces namespaces)
        {
            serializer.Serialize(stream, styles, namespaces);
        }
    }
}
