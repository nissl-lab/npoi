using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;
using System.IO;
using System.Xml;

namespace NPOI.OpenXmlFormats.Wordprocessing
{
    public class DocumentDocument
    {
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Document));
        //internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
        //    new XmlQualifiedName("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"), 
        //    new XmlQualifiedName("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
        //    new XmlQualifiedName("m", "http://schemas.openxmlformats.org/officeDocument/2006/math"),
        //    new XmlQualifiedName("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
        //    new XmlQualifiedName("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships") });

        CT_Document document = null;
        public DocumentDocument()
        {

        }
        public static DocumentDocument Parse(Stream stream)
        {
            XmlReaderSettings xmlrs = new XmlReaderSettings();
            xmlrs.IgnoreWhitespace = false;
            using (XmlReader xmlr = XmlReader.Create(stream, xmlrs))
            {
                CT_Document obj = (CT_Document)serializer.Deserialize(xmlr);
                return new DocumentDocument(obj);
            }
            //XmlTextReader xmlReader = new XmlTextReader(stream);
            //xmlReader.WhitespaceHandling = WhitespaceHandling.All;
            

            
        }

        public DocumentDocument(CT_Document document)
        {
            this.document = document;
        }
        public CT_Document Document
        {
            get
            {
                return this.document;
            }
        }
        public void Save(Stream stream, XmlSerializerNamespaces namespaces)
        {
            serializer.Serialize(stream, document, namespaces);
        }
    }
}
