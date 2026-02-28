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
        CT_Document document = null;
        public DocumentDocument()
        {

        }
        public static DocumentDocument Parse(XmlDocument doc, XmlNamespaceManager namespaceMgr)
        {
            CT_Document obj = CT_Document.Parse(doc.DocumentElement, namespaceMgr);
            return new DocumentDocument(obj);
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
        public void Save(Stream stream)
        {
            using (StreamWriter sw = new StreamWriter(stream))
            {
                document.Write(sw);
            }
        }
    }
}
