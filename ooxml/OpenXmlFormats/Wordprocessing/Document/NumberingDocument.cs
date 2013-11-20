using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Wordprocessing
{

    public class NumberingDocument
    {
        //TODO: add namespace according the documnet file.
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Numbering));
        //internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
        //    new XmlQualifiedName("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"), 
        //    new XmlQualifiedName("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
        //    new XmlQualifiedName("m", "http://schemas.openxmlformats.org/officeDocument/2006/math"),
        //    new XmlQualifiedName("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
        //    new XmlQualifiedName("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships") });

        CT_Numbering numbering = null;
        public NumberingDocument()
        {
            numbering = new CT_Numbering();
        }
        public NumberingDocument(CT_Numbering numbering)
        {
            this.numbering = numbering;
        }
        public CT_Numbering Numbering
        {
            get
            {
                return this.numbering;
            }
        }
        public void Save(Stream stream)
        {
            using (StreamWriter sw = new StreamWriter(stream))
            {
                numbering.Write(sw);
            }
        }

        public static NumberingDocument Parse(System.Xml.XmlDocument doc, System.Xml.XmlNamespaceManager NameSpaceManager)
        {
            CT_Numbering obj = CT_Numbering.Parse(doc, NameSpaceManager);
            return new NumberingDocument(obj);
        }
    }
}
