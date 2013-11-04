using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Wordprocessing
{
    public class FootnotesDocument
    {
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Footnotes));
        //internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
        //    new XmlQualifiedName("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"), 
        //    new XmlQualifiedName("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
        //    new XmlQualifiedName("m", "http://schemas.openxmlformats.org/officeDocument/2006/math"),
        //    new XmlQualifiedName("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
        //    new XmlQualifiedName("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships") });

        CT_Footnotes footnotes = null;
        public FootnotesDocument()
        {
            footnotes = new CT_Footnotes();
        }
        public static FootnotesDocument Parse(Stream stream)
        {
            CT_Footnotes obj = (CT_Footnotes)serializer.Deserialize(stream);

            return new FootnotesDocument(obj);
        }
        public FootnotesDocument(CT_Footnotes footnotes)
        {
            this.footnotes = footnotes;
        }
        public CT_Footnotes Footnotes
        {
            get
            {
                return this.footnotes;
            }
        }
        public void Save(Stream stream, XmlSerializerNamespaces namespaces)
        {
            serializer.Serialize(stream, footnotes, namespaces);
        }
    }

    public class EndnotesDocument
    {
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Endnotes));
        //internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
        //    new XmlQualifiedName("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"), 
        //    new XmlQualifiedName("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
        //    new XmlQualifiedName("m", "http://schemas.openxmlformats.org/officeDocument/2006/math"),
        //    new XmlQualifiedName("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
        //    new XmlQualifiedName("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships") });

        CT_Endnotes endnotes = null;
        public EndnotesDocument()
        {

        }
        public static EndnotesDocument Parse(Stream stream)
        {
            CT_Endnotes obj = (CT_Endnotes)serializer.Deserialize(stream);

            return new EndnotesDocument(obj);
        }
        public EndnotesDocument(CT_Endnotes endnotes)
        {
            this.endnotes = endnotes;
        }
        public CT_Endnotes Endnotes
        {
            get
            {
                return this.endnotes;
            }
        }
        public void Save(Stream stream, XmlSerializerNamespaces namespaces)
        {
            serializer.Serialize(stream, endnotes, namespaces);
        }
    }
    
}
