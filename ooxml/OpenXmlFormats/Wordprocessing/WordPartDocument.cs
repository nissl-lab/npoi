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
        //internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new[] {
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
    public class FootnotesDocument
    {
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Footnotes));
        //internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new[] {
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
        //internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new[] {
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
    
    public class StylesDocument
    {
        //TODO: add namespace according the documnet file.
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Styles));
        //internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new[] {
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

    public class NumberingDocument
    {
        //TODO: add namespace according the documnet file.
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Numbering));
        //internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new[] {
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
        public static NumberingDocument Parse(Stream stream)
        {
            CT_Numbering obj = (CT_Numbering)serializer.Deserialize(stream);

            return new NumberingDocument(obj);
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
        public void Save(Stream stream, XmlSerializerNamespaces namespaces)
        {
            serializer.Serialize(stream, numbering, namespaces);
        }
    }
    public class SettingsDocument
    {
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Settings));

        CT_Settings settings = null;
        public SettingsDocument()
        {
            settings = new CT_Settings();
        }
        public static SettingsDocument Parse(Stream stream)
        {
            CT_Settings obj = (CT_Settings)serializer.Deserialize(stream);

            return new SettingsDocument(obj);
        }
        public SettingsDocument(CT_Settings settings)
        {
            this.settings = settings;
        }
        public CT_Settings Settings
        {
            get
            {
                return this.settings;
            }
        }
        public void Save(Stream stream, XmlSerializerNamespaces namespaces)
        {
            serializer.Serialize(stream, settings, namespaces);
        }

    }

    public class CommentsDocument
    {
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Comments));

        CT_Comments comments = null;
        public CommentsDocument()
        {
            comments = new CT_Comments();
        }
        public static CommentsDocument Parse(Stream stream)
        {
            CT_Comments obj = (CT_Comments)serializer.Deserialize(stream);

            return new CommentsDocument(obj);
        }
        public CommentsDocument(CT_Comments comments)
        {
            this.comments = comments;
        }
        public CT_Comments Comments
        {
            get
            {
                return this.comments;
            }
        }
        public void Save(Stream stream, XmlSerializerNamespaces namespaces)
        {
            serializer.Serialize(stream, comments, namespaces);
        }
    }
    public class HdrDocument
    {
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Hdr));

        CT_Hdr hdr = null;
        public HdrDocument()
        {
            hdr = new CT_Hdr();
        }
        public static HdrDocument Parse(Stream stream)
        {
            CT_Hdr obj = (CT_Hdr)serializer.Deserialize(stream);

            return new HdrDocument(obj);
        }
        public HdrDocument(CT_Hdr hdr)
        {
            this.hdr = hdr;
        }
        public CT_Hdr Hdr
        {
            get
            {
                return this.hdr;
            }
        }
        public void Save(Stream stream, XmlSerializerNamespaces namespaces)
        {
            serializer.Serialize(stream, hdr, namespaces);
        }

        public void SetHdr(CT_Hdr hdr)
        {
            this.hdr = hdr;
        }
    }
    public class FtrDocument
    {
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Ftr));

        CT_Ftr ftr = null;
        public FtrDocument()
        {
            ftr = new CT_Ftr();
        }
        public static FtrDocument Parse(Stream stream)
        {
            CT_Ftr obj = (CT_Ftr)serializer.Deserialize(stream);

            return new FtrDocument(obj);
        }
        public FtrDocument(CT_Ftr ftr)
        {
            this.ftr = ftr;
        }
        public CT_Ftr Ftr
        {
            get
            {
                return this.ftr;
            }
        }
        public void Save(Stream stream, XmlSerializerNamespaces namespaces)
        {
            serializer.Serialize(stream, ftr, namespaces);
        }

        public void SetFtr(CT_Ftr ftr)
        {
            this.ftr = ftr;
        }
    }

}
