using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Wordprocessing
{
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
}
