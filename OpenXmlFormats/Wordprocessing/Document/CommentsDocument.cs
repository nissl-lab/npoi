using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Wordprocessing
{
    public class CommentsDocument
    {
        
        CT_Comments comments = null;
        public CommentsDocument()
        {
            comments = new CT_Comments();
        }
        public static CommentsDocument Parse(XmlDocument doc, XmlNamespaceManager NameSpaceManager)
        {
            CT_Comments obj = CT_Comments.Parse(doc.DocumentElement, NameSpaceManager);
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
        public void Save(Stream stream)
        {
            using (StreamWriter sw = new StreamWriter(stream))
            {
                comments.Write(sw);
            }
        }
    }
}
