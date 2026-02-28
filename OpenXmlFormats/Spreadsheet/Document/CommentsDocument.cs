using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class CommentsDocument
    {
        CT_Comments comments = null;

        public CommentsDocument()
        { 
        }
        public CommentsDocument(CT_Comments comments)
        {
            this.comments = comments;
        }
        public static CommentsDocument Parse(XmlDocument xmlDoc, XmlNamespaceManager namespaceManager)
        {
            CommentsDocument commentsDoc = new CommentsDocument();
            commentsDoc.comments = CT_Comments.Parse(xmlDoc.DocumentElement, namespaceManager);
            return commentsDoc;
        }
        public CT_Comments GetComments()
        {
            return comments;
        }
        public void SetComments(CT_Comments comments)
        {
            this.comments = comments;
        }
        public void Save(Stream stream)
        {
            using (StreamWriter sw = new StreamWriter(stream))
            {
                this.comments.Write(sw);
            }
        }
    }
}
