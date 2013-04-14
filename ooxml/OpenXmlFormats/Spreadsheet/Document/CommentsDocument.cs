using System.IO;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class CommentsDocument
    {
        static XmlSerializer serializer = new XmlSerializer(typeof(CT_Comments));
        CT_Comments comments = null;

        public CommentsDocument()
        { 
        }
        public CommentsDocument(CT_Comments comments)
        {
            this.comments = comments;
        }
        public static CommentsDocument Parse(Stream stream)
        {
            CT_Comments obj = (CT_Comments)serializer.Deserialize(stream);
            return new CommentsDocument(obj);
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
            serializer.Serialize(stream, comments);
        }
    }
}
