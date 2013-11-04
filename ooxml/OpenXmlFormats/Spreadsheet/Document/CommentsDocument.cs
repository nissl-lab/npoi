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
        public static CommentsDocument Parse(XmlDocument xmlDoc, XmlNamespaceManager NameSpaceManager)
        {
            CommentsDocument commentsDoc = new CommentsDocument();
            commentsDoc.comments = new CT_Comments();
            foreach (XmlElement node in xmlDoc.SelectNodes("//d:authors/d:author", NameSpaceManager))
            {
                commentsDoc.comments.authors.AddAuthor(node.InnerText);
            }
            foreach (XmlElement node in xmlDoc.SelectNodes("//d:commentList/d:comment", NameSpaceManager))
            {
                CT_Comment comment = commentsDoc.comments.commentList.AddNewComment();
                comment.authorId = uint.Parse(node.GetAttribute("authorId"));
                comment.@ref = node.GetAttribute("ref");
                comment.text = CT_Rst.Parse(node.ChildNodes[0], NameSpaceManager);
            }
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
            StreamWriter sw = new StreamWriter(stream);
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>");
            sw.Write("<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            sw.Write("<authors>");
            foreach (string author in comments.authors.author)
            {
                sw.Write("<author>");
                sw.Write(author);
                sw.Write("</author>");
            }
            sw.Write("</authors>");
            sw.Write("<commentList>");
            foreach (CT_Comment comment in comments.commentList.comment)
            {
                sw.Write(string.Format("<comment ref=\"{0}\" authorId=\"{1}\">", comment.@ref, comment.authorId));
                sw.Write("<text>");
                sw.Write(comment.text.XmlText);
                sw.Write("</text>");
                sw.Write("</comment>");
            }
            sw.Write("</commentList>");
            sw.Write("</comments>");
            sw.Flush();
        }
    }
}
