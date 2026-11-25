using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    /// <summary>
    /// Represents the root <c>ThreadedComments</c> element of a threaded comments part
    /// (namespace http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments).
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments")]
    [XmlRoot(Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments",
        ElementName = "ThreadedComments")]
    public class CT_ThreadedComments
    {
        private List<CT_ThreadedComment> threadedCommentField = null; // optional field [0..*]

        public static CT_ThreadedComments Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ThreadedComments ctObj = new CT_ThreadedComments();
            ctObj.threadedComment = new List<CT_ThreadedComment>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "threadedComment")
                    ctObj.threadedComment.Add(CT_ThreadedComment.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw)
        {
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>");
            sw.Write("<ThreadedComments xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            if (this.threadedComment != null)
            {
                foreach (CT_ThreadedComment x in this.threadedComment)
                {
                    x.Write(sw, "threadedComment");
                }
            }
            sw.Write("</ThreadedComments>");
        }

        [XmlElement("threadedComment")]
        public List<CT_ThreadedComment> threadedComment
        {
            get { return this.threadedCommentField; }
            set { this.threadedCommentField = value; }
        }

        public int SizeOfThreadedCommentArray()
        {
            return (null == threadedCommentField) ? 0 : threadedCommentField.Count;
        }

        public CT_ThreadedComment GetThreadedCommentArray(int index)
        {
            return threadedCommentField[index];
        }

        public CT_ThreadedComment InsertNewThreadedComment(int index)
        {
            if (null == threadedCommentField) { threadedCommentField = new List<CT_ThreadedComment>(); }
            CT_ThreadedComment ct = new CT_ThreadedComment();
            threadedCommentField.Insert(index, ct);
            return ct;
        }

        public CT_ThreadedComment AddNewThreadedComment()
        {
            if (null == threadedCommentField) { threadedCommentField = new List<CT_ThreadedComment>(); }
            CT_ThreadedComment ct = new CT_ThreadedComment();
            threadedCommentField.Add(ct);
            return ct;
        }

        public void RemoveThreadedComment(int index)
        {
            threadedCommentField.RemoveAt(index);
        }

        public CT_ThreadedComment[] GetThreadedCommentArray()
        {
            if (this.threadedCommentField == null)
                this.threadedCommentField = new List<CT_ThreadedComment>();
            return this.threadedCommentField.ToArray();
        }
    }
}
