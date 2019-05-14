using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    //[System.Diagnostics.DebuggerStepThrough]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ElementName = "comments")]
    public class CT_Comments
    {
        public static CT_Comments Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Comments ctObj = new CT_Comments();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(authors))
                    ctObj.authors = CT_Authors.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(commentList))
                    ctObj.commentList = CT_CommentList.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(extLst))
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }
        internal void Write(StreamWriter sw)
        {
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>");
            sw.Write("<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            this.authors?.Write(sw, nameof(authors));
            this.commentList?.Write(sw, nameof(commentList));
            this.extLst?.Write(sw, nameof(extLst));
            sw.Write("</comments>");
        }
        public CT_Authors AddNewAuthors()
        {
            this.authors = new CT_Authors();
            return this.authors;
        }
        public void AddNewCommentList()
        {
            this.commentList = new CT_CommentList();
        }

        [XmlElement(nameof(authors), Order = 0)]
        public CT_Authors authors { get; set; } = new CT_Authors();
        [XmlElement(nameof(commentList), Order = 1)]
        public CT_CommentList commentList { get; set; } = new CT_CommentList();

        [XmlElement(nameof(extLst), Order = 2)]
        public CT_ExtensionList extLst { get; set; } = null;
    }
}
