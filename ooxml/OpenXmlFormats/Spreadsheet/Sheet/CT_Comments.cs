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
        private CT_Authors authorsField = new CT_Authors(); // required field

        private CT_CommentList commentListField = new CT_CommentList(); // required field

        private CT_ExtensionList extLstField = null; // optional field
        public static CT_Comments Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Comments ctObj = new CT_Comments();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "authors")
                    ctObj.authors = CT_Authors.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "commentList")
                    ctObj.commentList = CT_CommentList.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }
        internal void Write(StreamWriter sw)
        {
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>");
            sw.Write("<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            if (this.authors != null)
                this.authors.Write(sw, "authors");
            if (this.commentList != null)
                this.commentList.Write(sw, "commentList");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            sw.Write("</comments>");
        }
        public CT_Authors AddNewAuthors()
        {
            this.authorsField = new CT_Authors();
            return this.authorsField;
        }
        public void AddNewCommentList()
        {
            this.commentListField = new CT_CommentList();
        }

        [XmlElement("authors", Order = 0)]
        public CT_Authors authors
        {
            get
            {
                return this.authorsField;
            }
            set
            {
                this.authorsField = value;
            }
        }
        [XmlElement("commentList", Order = 1)]
        public CT_CommentList commentList
        {
            get
            {
                return this.commentListField;
            }
            set
            {
                this.commentListField = value;
            }
        }

        [XmlElement("extLst", Order = 2)]
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }
    }
}
