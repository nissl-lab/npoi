using System;
using System.Collections.Generic;

using System.Text;
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

        //public CT_Comments()
        //{
        //    this.extLstField = new CT_ExtensionList();
        //    this.commentListField = new CT_CommentList();
        //    this.authorsField = new CT_Authors();
        //}
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
