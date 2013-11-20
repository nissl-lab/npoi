using System;
using System.Collections.Generic;
using System.Xml.Serialization;
using NPOI.OpenXmlFormats.Shared;
using System.Xml;
using NPOI.OpenXml4Net.Util;
using System.IO;

namespace NPOI.OpenXmlFormats.Wordprocessing
{

    [XmlInclude(typeof(CT_Bookmark))]
    [XmlInclude(typeof(CT_MoveBookmark))]

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_BookmarkRange : CT_MarkupRange
    {

        private string colFirstField;

        private string colLastField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string colFirst
        {
            get
            {
                return this.colFirstField;
            }
            set
            {
                this.colFirstField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string colLast
        {
            get
            {
                return this.colLastField;
            }
            set
            {
                this.colLastField = value;
            }
        }
    }

    [XmlInclude(typeof(CT_MoveBookmark))]

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Bookmark : CT_BookmarkRange
    {

        private string nameField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_MoveBookmark : CT_Bookmark
    {

        private string authorField;

        private string dateField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string author
        {
            get
            {
                return this.authorField;
            }
            set
            {
                this.authorField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string date
        {
            get
            {
                return this.dateField;
            }
            set
            {
                this.dateField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot("comments", Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = false)]
    public class CT_Comments
    {

        private List<CT_Comment> commentField;

        public CT_Comments()
        {
            //this.commentField = new List<CT_Comment>();
        }

        [XmlElement("comment", Order = 0)]
        public List<CT_Comment> comment
        {
            get
            {
                return this.commentField;
            }
            set
            {
                this.commentField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Comment : CT_TrackChange
    {


        private string initialsField;

        public CT_Comment()
        {
        }
        List<CT_MarkupRange> commentRangeStartField;
        public List<CT_MarkupRange> commentRangeStart
        {
            get { return this.commentRangeStartField; }
            set { this.commentRangeStartField = value; }
        }

        List<CT_TrackChange> customXmlDelRangeStartField;
        public List<CT_TrackChange> customXmlDelRangeStart
        {
            get { return this.customXmlDelRangeStartField; }
            set { this.customXmlDelRangeStartField = value; }
        }

        List<CT_TrackChange> customXmlInsRangeStartField;
        public List<CT_TrackChange> customXmlInsRangeStart
        {
            get { return this.customXmlInsRangeStartField; }
            set { this.customXmlInsRangeStartField = value; }
        }

        List<CT_Markup> customXmlMoveFromRangeEndField;
        public List<CT_Markup> customXmlMoveFromRangeEnd
        {
            get { return this.customXmlMoveFromRangeEndField; }
            set { this.customXmlMoveFromRangeEndField = value; }
        }

        List<CT_TrackChange> customXmlMoveFromRangeStartField;
        public List<CT_TrackChange> customXmlMoveFromRangeStart
        {
            get { return this.customXmlMoveFromRangeStartField; }
            set { this.customXmlMoveFromRangeStartField = value; }
        }

        List<CT_Markup> customXmlMoveToRangeEndField;
        public List<CT_Markup> customXmlMoveToRangeEnd
        {
            get { return this.customXmlMoveToRangeEndField; }
            set { this.customXmlMoveToRangeEndField = value; }
        }

        List<CT_TrackChange> customXmlMoveToRangeStartField;
        public List<CT_TrackChange> customXmlMoveToRangeStart
        {
            get { return this.customXmlMoveToRangeStartField; }
            set { this.customXmlMoveToRangeStartField = value; }
        }

        List<CT_RunTrackChange> delField;
        public List<CT_RunTrackChange> del
        {
            get { return this.delField; }
            set { this.delField = value; }
        }

        List<CT_RunTrackChange> insField;
        public List<CT_RunTrackChange> ins
        {
            get { return this.insField; }
            set { this.insField = value; }
        }

        List<CT_RunTrackChange> moveFromField;
        public List<CT_RunTrackChange> moveFrom
        {
            get { return this.moveFromField; }
            set { this.moveFromField = value; }
        }

        List<CT_Markup> customXmlDelRangeEndField;
        public List<CT_Markup> customXmlDelRangeEnd
        {
            get { return this.customXmlDelRangeEndField; }
            set { this.customXmlDelRangeEndField = value; }
        }

        List<CT_MoveBookmark> moveFromRangeStartField;
        public List<CT_MoveBookmark> moveFromRangeStart
        {
            get { return this.moveFromRangeStartField; }
            set { this.moveFromRangeStartField = value; }
        }

        List<CT_RunTrackChange> moveToField;
        public List<CT_RunTrackChange> moveTo
        {
            get { return this.moveToField; }
            set { this.moveToField = value; }
        }

        List<CT_MarkupRange> moveToRangeEndField;
        public List<CT_MarkupRange> moveToRangeEnd
        {
            get { return this.moveToRangeEndField; }
            set { this.moveToRangeEndField = value; }
        }

        List<CT_MoveBookmark> moveToRangeStartField;
        public List<CT_MoveBookmark> moveToRangeStart
        {
            get { return this.moveToRangeStartField; }
            set { this.moveToRangeStartField = value; }
        }

        List<CT_P> pField;
        public List<CT_P> p
        {
            get { return this.pField; }
            set { this.pField = value; }
        }

        List<CT_Perm> permEndField;
        public List<CT_Perm> permEnd
        {
            get { return this.permEndField; }
            set { this.permEndField = value; }
        }

        List<CT_PermStart> permStartField;
        public List<CT_PermStart> permStart
        {
            get { return this.permStartField; }
            set { this.permStartField = value; }
        }

        List<CT_ProofErr> proofErrField;
        public List<CT_ProofErr> proofErr
        {
            get { return this.proofErrField; }
            set { this.proofErrField = value; }
        }

        List<CT_SdtBlock> sdtField;
        public List<CT_SdtBlock> sdt
        {
            get { return this.sdtField; }
            set { this.sdtField = value; }
        }

        List<CT_Tbl> tblField;
        public List<CT_Tbl> tbl
        {
            get { return this.tblField; }
            set { this.tblField = value; }
        }

        List<CT_MarkupRange> moveFromRangeEndField;
        public List<CT_MarkupRange> moveFromRangeEnd
        {
            get { return this.moveFromRangeEndField; }
            set { this.moveFromRangeEndField = value; }
        }

        List<CT_OMath> oMathField;
        public List<CT_OMath> oMath
        {
            get { return this.oMathField; }
            set { this.oMathField = value; }
        }

        List<CT_OMathPara> oMathParaField;
        public List<CT_OMathPara> oMathPara
        {
            get { return this.oMathParaField; }
            set { this.oMathParaField = value; }
        }

        List<CT_AltChunk> altChunkField;
        public List<CT_AltChunk> altChunk
        {
            get { return this.altChunkField; }
            set { this.altChunkField = value; }
        }

        List<CT_Markup> customXmlInsRangeEndField;
        public List<CT_Markup> customXmlInsRangeEnd
        {
            get { return this.customXmlInsRangeEndField; }
            set { this.customXmlInsRangeEndField = value; }
        }

        List<CT_MarkupRange> bookmarkEndField;
        public List<CT_MarkupRange> bookmarkEnd
        {
            get { return this.bookmarkEndField; }
            set { this.bookmarkEndField = value; }
        }

        List<CT_Bookmark> bookmarkStartField;
        public List<CT_Bookmark> bookmarkStart
        {
            get { return this.bookmarkStartField; }
            set { this.bookmarkStartField = value; }
        }

        List<CT_MarkupRange> commentRangeEndField;
        public List<CT_MarkupRange> commentRangeEnd
        {
            get { return this.commentRangeEndField; }
            set { this.commentRangeEndField = value; }
        }

        List<CT_CustomXmlBlock> customXmlField;
        public List<CT_CustomXmlBlock> customXml
        {
            get { return this.customXmlField; }
            set { this.customXmlField = value; }
        }


        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string initials
        {
            get
            {
                return this.initialsField;
            }
            set
            {
                this.initialsField = value;
            }
        }

        public List<CT_P> GetPList()
        {
            return pField;
        }
    }

    [XmlInclude(typeof(CT_RunTrackChange))]
    [XmlInclude(typeof(CT_RPrChange))]
    [XmlInclude(typeof(CT_ParaRPrChange))]
    [XmlInclude(typeof(CT_PPrChange))]
    [XmlInclude(typeof(CT_SectPrChange))]
    [XmlInclude(typeof(CT_TblPrChange))]
    [XmlInclude(typeof(CT_TrPrChange))]
    [XmlInclude(typeof(CT_TcPrChange))]
    [XmlInclude(typeof(CT_TblPrExChange))]
    [XmlInclude(typeof(CT_TrackChangeNumbering))]
    [XmlInclude(typeof(CT_Comment))]
    [XmlInclude(typeof(CT_TrackChangeRange))]
    [XmlInclude(typeof(CT_CellMergeTrackChange))]

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TrackChange : CT_Markup
    {

        private string authorField;

        private string dateField;

        private bool dateFieldSpecified;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string author
        {
            get
            {
                return this.authorField;
            }
            set
            {
                this.authorField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string date
        {
            get
            {
                return this.dateField;
            }
            set
            {
                this.dateField = value;
            }
        }

        [XmlIgnore]
        public bool dateSpecified
        {
            get
            {
                return this.dateFieldSpecified;
            }
            set
            {
                this.dateFieldSpecified = value;
            }
        }
        public static CT_TrackChange Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TrackChange ctObj = new CT_TrackChange();
            ctObj.author = XmlHelper.ReadString(node.Attributes["w:author"]);
            ctObj.date = XmlHelper.ReadString(node.Attributes["w:date"]);
            ctObj.id = XmlHelper.ReadString(node.Attributes["r:id"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:author", this.author);
            XmlHelper.WriteAttribute(sw, "w:date", this.date);
            XmlHelper.WriteAttribute(sw, "r:id", this.id);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }



    }

    [XmlInclude(typeof(CT_TblGridChange))]
    [XmlInclude(typeof(CT_MarkupRange))]
    [XmlInclude(typeof(CT_BookmarkRange))]
    [XmlInclude(typeof(CT_Bookmark))]
    [XmlInclude(typeof(CT_MoveBookmark))]
    [XmlInclude(typeof(CT_TrackChange))]
    [XmlInclude(typeof(CT_RunTrackChange))]
    [XmlInclude(typeof(CT_RPrChange))]
    [XmlInclude(typeof(CT_ParaRPrChange))]
    [XmlInclude(typeof(CT_PPrChange))]
    [XmlInclude(typeof(CT_SectPrChange))]
    [XmlInclude(typeof(CT_TblPrChange))]
    [XmlInclude(typeof(CT_TrPrChange))]
    [XmlInclude(typeof(CT_TcPrChange))]
    [XmlInclude(typeof(CT_TblPrExChange))]
    [XmlInclude(typeof(CT_TrackChangeNumbering))]
    [XmlInclude(typeof(CT_Comment))]
    [XmlInclude(typeof(CT_TrackChangeRange))]
    [XmlInclude(typeof(CT_CellMergeTrackChange))]

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Markup
    {

        private string idField;

        public static CT_Markup Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Markup ctObj = new CT_Markup();
            ctObj.id = XmlHelper.ReadString(node.Attributes["r:id"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "r:id", this.id);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id
        {
            get
            {
                return this.idField;
            }
            set
            {
                this.idField = value;
            }
        }
    }




    [XmlInclude(typeof(CT_BookmarkRange))]
    [XmlInclude(typeof(CT_Bookmark))]
    [XmlInclude(typeof(CT_MoveBookmark))]

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_MarkupRange : CT_Markup
    {
        public static CT_MarkupRange Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_MarkupRange ctObj = new CT_MarkupRange();
            if (node.Attributes["w:displacedByCustomXml"] != null)
                ctObj.displacedByCustomXml = (ST_DisplacedByCustomXml)Enum.Parse(typeof(ST_DisplacedByCustomXml), node.Attributes["w:displacedByCustomXml"].Value);
            ctObj.id = XmlHelper.ReadString(node.Attributes["r:id"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:displacedByCustomXml", this.displacedByCustomXml.ToString());
            XmlHelper.WriteAttribute(sw, "r:id", this.id);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        private ST_DisplacedByCustomXml displacedByCustomXmlField;

        private bool displacedByCustomXmlFieldSpecified;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_DisplacedByCustomXml displacedByCustomXml
        {
            get
            {
                return this.displacedByCustomXmlField;
            }
            set
            {
                this.displacedByCustomXmlField = value;
            }
        }

        [XmlIgnore]
        public bool displacedByCustomXmlSpecified
        {
            get
            {
                return this.displacedByCustomXmlFieldSpecified;
            }
            set
            {
                this.displacedByCustomXmlFieldSpecified = value;
            }
        }
    }
}
