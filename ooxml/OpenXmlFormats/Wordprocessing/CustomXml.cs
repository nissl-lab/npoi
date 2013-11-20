using System;
using System.Collections.Generic;
using System.Xml.Serialization;
using NPOI.OpenXmlFormats.Shared;

namespace NPOI.OpenXmlFormats.Wordprocessing
{

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_CustomXmlRow
    {

        private CT_CustomXmlPr customXmlPrField;

        private object[] itemsField;

        private ItemsChoiceType21[] itemsElementNameField;

        private string uriField;

        private string elementField;

        public CT_CustomXmlRow()
        {
            this.itemsElementNameField = new ItemsChoiceType21[0];
            this.itemsField = new object[0];
            this.customXmlPrField = new CT_CustomXmlPr();
        }

        [XmlElement(Order = 0)]
        public CT_CustomXmlPr customXmlPr
        {
            get
            {
                return this.customXmlPrField;
            }
            set
            {
                this.customXmlPrField = value;
            }
        }

        [XmlElement("oMath", typeof(CT_OMath), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 1)]
        [XmlElement("oMathPara", typeof(CT_OMathPara), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 1)]
        [XmlElement("bookmarkEnd", typeof(CT_MarkupRange), Order = 1)]
        [XmlElement("bookmarkStart", typeof(CT_Bookmark), Order = 1)]
        [XmlElement("commentRangeEnd", typeof(CT_MarkupRange), Order = 1)]
        [XmlElement("commentRangeStart", typeof(CT_MarkupRange), Order = 1)]
        [XmlElement("customXml", typeof(CT_CustomXmlRow), Order = 1)]
        [XmlElement("customXmlDelRangeEnd", typeof(CT_Markup), Order = 1)]
        [XmlElement("customXmlDelRangeStart", typeof(CT_TrackChange), Order = 1)]
        [XmlElement("customXmlInsRangeEnd", typeof(CT_Markup), Order = 1)]
        [XmlElement("customXmlInsRangeStart", typeof(CT_TrackChange), Order = 1)]
        [XmlElement("customXmlMoveFromRangeEnd", typeof(CT_Markup), Order = 1)]
        [XmlElement("customXmlMoveFromRangeStart", typeof(CT_TrackChange), Order = 1)]
        [XmlElement("customXmlMoveToRangeEnd", typeof(CT_Markup), Order = 1)]
        [XmlElement("customXmlMoveToRangeStart", typeof(CT_TrackChange), Order = 1)]
        [XmlElement("del", typeof(CT_RunTrackChange), Order = 1)]
        [XmlElement("ins", typeof(CT_RunTrackChange), Order = 1)]
        [XmlElement("moveFrom", typeof(CT_RunTrackChange), Order = 1)]
        [XmlElement("moveFromRangeEnd", typeof(CT_MarkupRange), Order = 1)]
        [XmlElement("moveFromRangeStart", typeof(CT_MoveBookmark), Order = 1)]
        [XmlElement("moveTo", typeof(CT_RunTrackChange), Order = 1)]
        [XmlElement("moveToRangeEnd", typeof(CT_MarkupRange), Order = 1)]
        [XmlElement("moveToRangeStart", typeof(CT_MoveBookmark), Order = 1)]
        [XmlElement("permEnd", typeof(CT_Perm), Order = 1)]
        [XmlElement("permStart", typeof(CT_PermStart), Order = 1)]
        [XmlElement("proofErr", typeof(CT_ProofErr), Order = 1)]
        [XmlElement("sdt", typeof(CT_SdtRow), Order = 1)]
        [XmlElement("tr", typeof(CT_Row), Order = 1)]
        [XmlChoiceIdentifier("ItemsElementName")]
        public object[] Items
        {
            get
            {
                return this.itemsField;
            }
            set
            {
                this.itemsField = value;
            }
        }

        [XmlElement("ItemsElementName", Order = 2)]
        [XmlIgnore]
        public ItemsChoiceType21[] ItemsElementName
        {
            get
            {
                return this.itemsElementNameField;
            }
            set
            {
                this.itemsElementNameField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string uri
        {
            get
            {
                return this.uriField;
            }
            set
            {
                this.uriField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string element
        {
            get
            {
                return this.elementField;
            }
            set
            {
                this.elementField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_CustomXmlPr
    {

        private CT_String placeholderField;

        private List<CT_Attr> attrField;

        public CT_CustomXmlPr()
        {
            this.attrField = new List<CT_Attr>();
            this.placeholderField = new CT_String();
        }

        [XmlElement(Order = 0)]
        public CT_String placeholder
        {
            get
            {
                return this.placeholderField;
            }
            set
            {
                this.placeholderField = value;
            }
        }

        [XmlElement("attr", Order = 1)]
        public List<CT_Attr> attr
        {
            get
            {
                return this.attrField;
            }
            set
            {
                this.attrField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Attr
    {

        private string uriField;

        private string nameField;

        private string valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string uri
        {
            get
            {
                return this.uriField;
            }
            set
            {
                this.uriField = value;
            }
        }

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

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }



    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType21
    {

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:oMath")]
        oMath,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:oMathPara")]
        oMathPara,

    
        bookmarkEnd,

    
        bookmarkStart,

    
        commentRangeEnd,

    
        commentRangeStart,

    
        customXml,

    
        customXmlDelRangeEnd,

    
        customXmlDelRangeStart,

    
        customXmlInsRangeEnd,

    
        customXmlInsRangeStart,

    
        customXmlMoveFromRangeEnd,

    
        customXmlMoveFromRangeStart,

    
        customXmlMoveToRangeEnd,

    
        customXmlMoveToRangeStart,

    
        del,

    
        ins,

    
        moveFrom,

    
        moveFromRangeEnd,

    
        moveFromRangeStart,

    
        moveTo,

    
        moveToRangeEnd,

    
        moveToRangeStart,

    
        permEnd,

    
        permStart,

    
        proofErr,

    
        sdt,

    
        tr,
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType22
    {

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:oMath")]
        oMath,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:oMathPara")]
        oMathPara,

    
        bookmarkEnd,

    
        bookmarkStart,

    
        commentRangeEnd,

    
        commentRangeStart,

    
        customXml,

    
        customXmlDelRangeEnd,

    
        customXmlDelRangeStart,

    
        customXmlInsRangeEnd,

    
        customXmlInsRangeStart,

    
        customXmlMoveFromRangeEnd,

    
        customXmlMoveFromRangeStart,

    
        customXmlMoveToRangeEnd,

    
        customXmlMoveToRangeStart,

    
        del,

    
        ins,

    
        moveFrom,

    
        moveFromRangeEnd,

    
        moveFromRangeStart,

    
        moveTo,

    
        moveToRangeEnd,

    
        moveToRangeStart,

    
        permEnd,

    
        permStart,

    
        proofErr,

    
        sdt,

    
        tr,
    }



    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_CustomXmlRun
    {

        private CT_CustomXmlPr customXmlPrField;

        private string uriField;

        private string elementField;

        public CT_CustomXmlRun()
        {
            //this.customXmlPrField = new CT_CustomXmlPr();
        }

        [XmlElement(Order = 0)]
        public CT_CustomXmlPr customXmlPr
        {
            get
            {
                return this.customXmlPrField;
            }
            set
            {
                this.customXmlPrField = value;
            }
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

		List<CT_SimpleField> fldSimpleField;
		public List<CT_SimpleField> fldSimple
		{
			get { return this.fldSimpleField; }
			set { this.fldSimpleField = value; }
		}

		List<CT_Hyperlink1> hyperlinkField;
		public List<CT_Hyperlink1> hyperlink
		{
			get { return this.hyperlinkField; }
			set { this.hyperlinkField = value; }
		}

		List<CT_OMath> oMathField;
		public List<CT_OMath> oMath
		{
			get { return this.oMathField; }
			set { this.oMathField = value; }
		}

		List<CT_RunTrackChange> moveFromField;
		public List<CT_RunTrackChange> moveFrom
		{
			get { return this.moveFromField; }
			set { this.moveFromField = value; }
		}

		List<CT_MarkupRange> moveFromRangeEndField;
		public List<CT_MarkupRange> moveFromRangeEnd
		{
			get { return this.moveFromRangeEndField; }
			set { this.moveFromRangeEndField = value; }
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

		List<CT_R> rField;
		public List<CT_R> r
		{
			get { return this.rField; }
			set { this.rField = value; }
		}

		List<CT_SdtRun> sdtField;
		public List<CT_SdtRun> sdt
		{
			get { return this.sdtField; }
			set { this.sdtField = value; }
		}

		List<CT_SmartTagRun> smartTagField;
		public List<CT_SmartTagRun> smartTag
		{
			get { return this.smartTagField; }
			set { this.smartTagField = value; }
		}

		List<CT_Rel> subDocField;
		public List<CT_Rel> subDoc
		{
			get { return this.subDocField; }
			set { this.subDocField = value; }
		}

		List<CT_RunTrackChange> insField;
		public List<CT_RunTrackChange> ins
		{
			get { return this.insField; }
			set { this.insField = value; }
		}

		List<CT_OMathPara> oMathParaField;
		public List<CT_OMathPara> oMathPara
		{
			get { return this.oMathParaField; }
			set { this.oMathParaField = value; }
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

		List<CT_MarkupRange> commentRangeStartField;
		public List<CT_MarkupRange> commentRangeStart
		{
			get { return this.commentRangeStartField; }
			set { this.commentRangeStartField = value; }
		}

		List<CT_CustomXmlRun> customXmlField;
		public List<CT_CustomXmlRun> customXml
		{
			get { return this.customXmlField; }
			set { this.customXmlField = value; }
		}

		List<CT_Markup> customXmlDelRangeEndField;
		public List<CT_Markup> customXmlDelRangeEnd
		{
			get { return this.customXmlDelRangeEndField; }
			set { this.customXmlDelRangeEndField = value; }
		}

		List<CT_TrackChange> customXmlDelRangeStartField;
		public List<CT_TrackChange> customXmlDelRangeStart
		{
			get { return this.customXmlDelRangeStartField; }
			set { this.customXmlDelRangeStartField = value; }
		}

		List<CT_Markup> customXmlInsRangeEndField;
		public List<CT_Markup> customXmlInsRangeEnd
		{
			get { return this.customXmlInsRangeEndField; }
			set { this.customXmlInsRangeEndField = value; }
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



        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string uri
        {
            get
            {
                return this.uriField;
            }
            set
            {
                this.uriField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string element
        {
            get
            {
                return this.elementField;
            }
            set
            {
                this.elementField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType24
    {

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:oMath")]
        oMath,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:oMathPara")]
        oMathPara,

    
        bookmarkEnd,

    
        bookmarkStart,

    
        commentRangeEnd,

    
        commentRangeStart,

    
        customXml,

    
        customXmlDelRangeEnd,

    
        customXmlDelRangeStart,

    
        customXmlInsRangeEnd,

    
        customXmlInsRangeStart,

    
        customXmlMoveFromRangeEnd,

    
        customXmlMoveFromRangeStart,

    
        customXmlMoveToRangeEnd,

    
        customXmlMoveToRangeStart,

    
        del,

    
        fldSimple,

    
        hyperlink,

    
        ins,

    
        moveFrom,

    
        moveFromRangeEnd,

    
        moveFromRangeStart,

    
        moveTo,

    
        moveToRangeEnd,

    
        moveToRangeStart,

    
        permEnd,

    
        permStart,

    
        proofErr,

    
        r,

    
        sdt,

    
        smartTag,

    
        subDoc,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SmartTagRun
    {

        private List<CT_Attr> smartTagPrField;

        private object[] itemsField;

        private ItemsChoiceType25[] itemsElementNameField;

        private string uriField;

        private string elementField;

        public CT_SmartTagRun()
        {
            this.itemsElementNameField = new ItemsChoiceType25[0];
            this.itemsField = new object[0];
            this.smartTagPrField = new List<CT_Attr>();
        }

        [XmlArray(Order = 0)]
        [XmlArrayItem("attr", IsNullable = false)]
        public List<CT_Attr> smartTagPr
        {
            get
            {
                return this.smartTagPrField;
            }
            set
            {
                this.smartTagPrField = value;
            }
        }

        [XmlElement("oMath", typeof(CT_OMath), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 1)]
        [XmlElement("oMathPara", typeof(CT_OMathPara), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 1)]
        [XmlElement("bookmarkEnd", typeof(CT_MarkupRange), Order = 1)]
        [XmlElement("bookmarkStart", typeof(CT_Bookmark), Order = 1)]
        [XmlElement("commentRangeEnd", typeof(CT_MarkupRange), Order = 1)]
        [XmlElement("commentRangeStart", typeof(CT_MarkupRange), Order = 1)]
        [XmlElement("customXml", typeof(CT_CustomXmlRun), Order = 1)]
        [XmlElement("customXmlDelRangeEnd", typeof(CT_Markup), Order = 1)]
        [XmlElement("customXmlDelRangeStart", typeof(CT_TrackChange), Order = 1)]
        [XmlElement("customXmlInsRangeEnd", typeof(CT_Markup), Order = 1)]
        [XmlElement("customXmlInsRangeStart", typeof(CT_TrackChange), Order = 1)]
        [XmlElement("customXmlMoveFromRangeEnd", typeof(CT_Markup), Order = 1)]
        [XmlElement("customXmlMoveFromRangeStart", typeof(CT_TrackChange), Order = 1)]
        [XmlElement("customXmlMoveToRangeEnd", typeof(CT_Markup), Order = 1)]
        [XmlElement("customXmlMoveToRangeStart", typeof(CT_TrackChange), Order = 1)]
        [XmlElement("del", typeof(CT_RunTrackChange), Order = 1)]
        [XmlElement("fldSimple", typeof(CT_SimpleField), Order = 1)]
        [XmlElement("hyperlink", typeof(CT_Hyperlink1), Order = 1)]
        [XmlElement("ins", typeof(CT_RunTrackChange), Order = 1)]
        [XmlElement("moveFrom", typeof(CT_RunTrackChange), Order = 1)]
        [XmlElement("moveFromRangeEnd", typeof(CT_MarkupRange), Order = 1)]
        [XmlElement("moveFromRangeStart", typeof(CT_MoveBookmark), Order = 1)]
        [XmlElement("moveTo", typeof(CT_RunTrackChange), Order = 1)]
        [XmlElement("moveToRangeEnd", typeof(CT_MarkupRange), Order = 1)]
        [XmlElement("moveToRangeStart", typeof(CT_MoveBookmark), Order = 1)]
        [XmlElement("permEnd", typeof(CT_Perm), Order = 1)]
        [XmlElement("permStart", typeof(CT_PermStart), Order = 1)]
        [XmlElement("proofErr", typeof(CT_ProofErr), Order = 1)]
        [XmlElement("r", typeof(CT_R), Order = 1)]
        [XmlElement("sdt", typeof(CT_SdtRun), Order = 1)]
        [XmlElement("smartTag", typeof(CT_SmartTagRun), Order = 1)]
        [XmlElement("subDoc", typeof(CT_Rel), Order = 1)]
        [XmlChoiceIdentifier("ItemsElementName")]
        public object[] Items
        {
            get
            {
                return this.itemsField;
            }
            set
            {
                this.itemsField = value;
            }
        }

        [XmlElement("ItemsElementName", Order = 2)]
        [XmlIgnore]
        public ItemsChoiceType25[] ItemsElementName
        {
            get
            {
                return this.itemsElementNameField;
            }
            set
            {
                this.itemsElementNameField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string uri
        {
            get
            {
                return this.uriField;
            }
            set
            {
                this.uriField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string element
        {
            get
            {
                return this.elementField;
            }
            set
            {
                this.elementField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType25
    {

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:oMath")]
        oMath,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:oMathPara")]
        oMathPara,

    
        bookmarkEnd,

    
        bookmarkStart,

    
        commentRangeEnd,

    
        commentRangeStart,

    
        customXml,

    
        customXmlDelRangeEnd,

    
        customXmlDelRangeStart,

    
        customXmlInsRangeEnd,

    
        customXmlInsRangeStart,

    
        customXmlMoveFromRangeEnd,

    
        customXmlMoveFromRangeStart,

    
        customXmlMoveToRangeEnd,

    
        customXmlMoveToRangeStart,

    
        del,

    
        fldSimple,

    
        hyperlink,

    
        ins,

    
        moveFrom,

    
        moveFromRangeEnd,

    
        moveFromRangeStart,

    
        moveTo,

    
        moveToRangeEnd,

    
        moveToRangeStart,

    
        permEnd,

    
        permStart,

    
        proofErr,

    
        r,

    
        sdt,

    
        smartTag,

    
        subDoc,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_CustomXmlBlock
    {

        private CT_CustomXmlPr customXmlPrField;

        private object[] itemsField;

        private ItemsChoiceType26[] itemsElementNameField;

        private string uriField;

        private string elementField;

        public CT_CustomXmlBlock()
        {
            this.itemsElementNameField = new ItemsChoiceType26[0];
            this.itemsField = new object[0];
            this.customXmlPrField = new CT_CustomXmlPr();
        }

        [XmlElement(Order = 0)]
        public CT_CustomXmlPr customXmlPr
        {
            get
            {
                return this.customXmlPrField;
            }
            set
            {
                this.customXmlPrField = value;
            }
        }

        [XmlElement("oMath", typeof(CT_OMath), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 1)]
        [XmlElement("oMathPara", typeof(CT_OMathPara), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 1)]
        [XmlElement("bookmarkEnd", typeof(CT_MarkupRange), Order = 1)]
        [XmlElement("bookmarkStart", typeof(CT_Bookmark), Order = 1)]
        [XmlElement("commentRangeEnd", typeof(CT_MarkupRange), Order = 1)]
        [XmlElement("commentRangeStart", typeof(CT_MarkupRange), Order = 1)]
        [XmlElement("customXml", typeof(CT_CustomXmlBlock), Order = 1)]
        [XmlElement("customXmlDelRangeEnd", typeof(CT_Markup), Order = 1)]
        [XmlElement("customXmlDelRangeStart", typeof(CT_TrackChange), Order = 1)]
        [XmlElement("customXmlInsRangeEnd", typeof(CT_Markup), Order = 1)]
        [XmlElement("customXmlInsRangeStart", typeof(CT_TrackChange), Order = 1)]
        [XmlElement("customXmlMoveFromRangeEnd", typeof(CT_Markup), Order = 1)]
        [XmlElement("customXmlMoveFromRangeStart", typeof(CT_TrackChange), Order = 1)]
        [XmlElement("customXmlMoveToRangeEnd", typeof(CT_Markup), Order = 1)]
        [XmlElement("customXmlMoveToRangeStart", typeof(CT_TrackChange), Order = 1)]
        [XmlElement("del", typeof(CT_RunTrackChange), Order = 1)]
        [XmlElement("ins", typeof(CT_RunTrackChange), Order = 1)]
        [XmlElement("moveFrom", typeof(CT_RunTrackChange), Order = 1)]
        [XmlElement("moveFromRangeEnd", typeof(CT_MarkupRange), Order = 1)]
        [XmlElement("moveFromRangeStart", typeof(CT_MoveBookmark), Order = 1)]
        [XmlElement("moveTo", typeof(CT_RunTrackChange), Order = 1)]
        [XmlElement("moveToRangeEnd", typeof(CT_MarkupRange), Order = 1)]
        [XmlElement("moveToRangeStart", typeof(CT_MoveBookmark), Order = 1)]
        [XmlElement("p", typeof(CT_P), Order = 1)]
        [XmlElement("permEnd", typeof(CT_Perm), Order = 1)]
        [XmlElement("permStart", typeof(CT_PermStart), Order = 1)]
        [XmlElement("proofErr", typeof(CT_ProofErr), Order = 1)]
        [XmlElement("sdt", typeof(CT_SdtBlock), Order = 1)]
        [XmlElement("tbl", typeof(CT_Tbl), Order = 1)]
        [XmlChoiceIdentifier("ItemsElementName")]
        public object[] Items
        {
            get
            {
                return this.itemsField;
            }
            set
            {
                this.itemsField = value;
            }
        }

        [XmlElement("ItemsElementName", Order = 2)]
        [XmlIgnore]
        public ItemsChoiceType26[] ItemsElementName
        {
            get
            {
                return this.itemsElementNameField;
            }
            set
            {
                this.itemsElementNameField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string uri
        {
            get
            {
                return this.uriField;
            }
            set
            {
                this.uriField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string element
        {
            get
            {
                return this.elementField;
            }
            set
            {
                this.elementField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType26
    {

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:oMath")]
        oMath,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:oMathPara")]
        oMathPara,

    
        bookmarkEnd,

    
        bookmarkStart,

    
        commentRangeEnd,

    
        commentRangeStart,

    
        customXml,

    
        customXmlDelRangeEnd,

    
        customXmlDelRangeStart,

    
        customXmlInsRangeEnd,

    
        customXmlInsRangeStart,

    
        customXmlMoveFromRangeEnd,

    
        customXmlMoveFromRangeStart,

    
        customXmlMoveToRangeEnd,

    
        customXmlMoveToRangeStart,

    
        del,

    
        ins,

    
        moveFrom,

    
        moveFromRangeEnd,

    
        moveFromRangeStart,

    
        moveTo,

    
        moveToRangeEnd,

    
        moveToRangeStart,

    
        p,

    
        permEnd,

    
        permStart,

    
        proofErr,

    
        sdt,

    
        tbl,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_CustomXmlCell
    {

        private CT_CustomXmlPr customXmlPrField;

        private object[] itemsField;

        private ItemsChoiceType27[] itemsElementNameField;

        private string uriField;

        private string elementField;

        public CT_CustomXmlCell()
        {
            this.itemsElementNameField = new ItemsChoiceType27[0];
            this.itemsField = new object[0];
            this.customXmlPrField = new CT_CustomXmlPr();
        }

        [XmlElement(Order = 0)]
        public CT_CustomXmlPr customXmlPr
        {
            get
            {
                return this.customXmlPrField;
            }
            set
            {
                this.customXmlPrField = value;
            }
        }

        [XmlElement("oMath", typeof(CT_OMath), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 1)]
        [XmlElement("oMathPara", typeof(CT_OMathPara), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 1)]
        [XmlElement("bookmarkEnd", typeof(CT_MarkupRange), Order = 1)]
        [XmlElement("bookmarkStart", typeof(CT_Bookmark), Order = 1)]
        [XmlElement("commentRangeEnd", typeof(CT_MarkupRange), Order = 1)]
        [XmlElement("commentRangeStart", typeof(CT_MarkupRange), Order = 1)]
        [XmlElement("customXml", typeof(CT_CustomXmlCell), Order = 1)]
        [XmlElement("customXmlDelRangeEnd", typeof(CT_Markup), Order = 1)]
        [XmlElement("customXmlDelRangeStart", typeof(CT_TrackChange), Order = 1)]
        [XmlElement("customXmlInsRangeEnd", typeof(CT_Markup), Order = 1)]
        [XmlElement("customXmlInsRangeStart", typeof(CT_TrackChange), Order = 1)]
        [XmlElement("customXmlMoveFromRangeEnd", typeof(CT_Markup), Order = 1)]
        [XmlElement("customXmlMoveFromRangeStart", typeof(CT_TrackChange), Order = 1)]
        [XmlElement("customXmlMoveToRangeEnd", typeof(CT_Markup), Order = 1)]
        [XmlElement("customXmlMoveToRangeStart", typeof(CT_TrackChange), Order = 1)]
        [XmlElement("del", typeof(CT_RunTrackChange), Order = 1)]
        [XmlElement("ins", typeof(CT_RunTrackChange), Order = 1)]
        [XmlElement("moveFrom", typeof(CT_RunTrackChange), Order = 1)]
        [XmlElement("moveFromRangeEnd", typeof(CT_MarkupRange), Order = 1)]
        [XmlElement("moveFromRangeStart", typeof(CT_MoveBookmark), Order = 1)]
        [XmlElement("moveTo", typeof(CT_RunTrackChange), Order = 1)]
        [XmlElement("moveToRangeEnd", typeof(CT_MarkupRange), Order = 1)]
        [XmlElement("moveToRangeStart", typeof(CT_MoveBookmark), Order = 1)]
        [XmlElement("permEnd", typeof(CT_Perm), Order = 1)]
        [XmlElement("permStart", typeof(CT_PermStart), Order = 1)]
        [XmlElement("proofErr", typeof(CT_ProofErr), Order = 1)]
        [XmlElement("sdt", typeof(CT_SdtCell), Order = 1)]
        [XmlElement("tc", typeof(CT_Tc), Order = 1)]
        [XmlChoiceIdentifier("ItemsElementName")]
        public object[] Items
        {
            get
            {
                return this.itemsField;
            }
            set
            {
                this.itemsField = value;
            }
        }

        [XmlElement("ItemsElementName", Order = 2)]
        [XmlIgnore]
        public ItemsChoiceType27[] ItemsElementName
        {
            get
            {
                return this.itemsElementNameField;
            }
            set
            {
                this.itemsElementNameField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string uri
        {
            get
            {
                return this.uriField;
            }
            set
            {
                this.uriField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string element
        {
            get
            {
                return this.elementField;
            }
            set
            {
                this.elementField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType27
    {

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:oMath")]
        oMath,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:oMathPara")]
        oMathPara,

    
        bookmarkEnd,

    
        bookmarkStart,

    
        commentRangeEnd,

    
        commentRangeStart,

    
        customXml,

    
        customXmlDelRangeEnd,

    
        customXmlDelRangeStart,

    
        customXmlInsRangeEnd,

    
        customXmlInsRangeStart,

    
        customXmlMoveFromRangeEnd,

    
        customXmlMoveFromRangeStart,

    
        customXmlMoveToRangeEnd,

    
        customXmlMoveToRangeStart,

    
        del,

    
        ins,

    
        moveFrom,

    
        moveFromRangeEnd,

    
        moveFromRangeStart,

    
        moveTo,

    
        moveToRangeEnd,

    
        moveToRangeStart,

    
        permEnd,

    
        permStart,

    
        proofErr,

    
        sdt,

    
        tc,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SmartTagPr
    {

        private List<CT_Attr> attrField;

        public CT_SmartTagPr()
        {
            this.attrField = new List<CT_Attr>();
        }

        [XmlElement("attr", Order = 0)]
        public List<CT_Attr> attr
        {
            get
            {
                return this.attrField;
            }
            set
            {
                this.attrField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum Items1ChoiceType
    {

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:oMath")]
        oMath,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:oMathPara")]
        oMathPara,

    
        bookmarkEnd,

    
        bookmarkStart,

    
        commentRangeEnd,

    
        commentRangeStart,

    
        customXml,

    
        customXmlDelRangeEnd,

    
        customXmlDelRangeStart,

    
        customXmlInsRangeEnd,

    
        customXmlInsRangeStart,

    
        customXmlMoveFromRangeEnd,

    
        customXmlMoveFromRangeStart,

    
        customXmlMoveToRangeEnd,

    
        customXmlMoveToRangeStart,

    
        del,

    
        ins,

    
        moveFrom,

    
        moveFromRangeEnd,

    
        moveFromRangeStart,

    
        moveTo,

    
        moveToRangeEnd,

    
        moveToRangeStart,

    
        permEnd,

    
        permStart,

    
        proofErr,

    
        sdt,

    
        tr,
    }

}
