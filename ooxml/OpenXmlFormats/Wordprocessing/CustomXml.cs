using System;
using System.Collections.Generic;
using System.Xml.Serialization;
using NPOI.OpenXmlFormats.Shared;
using System.Collections;
using System.IO;
using System.Xml;
using NPOI.OpenXml4Net.Util;

namespace NPOI.OpenXmlFormats.Wordprocessing
{

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_CustomXmlRow
    {

        private CT_CustomXmlPr customXmlPrField;

        private ArrayList itemsField;

        private List<ItemsChoiceType21> itemsElementNameField;

        private string uriField;

        private string elementField;

        public CT_CustomXmlRow()
        {
            this.itemsElementNameField = new List<ItemsChoiceType21>();
            this.itemsField = new ArrayList();
            //this.customXmlPrField = new CT_CustomXmlPr();
        }

        public static CT_CustomXmlRow Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CustomXmlRow ctObj = new CT_CustomXmlRow();
            ctObj.uri = XmlHelper.ReadString(node.Attributes["w:uri"]);
            ctObj.element = XmlHelper.ReadString(node.Attributes["w:element"]);

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "customXmlMoveFromRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.customXmlMoveFromRangeStart);
                }
                else if (childNode.LocalName == "customXmlPr")
                {
                    ctObj.customXmlPr = CT_CustomXmlPr.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "oMathPara")
                {
                    ctObj.Items.Add(CT_OMathPara.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.oMathPara);
                }
                else if (childNode.LocalName == "sdt")
                {
                    ctObj.Items.Add(CT_SdtRow.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.sdt);
                }
                else if (childNode.LocalName == "customXmlDelRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.customXmlDelRangeStart);
                }
                else if (childNode.LocalName == "customXmlInsRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.customXmlInsRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.customXmlMoveFromRangeEnd);
                }
                else if (childNode.LocalName == "oMath")
                {
                    ctObj.Items.Add(CT_OMath.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.oMath);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.customXmlMoveToRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.customXmlMoveToRangeStart);
                }
                else if (childNode.LocalName == "customXmlDelRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.customXmlDelRangeEnd);
                }
                else if (childNode.LocalName == "ins")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.ins);
                }
                else if (childNode.LocalName == "moveFrom")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.moveFrom);
                }
                else if (childNode.LocalName == "moveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.moveFromRangeEnd);
                }
                else if (childNode.LocalName == "moveFromRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.moveFromRangeStart);
                }
                else if (childNode.LocalName == "moveTo")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.moveTo);
                }
                else if (childNode.LocalName == "moveToRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.moveToRangeEnd);
                }
                else if (childNode.LocalName == "moveToRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.moveToRangeStart);
                }
                else if (childNode.LocalName == "permEnd")
                {
                    ctObj.Items.Add(CT_Perm.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.permEnd);
                }
                else if (childNode.LocalName == "permStart")
                {
                    ctObj.Items.Add(CT_PermStart.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.permStart);
                }
                else if (childNode.LocalName == "proofErr")
                {
                    ctObj.Items.Add(CT_ProofErr.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.proofErr);
                }
                else if (childNode.LocalName == "customXmlInsRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.customXmlInsRangeEnd);
                }
                else if (childNode.LocalName == "tr")
                {
                    ctObj.Items.Add(CT_Row.Parse(childNode, namespaceManager, ctObj));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.tr);
                }
                else if (childNode.LocalName == "del")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.del);
                }
                else if (childNode.LocalName == "bookmarkEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.bookmarkEnd);
                }
                else if (childNode.LocalName == "bookmarkStart")
                {
                    ctObj.Items.Add(CT_Bookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.bookmarkStart);
                }
                else if (childNode.LocalName == "commentRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.commentRangeEnd);
                }
                else if (childNode.LocalName == "commentRangeStart")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.commentRangeStart);
                }
                else if (childNode.LocalName == "customXml")
                {
                    ctObj.Items.Add(CT_CustomXmlRow.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType21.customXml);
                }
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:uri", this.uri);
            XmlHelper.WriteAttribute(sw, "w:element", this.element);
            sw.Write(">");
            if (this.customXmlPr != null)
                this.customXmlPr.Write(sw, "customXmlPr");
            foreach (object o in this.Items)
            {
                if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveFromRangeStart");
                else if (o is CT_OMathPara)
                    ((CT_OMathPara)o).Write(sw, "oMathPara");
                else if (o is CT_SdtRow)
                    ((CT_SdtRow)o).Write(sw, "sdt");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlDelRangeStart");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlInsRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlMoveFromRangeEnd");
                else if (o is CT_OMath)
                    ((CT_OMath)o).Write(sw, "oMath");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlMoveToRangeEnd");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveToRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlDelRangeEnd");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "ins");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "moveFrom");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "moveFromRangeEnd");
                else if (o is CT_MoveBookmark)
                    ((CT_MoveBookmark)o).Write(sw, "moveFromRangeStart");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "moveTo");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "moveToRangeEnd");
                else if (o is CT_MoveBookmark)
                    ((CT_MoveBookmark)o).Write(sw, "moveToRangeStart");
                else if (o is CT_Perm)
                    ((CT_Perm)o).Write(sw, "permEnd");
                else if (o is CT_PermStart)
                    ((CT_PermStart)o).Write(sw, "permStart");
                else if (o is CT_ProofErr)
                    ((CT_ProofErr)o).Write(sw, "proofErr");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlInsRangeEnd");
                else if (o is CT_Row)
                    ((CT_Row)o).Write(sw, "tr");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "del");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "bookmarkEnd");
                else if (o is CT_Bookmark)
                    ((CT_Bookmark)o).Write(sw, "bookmarkStart");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeEnd");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeStart");
                else if (o is CT_CustomXmlRow)
                    ((CT_CustomXmlRow)o).Write(sw, "customXml");
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
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
        public ArrayList Items
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
        public List<ItemsChoiceType21> ItemsElementName
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
        {            //this.attrField = new List<CT_Attr>();
            //this.placeholderField = new CT_String();

        }
        public static CT_CustomXmlPr Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CustomXmlPr ctObj = new CT_CustomXmlPr();
            ctObj.attr = new List<CT_Attr>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "placeholder")
                    ctObj.placeholder = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "attr")
                    ctObj.attr.Add(CT_Attr.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.placeholder != null)
                this.placeholder.Write(sw, "placeholder");
            if (this.attr != null)
            {
                foreach (CT_Attr x in this.attr)
                {
                    x.Write(sw, "attr");
                }
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
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
        public static CT_Attr Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Attr ctObj = new CT_Attr();
            ctObj.uri = XmlHelper.ReadString(node.Attributes["w:uri"]);
            ctObj.name = XmlHelper.ReadString(node.Attributes["w:name"]);
            ctObj.val = XmlHelper.ReadString(node.Attributes["w:val"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:uri", this.uri);
            XmlHelper.WriteAttribute(sw, "w:name", this.name);
            XmlHelper.WriteAttribute(sw, "w:val", this.val);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
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

        private ArrayList itemsField;

        private List<ItemsChoiceType24> itemsElementNameField;

        private string uriField;

        private string elementField;

        public CT_CustomXmlRun()
        {
            this.itemsElementNameField = new List<ItemsChoiceType24>();
            this.itemsField = new ArrayList();
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
        public ArrayList Items
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
        public List<ItemsChoiceType24> ItemsElementName
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

        public static CT_CustomXmlRun Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CustomXmlRun ctObj = new CT_CustomXmlRun();
            ctObj.uri = XmlHelper.ReadString(node.Attributes["w:uri"]);
            ctObj.element = XmlHelper.ReadString(node.Attributes["w:element"]);

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "bookmarkStart")
                {
                    ctObj.Items.Add(CT_Bookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.bookmarkStart);
                }
                else if (childNode.LocalName == "moveFromRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.moveFromRangeStart);
                }
                else if (childNode.LocalName == "moveTo")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.moveTo);
                }
                else if (childNode.LocalName == "oMath")
                {
                    ctObj.Items.Add(CT_OMath.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.oMath);
                }
                else if (childNode.LocalName == "customXmlDelRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.customXmlDelRangeEnd);
                }
                else if (childNode.LocalName == "customXmlDelRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.customXmlDelRangeStart);
                }
                else if (childNode.LocalName == "customXmlInsRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.customXmlInsRangeEnd);
                }
                else if (childNode.LocalName == "customXmlInsRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.customXmlInsRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.customXmlMoveFromRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.customXmlMoveFromRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.customXmlMoveToRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.customXmlMoveToRangeStart);
                }
                else if (childNode.LocalName == "del")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.del);
                }
                else if (childNode.LocalName == "commentRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.commentRangeEnd);
                }
                else if (childNode.LocalName == "fldSimple")
                {
                    ctObj.Items.Add(CT_SimpleField.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.fldSimple);
                }
                else if (childNode.LocalName == "hyperlink")
                {
                    ctObj.Items.Add(CT_Hyperlink1.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.hyperlink);
                }
                else if (childNode.LocalName == "ins")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.ins);
                }
                else if (childNode.LocalName == "moveFrom")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.moveFrom);
                }
                else if (childNode.LocalName == "moveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.moveFromRangeEnd);
                }
                else if (childNode.LocalName == "customXml")
                {
                    ctObj.Items.Add(CT_CustomXmlRun.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.customXml);
                }
                else if (childNode.LocalName == "moveToRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.moveToRangeEnd);
                }
                else if (childNode.LocalName == "moveToRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.moveToRangeStart);
                }
                else if (childNode.LocalName == "permEnd")
                {
                    ctObj.Items.Add(CT_Perm.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.permEnd);
                }
                else if (childNode.LocalName == "permStart")
                {
                    ctObj.Items.Add(CT_PermStart.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.permStart);
                }
                else if (childNode.LocalName == "bookmarkEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.bookmarkEnd);
                }
                else if (childNode.LocalName == "r")
                {
                    ctObj.Items.Add(CT_R.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.r);
                }
                else if (childNode.LocalName == "sdt")
                {
                    ctObj.Items.Add(CT_SdtRun.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.sdt);
                }
                else if (childNode.LocalName == "smartTag")
                {
                    ctObj.Items.Add(CT_SmartTagRun.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.smartTag);
                }
                else if (childNode.LocalName == "subDoc")
                {
                    ctObj.Items.Add(CT_Rel.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.subDoc);
                }
                else if (childNode.LocalName == "proofErr")
                {
                    ctObj.Items.Add(CT_ProofErr.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.proofErr);
                }
                else if (childNode.LocalName == "commentRangeStart")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.commentRangeStart);
                }
                else if (childNode.LocalName == "oMathPara")
                {
                    ctObj.Items.Add(CT_OMathPara.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType24.oMathPara);
                }
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:uri", this.uri);
            XmlHelper.WriteAttribute(sw, "w:element", this.element);
            sw.Write(">");
            foreach (object o in this.Items)
            {
                if (o is CT_Bookmark)
                    ((CT_Bookmark)o).Write(sw, "bookmarkStart");
                else if (o is CT_MoveBookmark)
                    ((CT_MoveBookmark)o).Write(sw, "moveFromRangeStart");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "moveTo");
                else if (o is CT_OMath)
                    ((CT_OMath)o).Write(sw, "oMath");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlDelRangeEnd");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlDelRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlInsRangeEnd");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlInsRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlMoveFromRangeEnd");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveFromRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlMoveToRangeEnd");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveToRangeStart");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "del");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeEnd");
                else if (o is CT_SimpleField)
                    ((CT_SimpleField)o).Write(sw, "fldSimple");
                else if (o is CT_Hyperlink1)
                    ((CT_Hyperlink1)o).Write(sw, "hyperlink");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "ins");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "moveFrom");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "moveFromRangeEnd");
                else if (o is CT_CustomXmlRun)
                    ((CT_CustomXmlRun)o).Write(sw, "customXml");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "moveToRangeEnd");
                else if (o is CT_MoveBookmark)
                    ((CT_MoveBookmark)o).Write(sw, "moveToRangeStart");
                else if (o is CT_Perm)
                    ((CT_Perm)o).Write(sw, "permEnd");
                else if (o is CT_PermStart)
                    ((CT_PermStart)o).Write(sw, "permStart");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "bookmarkEnd");
                else if (o is CT_R)
                    ((CT_R)o).Write(sw, "r");
                else if (o is CT_SdtRun)
                    ((CT_SdtRun)o).Write(sw, "sdt");
                else if (o is CT_SmartTagRun)
                    ((CT_SmartTagRun)o).Write(sw, "smartTag");
                else if (o is CT_Rel)
                    ((CT_Rel)o).Write(sw, "subDoc");
                else if (o is CT_ProofErr)
                    ((CT_ProofErr)o).Write(sw, "proofErr");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeStart");
                else if (o is CT_OMathPara)
                    ((CT_OMathPara)o).Write(sw, "oMathPara");
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
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

        private ArrayList itemsField;

        private List<ItemsChoiceType25> itemsElementNameField;

        private string uriField;

        private string elementField;

        public CT_SmartTagRun()
        {
            this.itemsElementNameField = new List<ItemsChoiceType25>();
            this.itemsField = new ArrayList();
            //this.smartTagPrField = new List<CT_Attr>();
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
        public ArrayList Items
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
        public List<ItemsChoiceType25> ItemsElementName
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

        public static CT_SmartTagRun Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SmartTagRun ctObj = new CT_SmartTagRun();
            ctObj.uri = XmlHelper.ReadString(node.Attributes["w:uri"]);
            ctObj.element = XmlHelper.ReadString(node.Attributes["w:element"]);
            ctObj.smartTagPr = new List<CT_Attr>();

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "customXmlMoveToRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.customXmlMoveToRangeStart);
                }
                else if (childNode.LocalName == "smartTagPr")
                {
                    ctObj.smartTagPr.Add(CT_Attr.Parse(childNode, namespaceManager));
                }
                else if (childNode.LocalName == "fldSimple")
                {
                    ctObj.Items.Add(CT_SimpleField.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.fldSimple);
                }
                else if (childNode.LocalName == "hyperlink")
                {
                    ctObj.Items.Add(CT_Hyperlink1.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.hyperlink);
                }
                else if (childNode.LocalName == "ins")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.ins);
                }
                else if (childNode.LocalName == "moveFrom")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.moveFrom);
                }
                else if (childNode.LocalName == "moveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.moveFromRangeEnd);
                }
                else if (childNode.LocalName == "moveFromRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.moveFromRangeStart);
                }
                else if (childNode.LocalName == "moveTo")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.moveTo);
                }
                else if (childNode.LocalName == "moveToRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.moveToRangeEnd);
                }
                else if (childNode.LocalName == "moveToRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.moveToRangeStart);
                }
                else if (childNode.LocalName == "permEnd")
                {
                    ctObj.Items.Add(CT_Perm.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.permEnd);
                }
                else if (childNode.LocalName == "permStart")
                {
                    ctObj.Items.Add(CT_PermStart.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.permStart);
                }
                else if (childNode.LocalName == "proofErr")
                {
                    ctObj.Items.Add(CT_ProofErr.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.proofErr);
                }
                else if (childNode.LocalName == "r")
                {
                    ctObj.Items.Add(CT_R.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.r);
                }
                else if (childNode.LocalName == "sdt")
                {
                    ctObj.Items.Add(CT_SdtRun.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.sdt);
                }
                else if (childNode.LocalName == "smartTag")
                {
                    ctObj.Items.Add(CT_SmartTagRun.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.smartTag);
                }
                else if (childNode.LocalName == "subDoc")
                {
                    ctObj.Items.Add(CT_Rel.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.subDoc);
                }
                else if (childNode.LocalName == "oMath")
                {
                    ctObj.Items.Add(CT_OMath.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.oMath);
                }
                else if (childNode.LocalName == "del")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.del);
                }
                else if (childNode.LocalName == "oMathPara")
                {
                    ctObj.Items.Add(CT_OMathPara.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.oMathPara);
                }
                else if (childNode.LocalName == "bookmarkEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.bookmarkEnd);
                }
                else if (childNode.LocalName == "bookmarkStart")
                {
                    ctObj.Items.Add(CT_Bookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.bookmarkStart);
                }
                else if (childNode.LocalName == "commentRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.commentRangeEnd);
                }
                else if (childNode.LocalName == "commentRangeStart")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.commentRangeStart);
                }
                else if (childNode.LocalName == "customXml")
                {
                    ctObj.Items.Add(CT_CustomXmlRun.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.customXml);
                }
                else if (childNode.LocalName == "customXmlDelRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.customXmlDelRangeEnd);
                }
                else if (childNode.LocalName == "customXmlDelRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.customXmlDelRangeStart);
                }
                else if (childNode.LocalName == "customXmlInsRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.customXmlInsRangeEnd);
                }
                else if (childNode.LocalName == "customXmlInsRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.customXmlInsRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.customXmlMoveFromRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.customXmlMoveFromRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType25.customXmlMoveToRangeEnd);
                }
            }
            return ctObj;
        }
        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:uri", this.uri);
            XmlHelper.WriteAttribute(sw, "w:element", this.element);
            sw.Write(">");
            if (this.smartTagPr != null)
            {
                foreach (CT_Attr x in this.smartTagPr)
                {
                    x.Write(sw, "smartTagPr");
                }
            }
            foreach (object o in this.Items)
            {
                if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveToRangeStart");
                else if (o is CT_SimpleField)
                    ((CT_SimpleField)o).Write(sw, "fldSimple");
                else if (o is CT_Hyperlink1)
                    ((CT_Hyperlink1)o).Write(sw, "hyperlink");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "ins");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "moveFrom");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "moveFromRangeEnd");
                else if (o is CT_MoveBookmark)
                    ((CT_MoveBookmark)o).Write(sw, "moveFromRangeStart");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "moveTo");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "moveToRangeEnd");
                else if (o is CT_MoveBookmark)
                    ((CT_MoveBookmark)o).Write(sw, "moveToRangeStart");
                else if (o is CT_Perm)
                    ((CT_Perm)o).Write(sw, "permEnd");
                else if (o is CT_PermStart)
                    ((CT_PermStart)o).Write(sw, "permStart");
                else if (o is CT_ProofErr)
                    ((CT_ProofErr)o).Write(sw, "proofErr");
                else if (o is CT_R)
                    ((CT_R)o).Write(sw, "r");
                else if (o is CT_SdtRun)
                    ((CT_SdtRun)o).Write(sw, "sdt");
                else if (o is CT_SmartTagRun)
                    ((CT_SmartTagRun)o).Write(sw, "smartTag");
                else if (o is CT_Rel)
                    ((CT_Rel)o).Write(sw, "subDoc");
                else if (o is CT_OMath)
                    ((CT_OMath)o).Write(sw, "oMath");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "del");
                else if (o is CT_OMathPara)
                    ((CT_OMathPara)o).Write(sw, "oMathPara");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "bookmarkEnd");
                else if (o is CT_Bookmark)
                    ((CT_Bookmark)o).Write(sw, "bookmarkStart");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeEnd");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeStart");
                else if (o is CT_CustomXmlRun)
                    ((CT_CustomXmlRun)o).Write(sw, "customXml");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlDelRangeEnd");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlDelRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlInsRangeEnd");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlInsRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlMoveFromRangeEnd");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveFromRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlMoveToRangeEnd");
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
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

        private ArrayList itemsField;

        private List<ItemsChoiceType26> itemsElementNameField;

        private string uriField;

        private string elementField;

        public CT_CustomXmlBlock()
        {
            this.itemsElementNameField = new List<ItemsChoiceType26>();
            this.itemsField = new ArrayList();
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
        public static CT_CustomXmlBlock Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CustomXmlBlock ctObj = new CT_CustomXmlBlock();
            ctObj.uri = XmlHelper.ReadString(node.Attributes["w:uri"]);
            ctObj.element = XmlHelper.ReadString(node.Attributes["w:element"]);

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "del")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.del);
                }
                else if (childNode.LocalName == "moveFrom")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.moveFrom);
                }
                else if (childNode.LocalName == "moveTo")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.moveTo);
                }
                else if (childNode.LocalName == "bookmarkStart")
                {
                    ctObj.Items.Add(CT_Bookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.bookmarkStart);
                }
                else if (childNode.LocalName == "commentRangeStart")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.commentRangeStart);
                }
                else if (childNode.LocalName == "customXml")
                {
                    ctObj.Items.Add(CT_CustomXmlBlock.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.customXml);
                }
                else if (childNode.LocalName == "customXmlDelRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.customXmlDelRangeEnd);
                }
                else if (childNode.LocalName == "customXmlDelRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.customXmlDelRangeStart);
                }
                else if (childNode.LocalName == "customXmlInsRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.customXmlInsRangeEnd);
                }
                else if (childNode.LocalName == "customXmlInsRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.customXmlInsRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.customXmlMoveFromRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.customXmlMoveFromRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.customXmlMoveToRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.customXmlMoveToRangeStart);
                }
                else if (childNode.LocalName == "oMath")
                {
                    ctObj.Items.Add(CT_OMath.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.oMath);
                }
                else if (childNode.LocalName == "ins")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.ins);
                }
                else if (childNode.LocalName == "moveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.moveFromRangeEnd);
                }
                else if (childNode.LocalName == "moveFromRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.moveFromRangeStart);
                }
                else if (childNode.LocalName == "commentRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.commentRangeEnd);
                }
                else if (childNode.LocalName == "moveToRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.moveToRangeEnd);
                }
                else if (childNode.LocalName == "moveToRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.moveToRangeStart);
                }
                else if (childNode.LocalName == "p")
                {
                    ctObj.Items.Add(CT_P.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.p);
                }
                else if (childNode.LocalName == "permEnd")
                {
                    ctObj.Items.Add(CT_Perm.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.permEnd);
                }
                else if (childNode.LocalName == "bookmarkEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.bookmarkEnd);
                }
                else if (childNode.LocalName == "proofErr")
                {
                    ctObj.Items.Add(CT_ProofErr.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.proofErr);
                }
                else if (childNode.LocalName == "sdt")
                {
                    ctObj.Items.Add(CT_SdtBlock.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.sdt);
                }
                else if (childNode.LocalName == "tbl")
                {
                    ctObj.Items.Add(CT_Tbl.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.tbl);
                }
                else if (childNode.LocalName == "permStart")
                {
                    ctObj.Items.Add(CT_PermStart.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.permStart);
                }
                else if (childNode.LocalName == "oMathPara")
                {
                    ctObj.Items.Add(CT_OMathPara.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType26.oMathPara);
                }
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:uri", this.uri);
            XmlHelper.WriteAttribute(sw, "w:element", this.element);
            sw.Write(">");
            foreach (object o in this.Items)
            {
                if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "del");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "moveFrom");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "moveTo");
                else if (o is CT_Bookmark)
                    ((CT_Bookmark)o).Write(sw, "bookmarkStart");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeStart");
                else if (o is CT_CustomXmlBlock)
                    ((CT_CustomXmlBlock)o).Write(sw, "customXml");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlDelRangeEnd");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlDelRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlInsRangeEnd");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlInsRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlMoveFromRangeEnd");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveFromRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlMoveToRangeEnd");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveToRangeStart");
                else if (o is CT_OMath)
                    ((CT_OMath)o).Write(sw, "oMath");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "ins");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "moveFromRangeEnd");
                else if (o is CT_MoveBookmark)
                    ((CT_MoveBookmark)o).Write(sw, "moveFromRangeStart");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeEnd");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "moveToRangeEnd");
                else if (o is CT_MoveBookmark)
                    ((CT_MoveBookmark)o).Write(sw, "moveToRangeStart");
                else if (o is CT_P)
                    ((CT_P)o).Write(sw, "p");
                else if (o is CT_Perm)
                    ((CT_Perm)o).Write(sw, "permEnd");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "bookmarkEnd");
                else if (o is CT_ProofErr)
                    ((CT_ProofErr)o).Write(sw, "proofErr");
                else if (o is CT_SdtBlock)
                    ((CT_SdtBlock)o).Write(sw, "sdt");
                else if (o is CT_Tbl)
                    ((CT_Tbl)o).Write(sw, "tbl");
                else if (o is CT_PermStart)
                    ((CT_PermStart)o).Write(sw, "permStart");
                else if (o is CT_OMathPara)
                    ((CT_OMathPara)o).Write(sw, "oMathPara");
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
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
        public ArrayList Items
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
        public List<ItemsChoiceType26> ItemsElementName
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

        private ArrayList itemsField;

        private List<ItemsChoiceType27> itemsElementNameField;

        private string uriField;

        private string elementField;

        public CT_CustomXmlCell()
        {
            this.itemsElementNameField = new List<ItemsChoiceType27>();
            this.itemsField = new ArrayList();
            //this.customXmlPrField = new CT_CustomXmlPr();
        }
        public static CT_CustomXmlCell Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CustomXmlCell ctObj = new CT_CustomXmlCell();
            ctObj.uri = XmlHelper.ReadString(node.Attributes["w:uri"]);
            ctObj.element = XmlHelper.ReadString(node.Attributes["w:element"]);

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "customXmlMoveToRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.customXmlMoveToRangeStart);
                }
                else if (childNode.LocalName == "customXmlPr")
                {
                    ctObj.customXmlPr = CT_CustomXmlPr.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.customXmlMoveFromRangeEnd);
                }
                else if (childNode.LocalName == "sdt")
                {
                    ctObj.Items.Add(CT_SdtCell.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.sdt);
                }
                else if (childNode.LocalName == "tc")
                {
                    ctObj.Items.Add(CT_Tc.Parse(childNode, namespaceManager, ctObj));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.tc);
                }
                else if (childNode.LocalName == "permEnd")
                {
                    ctObj.Items.Add(CT_Perm.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.permEnd);
                }
                else if (childNode.LocalName == "customXmlInsRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.customXmlInsRangeStart);
                }
                else if (childNode.LocalName == "del")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.del);
                }
                else if (childNode.LocalName == "ins")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.ins);
                }
                else if (childNode.LocalName == "moveFrom")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.moveFrom);
                }
                else if (childNode.LocalName == "moveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.moveFromRangeEnd);
                }
                else if (childNode.LocalName == "moveFromRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.moveFromRangeStart);
                }
                else if (childNode.LocalName == "moveTo")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.moveTo);
                }
                else if (childNode.LocalName == "moveToRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.moveToRangeEnd);
                }
                else if (childNode.LocalName == "moveToRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.moveToRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.customXmlMoveToRangeEnd);
                }
                else if (childNode.LocalName == "permStart")
                {
                    ctObj.Items.Add(CT_PermStart.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.permStart);
                }
                else if (childNode.LocalName == "proofErr")
                {
                    ctObj.Items.Add(CT_ProofErr.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.proofErr);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.customXmlMoveFromRangeStart);
                }
                else if (childNode.LocalName == "oMath")
                {
                    ctObj.Items.Add(CT_OMath.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.oMath);
                }
                else if (childNode.LocalName == "oMathPara")
                {
                    ctObj.Items.Add(CT_OMathPara.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.oMathPara);
                }
                else if (childNode.LocalName == "bookmarkEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.bookmarkEnd);
                }
                else if (childNode.LocalName == "bookmarkStart")
                {
                    ctObj.Items.Add(CT_Bookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.bookmarkStart);
                }
                else if (childNode.LocalName == "commentRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.commentRangeEnd);
                }
                else if (childNode.LocalName == "commentRangeStart")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.commentRangeStart);
                }
                else if (childNode.LocalName == "customXml")
                {
                    ctObj.Items.Add(CT_CustomXmlCell.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.customXml);
                }
                else if (childNode.LocalName == "customXmlDelRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.customXmlDelRangeEnd);
                }
                else if (childNode.LocalName == "customXmlDelRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.customXmlDelRangeStart);
                }
                else if (childNode.LocalName == "customXmlInsRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType27.customXmlInsRangeEnd);
                }
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:uri", this.uri);
            XmlHelper.WriteAttribute(sw, "w:element", this.element);
            sw.Write(">");
            if (this.customXmlPr != null)
                this.customXmlPr.Write(sw, "customXmlPr");
            foreach (object o in this.Items)
            {
                if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveToRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlMoveFromRangeEnd");
                else if (o is CT_SdtCell)
                    ((CT_SdtCell)o).Write(sw, "sdt");
                else if (o is CT_Tc)
                    ((CT_Tc)o).Write(sw, "tc");
                else if (o is CT_Perm)
                    ((CT_Perm)o).Write(sw, "permEnd");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlInsRangeStart");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "del");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "ins");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "moveFrom");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "moveFromRangeEnd");
                else if (o is CT_MoveBookmark)
                    ((CT_MoveBookmark)o).Write(sw, "moveFromRangeStart");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "moveTo");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "moveToRangeEnd");
                else if (o is CT_MoveBookmark)
                    ((CT_MoveBookmark)o).Write(sw, "moveToRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlMoveToRangeEnd");
                else if (o is CT_PermStart)
                    ((CT_PermStart)o).Write(sw, "permStart");
                else if (o is CT_ProofErr)
                    ((CT_ProofErr)o).Write(sw, "proofErr");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveFromRangeStart");
                else if (o is CT_OMath)
                    ((CT_OMath)o).Write(sw, "oMath");
                else if (o is CT_OMathPara)
                    ((CT_OMathPara)o).Write(sw, "oMathPara");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "bookmarkEnd");
                else if (o is CT_Bookmark)
                    ((CT_Bookmark)o).Write(sw, "bookmarkStart");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeEnd");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeStart");
                else if (o is CT_CustomXmlCell)
                    ((CT_CustomXmlCell)o).Write(sw, "customXml");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlDelRangeEnd");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlDelRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlInsRangeEnd");
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
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
        public ArrayList Items
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
        public List<ItemsChoiceType27> ItemsElementName
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