using System;
using System.Collections.Generic;
using System.Xml.Serialization;
using NPOI.OpenXmlFormats.Shared;
using System.Xml;
using NPOI.OpenXml4Net.Util;
using System.IO;
using System.Collections;

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
        public static new CT_Bookmark Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Bookmark ctObj = new CT_Bookmark();
            ctObj.name = XmlHelper.ReadString(node.Attributes["w:name"]);
            ctObj.colFirst = XmlHelper.ReadString(node.Attributes["w:colFirst"]);
            ctObj.colLast = XmlHelper.ReadString(node.Attributes["w:colLast"]);
            if (node.Attributes["w:displacedByCustomXml"] != null)
                ctObj.displacedByCustomXml = (ST_DisplacedByCustomXml)Enum.Parse(typeof(ST_DisplacedByCustomXml), node.Attributes["w:displacedByCustomXml"].Value);
            ctObj.id = XmlHelper.ReadString(node.Attributes["w:id"]);
            return ctObj;
        }



        internal new void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:id", this.id);
            XmlHelper.WriteAttribute(sw, "w:name", this.name);
            XmlHelper.WriteAttribute(sw, "w:colFirst", this.colFirst);
            XmlHelper.WriteAttribute(sw, "w:colLast", this.colLast);
            if (this.displacedByCustomXml!= ST_DisplacedByCustomXml.next)
                XmlHelper.WriteAttribute(sw, "w:displacedByCustomXml", this.displacedByCustomXml.ToString());
            sw.Write("/>");
        }

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
        public static new CT_MoveBookmark Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_MoveBookmark ctObj = new CT_MoveBookmark();
            ctObj.author = XmlHelper.ReadString(node.Attributes["w:author"]);
            ctObj.date = XmlHelper.ReadString(node.Attributes["w:date"]);
            ctObj.name = XmlHelper.ReadString(node.Attributes["w:name"]);
            ctObj.colFirst = XmlHelper.ReadString(node.Attributes["w:colFirst"]);
            ctObj.colLast = XmlHelper.ReadString(node.Attributes["w:colLast"]);
            if (node.Attributes["w:displacedByCustomXml"] != null)
                ctObj.displacedByCustomXml = (ST_DisplacedByCustomXml)Enum.Parse(typeof(ST_DisplacedByCustomXml), node.Attributes["w:displacedByCustomXml"].Value);
            ctObj.id = XmlHelper.ReadString(node.Attributes["w:id"]);
            return ctObj;
        }



        internal new void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:author", this.author);
            XmlHelper.WriteAttribute(sw, "w:date", this.date);
            XmlHelper.WriteAttribute(sw, "w:name", this.name);
            XmlHelper.WriteAttribute(sw, "w:colFirst", this.colFirst);
            XmlHelper.WriteAttribute(sw, "w:colLast", this.colLast);
            XmlHelper.WriteAttribute(sw, "w:displacedByCustomXml", this.displacedByCustomXml.ToString());
            XmlHelper.WriteAttribute(sw, "w:id", this.id);
            sw.Write("/>");
        }

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
        public static CT_Comments Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Comments ctObj = new CT_Comments();
            ctObj.comment = new List<CT_Comment>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "comment")
                    ctObj.comment.Add(CT_Comment.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw)
        {
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write(string.Format("<w:comments "));
            sw.Write("xmlns:ve=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" ");
            sw.Write("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" ");
            sw.Write("xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" ");
            sw.Write("xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" ");
            sw.Write("xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\">");
            sw.Write(">");
            if (this.comment != null)
            {
                foreach (CT_Comment x in this.comment)
                {
                    x.Write(sw, "comment");
                }
            }
            sw.Write("</w:comments>");
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

        private ArrayList itemsField;

        private List<ItemsChoiceType50> itemsElementNameField;

        private string initialsField;

        public CT_Comment()
        {
            this.itemsElementNameField = new List<ItemsChoiceType50>();
            this.itemsField = new ArrayList();
        }
        public static new CT_Comment Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Comment ctObj = new CT_Comment();
            ctObj.initials = XmlHelper.ReadString(node.Attributes["w:initials"]);
            ctObj.author = XmlHelper.ReadString(node.Attributes["w:author"]);
            ctObj.date = XmlHelper.ReadString(node.Attributes["w:date"]);
            ctObj.id = XmlHelper.ReadString(node.Attributes["w:id"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "permStart")
                {
                    ctObj.Items.Add(CT_PermStart.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.permStart);
                }
                else if (childNode.LocalName == "customXmlInsRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.customXmlInsRangeEnd);
                }
                else if (childNode.LocalName == "customXmlInsRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.customXmlInsRangeStart);
                }
                else if (childNode.LocalName == "permEnd")
                {
                    ctObj.Items.Add(CT_Perm.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.permEnd);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.customXmlMoveFromRangeEnd);
                }
                else if (childNode.LocalName == "proofErr")
                {
                    ctObj.Items.Add(CT_ProofErr.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.proofErr);
                }
                else if (childNode.LocalName == "sdt")
                {
                    ctObj.Items.Add(CT_SdtBlock.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.sdt);
                }
                else if (childNode.LocalName == "tbl")
                {
                    ctObj.Items.Add(CT_Tbl.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.tbl);
                }
                else if (childNode.LocalName == "moveFromRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.moveFromRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.customXmlMoveToRangeEnd);
                }
                else if (childNode.LocalName == "oMath")
                {
                    ctObj.Items.Add(CT_OMath.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.oMath);
                }
                else if (childNode.LocalName == "oMathPara")
                {
                    ctObj.Items.Add(CT_OMathPara.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.oMathPara);
                }
                else if (childNode.LocalName == "altChunk")
                {
                    ctObj.Items.Add(CT_AltChunk.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.altChunk);
                }
                else if (childNode.LocalName == "bookmarkEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.bookmarkEnd);
                }
                else if (childNode.LocalName == "bookmarkStart")
                {
                    ctObj.Items.Add(CT_Bookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.bookmarkStart);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.customXmlMoveToRangeStart);
                }
                else if (childNode.LocalName == "commentRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.commentRangeEnd);
                }
                else if (childNode.LocalName == "commentRangeStart")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.commentRangeStart);
                }
                else if (childNode.LocalName == "moveTo")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.moveTo);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.customXmlMoveFromRangeStart);
                }
                else if (childNode.LocalName == "moveToRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.moveToRangeEnd);
                }
                else if (childNode.LocalName == "customXml")
                {
                    ctObj.Items.Add(CT_CustomXmlBlock.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.customXml);
                }
                else if (childNode.LocalName == "moveToRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.moveToRangeStart);
                }
                else if (childNode.LocalName == "p")
                {
                    ctObj.Items.Add(CT_P.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.p);
                }
                else if (childNode.LocalName == "customXmlDelRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.customXmlDelRangeEnd);
                }
                else if (childNode.LocalName == "del")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.del);
                }
                else if (childNode.LocalName == "ins")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.ins);
                }
                else if (childNode.LocalName == "moveFrom")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.moveFrom);
                }
                else if (childNode.LocalName == "customXmlDelRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.customXmlDelRangeStart);
                }
                else if (childNode.LocalName == "moveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType50.moveFromRangeEnd);
                }
            }
            return ctObj;
        }

        internal new void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:initials", this.initials);
            XmlHelper.WriteAttribute(sw, "w:author", this.author);
            XmlHelper.WriteAttribute(sw, "w:date", this.date);
            XmlHelper.WriteAttribute(sw, "w:id", this.id);
            sw.Write(">");
            foreach (object o in this.Items)
            {
                if (o is CT_PermStart)
                    ((CT_PermStart)o).Write(sw, "permStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlInsRangeEnd");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlInsRangeStart");
                else if (o is CT_Perm)
                    ((CT_Perm)o).Write(sw, "permEnd");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlMoveFromRangeEnd");
                else if (o is CT_ProofErr)
                    ((CT_ProofErr)o).Write(sw, "proofErr");
                else if (o is CT_SdtBlock)
                    ((CT_SdtBlock)o).Write(sw, "sdt");
                else if (o is CT_Tbl)
                    ((CT_Tbl)o).Write(sw, "tbl");
                else if (o is CT_MoveBookmark)
                    ((CT_MoveBookmark)o).Write(sw, "moveFromRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlMoveToRangeEnd");
                else if (o is CT_OMath)
                    ((CT_OMath)o).Write(sw, "oMath");
                else if (o is CT_OMathPara)
                    ((CT_OMathPara)o).Write(sw, "oMathPara");
                else if (o is CT_AltChunk)
                    ((CT_AltChunk)o).Write(sw, "altChunk");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "bookmarkEnd");
                else if (o is CT_Bookmark)
                    ((CT_Bookmark)o).Write(sw, "bookmarkStart");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveToRangeStart");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeEnd");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeStart");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "moveTo");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveFromRangeStart");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "moveToRangeEnd");
                else if (o is CT_CustomXmlBlock)
                    ((CT_CustomXmlBlock)o).Write(sw, "customXml");
                else if (o is CT_MoveBookmark)
                    ((CT_MoveBookmark)o).Write(sw, "moveToRangeStart");
                else if (o is CT_P)
                    ((CT_P)o).Write(sw, "p");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlDelRangeEnd");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "del");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "ins");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "moveFrom");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlDelRangeStart");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "moveFromRangeEnd");
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [System.Xml.Serialization.XmlElementAttribute("oMath", typeof(CT_OMath), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("oMathPara", typeof(CT_OMathPara), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("altChunk", typeof(CT_AltChunk), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("bookmarkEnd", typeof(CT_MarkupRange), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("bookmarkStart", typeof(CT_Bookmark), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("commentRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("commentRangeStart", typeof(CT_MarkupRange), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("customXml", typeof(CT_CustomXmlBlock), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("customXmlDelRangeEnd", typeof(CT_Markup), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("customXmlDelRangeStart", typeof(CT_TrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("customXmlInsRangeEnd", typeof(CT_Markup), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("customXmlInsRangeStart", typeof(CT_TrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("customXmlMoveFromRangeEnd", typeof(CT_Markup), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("customXmlMoveFromRangeStart", typeof(CT_TrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("customXmlMoveToRangeEnd", typeof(CT_Markup), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("customXmlMoveToRangeStart", typeof(CT_TrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("del", typeof(CT_RunTrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("ins", typeof(CT_RunTrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("moveFrom", typeof(CT_RunTrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("moveFromRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("moveFromRangeStart", typeof(CT_MoveBookmark), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("moveTo", typeof(CT_RunTrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("moveToRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("moveToRangeStart", typeof(CT_MoveBookmark), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("p", typeof(CT_P), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("permEnd", typeof(CT_Perm), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("permStart", typeof(CT_PermStart), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("proofErr", typeof(CT_ProofErr), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("sdt", typeof(CT_SdtBlock), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("tbl", typeof(CT_Tbl), Order = 0)]
        [System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemsElementName")]
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

        [System.Xml.Serialization.XmlElementAttribute("ItemsElementName", Order = 1)]
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public List<ItemsChoiceType50> ItemsElementName
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

        public IList<CT_P> GetPList()
        {
            return GetObjectList<CT_P>(ItemsChoiceType50.p);
        }
        #region Generic methods for object operation

        private List<T> GetObjectList<T>(ItemsChoiceType50 type) where T : class
        {
            lock (this)
            {
                List<T> list = new List<T>();
                for (int i = 0; i < itemsElementNameField.Count; i++)
                {
                    if (itemsElementNameField[i] == type)
                        list.Add(itemsField[i] as T);
                }
                return list;
            }
        }
        private int SizeOfArray(ItemsChoiceType50 type)
        {
            lock (this)
            {
                int size = 0;
                for (int i = 0; i < itemsElementNameField.Count; i++)
                {
                    if (itemsElementNameField[i] == type)
                        size++;
                }
                return size;
            }
        }
        private T GetObjectArray<T>(int p, ItemsChoiceType50 type) where T : class
        {
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                if (pos < 0 || pos >= this.itemsField.Count)
                    return null;
                return itemsField[pos] as T;
            }
        }
        private T InsertNewObject<T>(ItemsChoiceType50 type, int p) where T : class, new()
        {
            T t = new T();
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                this.itemsElementNameField.Insert(pos, type);
                this.itemsField.Insert(pos, t);
            }
            return t;
        }
        private T AddNewObject<T>(ItemsChoiceType50 type) where T : class, new()
        {
            T t = new T();
            lock (this)
            {
                this.itemsElementNameField.Add(type);
                this.itemsField.Add(t);
            }
            return t;
        }
        private void SetObject<T>(ItemsChoiceType50 type, int p, T obj) where T : class
        {
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                if (pos < 0 || pos >= this.itemsField.Count)
                    return;
                if (this.itemsField[pos] is T)
                    this.itemsField[pos] = obj;
                else
                    throw new Exception(string.Format(@"object types are difference, itemsField[{0}] is {1}, and parameter obj is {2}",
                        pos, this.itemsField[pos].GetType().Name, typeof(T).Name));
            }
        }
        private int GetObjectIndex(ItemsChoiceType50 type, int p)
        {
            int index = -1;
            int pos = 0;
            for (int i = 0; i < itemsElementNameField.Count; i++)
            {
                if (itemsElementNameField[i] == type)
                {
                    if (pos == p)
                    {
                        index = i;
                        break;
                    }
                    else
                        pos++;
                }
            }
            return index;
        }
        private void RemoveObject(ItemsChoiceType50 type, int p)
        {
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                if (pos < 0 || pos >= this.itemsField.Count)
                    return;
                itemsElementNameField.RemoveAt(pos);
                itemsField.RemoveAt(pos);
            }
        }
        #endregion
    }
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType50
    {

        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:oMath")]
        oMath,

        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:oMathPara")]
        oMathPara,

        /// <remarks/>
        altChunk,

        /// <remarks/>
        bookmarkEnd,

        /// <remarks/>
        bookmarkStart,

        /// <remarks/>
        commentRangeEnd,

        /// <remarks/>
        commentRangeStart,

        /// <remarks/>
        customXml,

        /// <remarks/>
        customXmlDelRangeEnd,

        /// <remarks/>
        customXmlDelRangeStart,

        /// <remarks/>
        customXmlInsRangeEnd,

        /// <remarks/>
        customXmlInsRangeStart,

        /// <remarks/>
        customXmlMoveFromRangeEnd,

        /// <remarks/>
        customXmlMoveFromRangeStart,

        /// <remarks/>
        customXmlMoveToRangeEnd,

        /// <remarks/>
        customXmlMoveToRangeStart,

        /// <remarks/>
        del,

        /// <remarks/>
        ins,

        /// <remarks/>
        moveFrom,

        /// <remarks/>
        moveFromRangeEnd,

        /// <remarks/>
        moveFromRangeStart,

        /// <remarks/>
        moveTo,

        /// <remarks/>
        moveToRangeEnd,

        /// <remarks/>
        moveToRangeStart,

        /// <remarks/>
        p,

        /// <remarks/>
        permEnd,

        /// <remarks/>
        permStart,

        /// <remarks/>
        proofErr,

        /// <remarks/>
        sdt,

        /// <remarks/>
        tbl,
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
        public new static CT_TrackChange Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TrackChange ctObj = new CT_TrackChange();
            ctObj.author = XmlHelper.ReadString(node.Attributes["w:author"]);
            ctObj.date = XmlHelper.ReadString(node.Attributes["w:date"]);
            ctObj.id = XmlHelper.ReadString(node.Attributes["w:id"]);
            return ctObj;
        }



        internal new void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:author", this.author, true);
            XmlHelper.WriteAttribute(sw, "w:date", this.date);
            XmlHelper.WriteAttribute(sw, "w:id", this.id, true);
            sw.Write("/>");
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
            ctObj.id = XmlHelper.ReadString(node.Attributes["w:id"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:id", this.id);
            sw.Write("/>");
        }

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
        public static new CT_MarkupRange Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_MarkupRange ctObj = new CT_MarkupRange();
            if (node.Attributes["w:displacedByCustomXml"] != null)
                ctObj.displacedByCustomXml = (ST_DisplacedByCustomXml)Enum.Parse(typeof(ST_DisplacedByCustomXml), node.Attributes["w:displacedByCustomXml"].Value);
            ctObj.id = XmlHelper.ReadString(node.Attributes["w:id"]);
            return ctObj;
        }



        internal new void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            if (this.displacedByCustomXml!= ST_DisplacedByCustomXml.next)
                XmlHelper.WriteAttribute(sw, "w:displacedByCustomXml", this.displacedByCustomXml.ToString());
            XmlHelper.WriteAttribute(sw, "w:id", this.id);
            sw.Write("/>");
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
