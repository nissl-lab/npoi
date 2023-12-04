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
            this.commentField = new List<CT_Comment>();
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

        public CT_Comment AddNewComment()
        {
            var comment = new CT_Comment();
            commentField.Add(comment);
            return comment;
        }

        public void RemoveComment(int i)
        {
            commentField.RemoveAt(i);
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
            sw.WriteEndW(nodeName);
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

        #pragma warning disable format
        public CT_AltChunk       AddNewAltChunk()                    { return AddNewObject<CT_AltChunk>      (ItemsChoiceType50.altChunk); }
        public CT_MarkupRange    AddNewBookmarkEnd()                 { return AddNewObject<CT_MarkupRange>   (ItemsChoiceType50.bookmarkEnd); }
        public CT_Bookmark       AddNewBookmarkStart()               { return AddNewObject<CT_Bookmark>      (ItemsChoiceType50.bookmarkStart); }
        public CT_MarkupRange    AddNewCommentRangeEnd()             { return AddNewObject<CT_MarkupRange>   (ItemsChoiceType50.commentRangeEnd); }
        public CT_MarkupRange    AddNewCommentRangeStart()           { return AddNewObject<CT_MarkupRange>   (ItemsChoiceType50.commentRangeStart); }
        public CT_CustomXmlBlock AddNewCustomXml()                   { return AddNewObject<CT_CustomXmlBlock>(ItemsChoiceType50.customXml); }
        public CT_Markup         AddNewCustomXmlDelRangeEnd()        { return AddNewObject<CT_Markup>        (ItemsChoiceType50.customXmlDelRangeEnd); }
        public CT_TrackChange    AddNewCustomXmlDelRangeStart()      { return AddNewObject<CT_TrackChange>   (ItemsChoiceType50.customXmlDelRangeStart); }
        public CT_Markup         AddNewCustomXmlInsRangeEnd()        { return AddNewObject<CT_Markup>        (ItemsChoiceType50.customXmlInsRangeEnd); }
        public CT_TrackChange    AddNewCustomXmlInsRangeStart()      { return AddNewObject<CT_TrackChange>   (ItemsChoiceType50.customXmlInsRangeStart); }
        public CT_Markup         AddNewCustomXmlMoveFromRangeEnd()   { return AddNewObject<CT_Markup>        (ItemsChoiceType50.customXmlMoveFromRangeEnd); }
        public CT_TrackChange    AddNewCustomXmlMoveFromRangeStart() { return AddNewObject<CT_TrackChange>   (ItemsChoiceType50.customXmlMoveFromRangeStart); }
        public CT_Markup         AddNewCustomXmlMoveToRangeEnd()     { return AddNewObject<CT_Markup>        (ItemsChoiceType50.customXmlMoveToRangeEnd); }
        public CT_TrackChange    AddNewCustomXmlMoveToRangeStart()   { return AddNewObject<CT_TrackChange>   (ItemsChoiceType50.customXmlMoveToRangeStart); }
        public CT_RunTrackChange AddNewDel()                         { return AddNewObject<CT_RunTrackChange>(ItemsChoiceType50.del); }
        public CT_RunTrackChange AddNewIns()                         { return AddNewObject<CT_RunTrackChange>(ItemsChoiceType50.ins); }
        public CT_RunTrackChange AddNewMoveFrom()                    { return AddNewObject<CT_RunTrackChange>(ItemsChoiceType50.moveFrom); }
        public CT_MarkupRange    AddNewMoveFromRangeEnd()            { return AddNewObject<CT_MarkupRange>   (ItemsChoiceType50.moveFromRangeEnd); }
        public CT_MoveBookmark   AddNewMoveFromRangeStart()          { return AddNewObject<CT_MoveBookmark>  (ItemsChoiceType50.moveFromRangeStart); }
        public CT_RunTrackChange AddNewMoveTo()                      { return AddNewObject<CT_RunTrackChange>(ItemsChoiceType50.moveTo); }
        public CT_MarkupRange    AddNewMoveToRangeEnd()              { return AddNewObject<CT_MarkupRange>   (ItemsChoiceType50.moveToRangeEnd); }
        public CT_MoveBookmark   AddNewMoveToRangeStart()            { return AddNewObject<CT_MoveBookmark>  (ItemsChoiceType50.moveToRangeStart); }
        public CT_OMath          AddNewOMath()                       { return AddNewObject<CT_OMath>         (ItemsChoiceType50.oMath); }
        public CT_OMathPara      AddNewOMathPara()                   { return AddNewObject<CT_OMathPara>     (ItemsChoiceType50.oMathPara); }
        public CT_P              AddNewP()                           { return AddNewObject<CT_P>             (ItemsChoiceType50.p); }
        public CT_Perm           AddNewPermEnd()                     { return AddNewObject<CT_Perm>          (ItemsChoiceType50.permEnd); }
        public CT_PermStart      AddNewPermStart()                   { return AddNewObject<CT_PermStart>     (ItemsChoiceType50.permStart); }
        public CT_ProofErr       AddNewProofErr()                    { return AddNewObject<CT_ProofErr>      (ItemsChoiceType50.proofErr); }
        public CT_SdtBlock       AddNewSdt()                         { return AddNewObject<CT_SdtBlock>      (ItemsChoiceType50.sdt); }
        public CT_Tbl            AddNewTbl()                         { return AddNewObject<CT_Tbl>           (ItemsChoiceType50.tbl); }

        public CT_AltChunk       GetAltChunkArray(int p)                    { return GetObjectArray<CT_AltChunk>      (p, ItemsChoiceType50.altChunk); }
        public CT_MarkupRange    GetBookmarkEndArray(int p)                 { return GetObjectArray<CT_MarkupRange>   (p, ItemsChoiceType50.bookmarkEnd); }
        public CT_Bookmark       GetBookmarkStartArray(int p)               { return GetObjectArray<CT_Bookmark>      (p, ItemsChoiceType50.bookmarkStart); }
        public CT_MarkupRange    GetCommentRangeEndArray(int p)             { return GetObjectArray<CT_MarkupRange>   (p, ItemsChoiceType50.commentRangeEnd); }
        public CT_MarkupRange    GetCommentRangeStartArray(int p)           { return GetObjectArray<CT_MarkupRange>   (p, ItemsChoiceType50.commentRangeStart); }
        public CT_CustomXmlBlock GetCustomXmlArray(int p)                   { return GetObjectArray<CT_CustomXmlBlock>(p, ItemsChoiceType50.customXml); }
        public CT_Markup         GetCustomXmlDelRangeEndArray(int p)        { return GetObjectArray<CT_Markup>        (p, ItemsChoiceType50.customXmlDelRangeEnd); }
        public CT_TrackChange    GetCustomXmlDelRangeStartArray(int p)      { return GetObjectArray<CT_TrackChange>   (p, ItemsChoiceType50.customXmlDelRangeStart); }
        public CT_Markup         GetCustomXmlInsRangeEndArray(int p)        { return GetObjectArray<CT_Markup>        (p, ItemsChoiceType50.customXmlInsRangeEnd); }
        public CT_TrackChange    GetCustomXmlInsRangeStartArray(int p)      { return GetObjectArray<CT_TrackChange>   (p, ItemsChoiceType50.customXmlInsRangeStart); }
        public CT_Markup         GetCustomXmlMoveFromRangeEndArray(int p)   { return GetObjectArray<CT_Markup>        (p, ItemsChoiceType50.customXmlMoveFromRangeEnd); }
        public CT_TrackChange    GetCustomXmlMoveFromRangeStartArray(int p) { return GetObjectArray<CT_TrackChange>   (p, ItemsChoiceType50.customXmlMoveFromRangeStart); }
        public CT_Markup         GetCustomXmlMoveToRangeEndArray(int p)     { return GetObjectArray<CT_Markup>        (p, ItemsChoiceType50.customXmlMoveToRangeEnd); }
        public CT_TrackChange    GetCustomXmlMoveToRangeStartArray(int p)   { return GetObjectArray<CT_TrackChange>   (p, ItemsChoiceType50.customXmlMoveToRangeStart); }
        public CT_RunTrackChange GetDelArray(int p)                         { return GetObjectArray<CT_RunTrackChange>(p, ItemsChoiceType50.del); }
        public CT_RunTrackChange GetInsArray(int p)                         { return GetObjectArray<CT_RunTrackChange>(p, ItemsChoiceType50.ins); }
        public CT_RunTrackChange GetMoveFromArray(int p)                    { return GetObjectArray<CT_RunTrackChange>(p, ItemsChoiceType50.moveFrom); }
        public CT_MarkupRange    GetMoveFromRangeEndArray(int p)            { return GetObjectArray<CT_MarkupRange>   (p, ItemsChoiceType50.moveFromRangeEnd); }
        public CT_MoveBookmark   GetMoveFromRangeStartArray(int p)          { return GetObjectArray<CT_MoveBookmark>  (p, ItemsChoiceType50.moveFromRangeStart); }
        public CT_RunTrackChange GetMoveToArray(int p)                      { return GetObjectArray<CT_RunTrackChange>(p, ItemsChoiceType50.moveTo); }
        public CT_MarkupRange    GetMoveToRangeEndArray(int p)              { return GetObjectArray<CT_MarkupRange>   (p, ItemsChoiceType50.moveToRangeEnd); }
        public CT_MoveBookmark   GetMoveToRangeStartArray(int p)            { return GetObjectArray<CT_MoveBookmark>  (p, ItemsChoiceType50.moveToRangeStart); }
        public CT_OMath          GetOMathArray(int p)                       { return GetObjectArray<CT_OMath>         (p, ItemsChoiceType50.oMath); }
        public CT_OMathPara      GetOMathParaArray(int p)                   { return GetObjectArray<CT_OMathPara>     (p, ItemsChoiceType50.oMathPara); }
        public CT_P              GetPArray(int p)                           { return GetObjectArray<CT_P>             (p, ItemsChoiceType50.p); }
        public CT_Perm           GetPermEndArray(int p)                     { return GetObjectArray<CT_Perm>          (p, ItemsChoiceType50.permEnd); }
        public CT_PermStart      GetPermStartArray(int p)                   { return GetObjectArray<CT_PermStart>     (p, ItemsChoiceType50.permStart); }
        public CT_ProofErr       GetProofErrArray(int p)                    { return GetObjectArray<CT_ProofErr>      (p, ItemsChoiceType50.proofErr); }
        public CT_SdtBlock       GetSdtArray(int p)                         { return GetObjectArray<CT_SdtBlock>      (p, ItemsChoiceType50.sdt); }
        public CT_Tbl            GetTblArray(int p)                         { return GetObjectArray<CT_Tbl>           (p, ItemsChoiceType50.tbl); }

        public IList<CT_AltChunk>       GetAltChunkList()                    { return GetObjectList<CT_AltChunk>      (ItemsChoiceType50.altChunk); }
        public IList<CT_MarkupRange>    GetBookmarkEndList()                 { return GetObjectList<CT_MarkupRange>   (ItemsChoiceType50.bookmarkEnd); }
        public IList<CT_Bookmark>       GetBookmarkStartList()               { return GetObjectList<CT_Bookmark>      (ItemsChoiceType50.bookmarkStart); }
        public IList<CT_MarkupRange>    GetCommentRangeEndList()             { return GetObjectList<CT_MarkupRange>   (ItemsChoiceType50.commentRangeEnd); }
        public IList<CT_MarkupRange>    GetCommentRangeStartList()           { return GetObjectList<CT_MarkupRange>   (ItemsChoiceType50.commentRangeStart); }
        public IList<CT_CustomXmlBlock> GetCustomXmlList()                   { return GetObjectList<CT_CustomXmlBlock>(ItemsChoiceType50.customXml); }
        public IList<CT_Markup>         GetCustomXmlDelRangeEndList()        { return GetObjectList<CT_Markup>        (ItemsChoiceType50.customXmlDelRangeEnd); }
        public IList<CT_TrackChange>    GetCustomXmlDelRangeStartList()      { return GetObjectList<CT_TrackChange>   (ItemsChoiceType50.customXmlDelRangeStart); }
        public IList<CT_Markup>         GetCustomXmlInsRangeEndList()        { return GetObjectList<CT_Markup>        (ItemsChoiceType50.customXmlInsRangeEnd); }
        public IList<CT_TrackChange>    GetCustomXmlInsRangeStartList()      { return GetObjectList<CT_TrackChange>   (ItemsChoiceType50.customXmlInsRangeStart); }
        public IList<CT_Markup>         GetCustomXmlMoveFromRangeEndList()   { return GetObjectList<CT_Markup>        (ItemsChoiceType50.customXmlMoveFromRangeEnd); }
        public IList<CT_TrackChange>    GetCustomXmlMoveFromRangeStartList() { return GetObjectList<CT_TrackChange>   (ItemsChoiceType50.customXmlMoveFromRangeStart); }
        public IList<CT_Markup>         GetCustomXmlMoveToRangeEndList()     { return GetObjectList<CT_Markup>        (ItemsChoiceType50.customXmlMoveToRangeEnd); }
        public IList<CT_TrackChange>    GetCustomXmlMoveToRangeStartList()   { return GetObjectList<CT_TrackChange>   (ItemsChoiceType50.customXmlMoveToRangeStart); }
        public IList<CT_RunTrackChange> GetDelList()                         { return GetObjectList<CT_RunTrackChange>(ItemsChoiceType50.del); }
        public IList<CT_RunTrackChange> GetInsList()                         { return GetObjectList<CT_RunTrackChange>(ItemsChoiceType50.ins); }
        public IList<CT_RunTrackChange> GetMoveFromList()                    { return GetObjectList<CT_RunTrackChange>(ItemsChoiceType50.moveFrom); }
        public IList<CT_MarkupRange>    GetMoveFromRangeEndList()            { return GetObjectList<CT_MarkupRange>   (ItemsChoiceType50.moveFromRangeEnd); }
        public IList<CT_MoveBookmark>   GetMoveFromRangeStartList()          { return GetObjectList<CT_MoveBookmark>  (ItemsChoiceType50.moveFromRangeStart); }
        public IList<CT_RunTrackChange> GetMoveToList()                      { return GetObjectList<CT_RunTrackChange>(ItemsChoiceType50.moveTo); }
        public IList<CT_MarkupRange>    GetMoveToRangeEndList()              { return GetObjectList<CT_MarkupRange>   (ItemsChoiceType50.moveToRangeEnd); }
        public IList<CT_MoveBookmark>   GetMoveToRangeStartList()            { return GetObjectList<CT_MoveBookmark>  (ItemsChoiceType50.moveToRangeStart); }
        public IList<CT_OMath>          GetOMathList()                       { return GetObjectList<CT_OMath>         (ItemsChoiceType50.oMath); }
        public IList<CT_OMathPara>      GetOMathParaList()                   { return GetObjectList<CT_OMathPara>     (ItemsChoiceType50.oMathPara); }
        public IList<CT_P>              GetPList()                           { return GetObjectList<CT_P>             (ItemsChoiceType50.p); }
        public IList<CT_Perm>           GetPermEndList()                     { return GetObjectList<CT_Perm>          (ItemsChoiceType50.permEnd); }
        public IList<CT_PermStart>      GetPermStartList()                   { return GetObjectList<CT_PermStart>     (ItemsChoiceType50.permStart); }
        public IList<CT_ProofErr>       GetProofErrList()                    { return GetObjectList<CT_ProofErr>      (ItemsChoiceType50.proofErr); }
        public IList<CT_SdtBlock>       GetSdtList()                         { return GetObjectList<CT_SdtBlock>      (ItemsChoiceType50.sdt); }
        public IList<CT_Tbl>            GetTblList()                         { return GetObjectList<CT_Tbl>           (ItemsChoiceType50.tbl); }

        public CT_AltChunk       InsertNewAltChunk(int p)                    { return InsertNewObject<CT_AltChunk>      (ItemsChoiceType50.altChunk, p); }
        public CT_MarkupRange    InsertNewBookmarkEnd(int p)                 { return InsertNewObject<CT_MarkupRange>   (ItemsChoiceType50.bookmarkEnd, p); }
        public CT_Bookmark       InsertNewBookmarkStart(int p)               { return InsertNewObject<CT_Bookmark>      (ItemsChoiceType50.bookmarkStart, p); }
        public CT_MarkupRange    InsertNewCommentRangeEnd(int p)             { return InsertNewObject<CT_MarkupRange>   (ItemsChoiceType50.commentRangeEnd, p); }
        public CT_MarkupRange    InsertNewCommentRangeStart(int p)           { return InsertNewObject<CT_MarkupRange>   (ItemsChoiceType50.commentRangeStart, p); }
        public CT_CustomXmlBlock InsertNewCustomXml(int p)                   { return InsertNewObject<CT_CustomXmlBlock>(ItemsChoiceType50.customXml, p); }
        public CT_Markup         InsertNewCustomXmlDelRangeEnd(int p)        { return InsertNewObject<CT_Markup>        (ItemsChoiceType50.customXmlDelRangeEnd, p); }
        public CT_TrackChange    InsertNewCustomXmlDelRangeStart(int p)      { return InsertNewObject<CT_TrackChange>   (ItemsChoiceType50.customXmlDelRangeStart, p); }
        public CT_Markup         InsertNewCustomXmlInsRangeEnd(int p)        { return InsertNewObject<CT_Markup>        (ItemsChoiceType50.customXmlInsRangeEnd, p); }
        public CT_TrackChange    InsertNewCustomXmlInsRangeStart(int p)      { return InsertNewObject<CT_TrackChange>   (ItemsChoiceType50.customXmlInsRangeStart, p); }
        public CT_Markup         InsertNewCustomXmlMoveFromRangeEnd(int p)   { return InsertNewObject<CT_Markup>        (ItemsChoiceType50.customXmlMoveFromRangeEnd, p); }
        public CT_TrackChange    InsertNewCustomXmlMoveFromRangeStart(int p) { return InsertNewObject<CT_TrackChange>   (ItemsChoiceType50.customXmlMoveFromRangeStart, p); }
        public CT_Markup         InsertNewCustomXmlMoveToRangeEnd(int p)     { return InsertNewObject<CT_Markup>        (ItemsChoiceType50.customXmlMoveToRangeEnd, p); }
        public CT_TrackChange    InsertNewCustomXmlMoveToRangeStart(int p)   { return InsertNewObject<CT_TrackChange>   (ItemsChoiceType50.customXmlMoveToRangeStart, p); }
        public CT_RunTrackChange InsertNewDel(int p)                         { return InsertNewObject<CT_RunTrackChange>(ItemsChoiceType50.del, p); }
        public CT_RunTrackChange InsertNewIns(int p)                         { return InsertNewObject<CT_RunTrackChange>(ItemsChoiceType50.ins, p); }
        public CT_RunTrackChange InsertNewMoveFrom(int p)                    { return InsertNewObject<CT_RunTrackChange>(ItemsChoiceType50.moveFrom, p); }
        public CT_MarkupRange    InsertNewMoveFromRangeEnd(int p)            { return InsertNewObject<CT_MarkupRange>   (ItemsChoiceType50.moveFromRangeEnd, p); }
        public CT_MoveBookmark   InsertNewMoveFromRangeStart(int p)          { return InsertNewObject<CT_MoveBookmark>  (ItemsChoiceType50.moveFromRangeStart, p); }
        public CT_RunTrackChange InsertNewMoveTo(int p)                      { return InsertNewObject<CT_RunTrackChange>(ItemsChoiceType50.moveTo, p); }
        public CT_MarkupRange    InsertNewMoveToRangeEnd(int p)              { return InsertNewObject<CT_MarkupRange>   (ItemsChoiceType50.moveToRangeEnd, p); }
        public CT_MoveBookmark   InsertNewMoveToRangeStart(int p)            { return InsertNewObject<CT_MoveBookmark>  (ItemsChoiceType50.moveToRangeStart, p); }
        public CT_OMath          InsertNewOMath(int p)                       { return InsertNewObject<CT_OMath>         (ItemsChoiceType50.oMath, p); }
        public CT_OMathPara      InsertNewOMathPara(int p)                   { return InsertNewObject<CT_OMathPara>     (ItemsChoiceType50.oMathPara, p); }
        public CT_P              InsertNewP(int p)                           { return InsertNewObject<CT_P>             (ItemsChoiceType50.p, p); }
        public CT_Perm           InsertNewPermEnd(int p)                     { return InsertNewObject<CT_Perm>          (ItemsChoiceType50.permEnd, p); }
        public CT_PermStart      InsertNewPermStart(int p)                   { return InsertNewObject<CT_PermStart>     (ItemsChoiceType50.permStart, p); }
        public CT_ProofErr       InsertNewProofErr(int p)                    { return InsertNewObject<CT_ProofErr>      (ItemsChoiceType50.proofErr, p); }
        public CT_SdtBlock       InsertNewSdt(int p)                         { return InsertNewObject<CT_SdtBlock>      (ItemsChoiceType50.sdt, p); }
        public CT_Tbl            InsertNewTbl(int p)                         { return InsertNewObject<CT_Tbl>           (ItemsChoiceType50.tbl, p); }

        public void RemoveAltChunk(int p)                    { RemoveObject(ItemsChoiceType50.altChunk, p); }
        public void RemoveBookmarkEnd(int p)                 { RemoveObject(ItemsChoiceType50.bookmarkEnd, p); }
        public void RemoveBookmarkStart(int p)               { RemoveObject(ItemsChoiceType50.bookmarkStart, p); }
        public void RemoveCommentRangeEnd(int p)             { RemoveObject(ItemsChoiceType50.commentRangeEnd, p); }
        public void RemoveCommentRangeStart(int p)           { RemoveObject(ItemsChoiceType50.commentRangeStart, p); }
        public void RemoveCustomXml(int p)                   { RemoveObject(ItemsChoiceType50.customXml, p); }
        public void RemoveCustomXmlDelRangeEnd(int p)        { RemoveObject(ItemsChoiceType50.customXmlDelRangeEnd, p); }
        public void RemoveCustomXmlDelRangeStart(int p)      { RemoveObject(ItemsChoiceType50.customXmlDelRangeStart, p); }
        public void RemoveCustomXmlInsRangeEnd(int p)        { RemoveObject(ItemsChoiceType50.customXmlInsRangeEnd, p); }
        public void RemoveCustomXmlInsRangeStart(int p)      { RemoveObject(ItemsChoiceType50.customXmlInsRangeStart, p); }
        public void RemoveCustomXmlMoveFromRangeEnd(int p)   { RemoveObject(ItemsChoiceType50.customXmlMoveFromRangeEnd, p); }
        public void RemoveCustomXmlMoveFromRangeStart(int p) { RemoveObject(ItemsChoiceType50.customXmlMoveFromRangeStart, p); }
        public void RemoveCustomXmlMoveToRangeEnd(int p)     { RemoveObject(ItemsChoiceType50.customXmlMoveToRangeEnd, p); }
        public void RemoveCustomXmlMoveToRangeStart(int p)   { RemoveObject(ItemsChoiceType50.customXmlMoveToRangeStart, p); }
        public void RemoveDel(int p)                         { RemoveObject(ItemsChoiceType50.del, p); }
        public void RemoveIns(int p)                         { RemoveObject(ItemsChoiceType50.ins, p); }
        public void RemoveMoveFrom(int p)                    { RemoveObject(ItemsChoiceType50.moveFrom, p); }
        public void RemoveMoveFromRangeEnd(int p)            { RemoveObject(ItemsChoiceType50.moveFromRangeEnd, p); }
        public void RemoveMoveFromRangeStart(int p)          { RemoveObject(ItemsChoiceType50.moveFromRangeStart, p); }
        public void RemoveMoveTo(int p)                      { RemoveObject(ItemsChoiceType50.moveTo, p); }
        public void RemoveMoveToRangeEnd(int p)              { RemoveObject(ItemsChoiceType50.moveToRangeEnd, p); }
        public void RemoveMoveToRangeStart(int p)            { RemoveObject(ItemsChoiceType50.moveToRangeStart, p); }
        public void RemoveOMath(int p)                       { RemoveObject(ItemsChoiceType50.oMath, p); }
        public void RemoveOMathPara(int p)                   { RemoveObject(ItemsChoiceType50.oMathPara, p); }
        public void RemoveP(int p)                           { RemoveObject(ItemsChoiceType50.p, p); }
        public void RemovePermEnd(int p)                     { RemoveObject(ItemsChoiceType50.permEnd, p); }
        public void RemovePermStart(int p)                   { RemoveObject(ItemsChoiceType50.permStart, p); }
        public void RemoveProofErr(int p)                    { RemoveObject(ItemsChoiceType50.proofErr, p); }
        public void RemoveSdt(int p)                         { RemoveObject(ItemsChoiceType50.sdt, p); }
        public void RemoveTbl(int p)                         { RemoveObject(ItemsChoiceType50.tbl, p); }

        public void SetAltChunkArray(int p, CT_AltChunk obj)                       { SetObject(ItemsChoiceType50.altChunk, p, obj); }
        public void SetBookmarkEndArray(int p, CT_MarkupRange obj)                 { SetObject(ItemsChoiceType50.bookmarkEnd, p, obj); }
        public void SetBookmarkStartArray(int p, CT_Bookmark obj)                  { SetObject(ItemsChoiceType50.bookmarkStart, p, obj); }
        public void SetCommentRangeEndArray(int p, CT_MarkupRange obj)             { SetObject(ItemsChoiceType50.commentRangeEnd, p, obj); }
        public void SetCommentRangeStartArray(int p, CT_MarkupRange obj)           { SetObject(ItemsChoiceType50.commentRangeStart, p, obj); }
        public void SetCustomXmlArray(int p, CT_CustomXmlBlock obj)                { SetObject(ItemsChoiceType50.customXml, p, obj); }
        public void SetCustomXmlDelRangeEndArray(int p, CT_Markup obj)             { SetObject(ItemsChoiceType50.customXmlDelRangeEnd, p, obj); }
        public void SetCustomXmlDelRangeStartArray(int p, CT_TrackChange obj)      { SetObject(ItemsChoiceType50.customXmlDelRangeStart, p, obj); }
        public void SetCustomXmlInsRangeEndArray(int p, CT_Markup obj)             { SetObject(ItemsChoiceType50.customXmlInsRangeEnd, p, obj); }
        public void SetCustomXmlInsRangeStartArray(int p, CT_TrackChange obj)      { SetObject(ItemsChoiceType50.customXmlInsRangeStart, p, obj); }
        public void SetCustomXmlMoveFromRangeEndArray(int p, CT_Markup obj)        { SetObject(ItemsChoiceType50.customXmlMoveFromRangeEnd, p, obj); }
        public void SetCustomXmlMoveFromRangeStartArray(int p, CT_TrackChange obj) { SetObject(ItemsChoiceType50.customXmlMoveFromRangeStart, p, obj); }
        public void SetCustomXmlMoveToRangeEndArray(int p, CT_Markup obj)          { SetObject(ItemsChoiceType50.customXmlMoveToRangeEnd, p, obj); }
        public void SetCustomXmlMoveToRangeStartArray(int p, CT_TrackChange obj)   { SetObject(ItemsChoiceType50.customXmlMoveToRangeStart, p, obj); }
        public void SetDelArray(int p, CT_RunTrackChange obj)                      { SetObject(ItemsChoiceType50.del, p, obj); }
        public void SetInsArray(int p, CT_RunTrackChange obj)                      { SetObject(ItemsChoiceType50.ins, p, obj); }
        public void SetMoveFromArray(int p, CT_RunTrackChange obj)                 { SetObject(ItemsChoiceType50.moveFrom, p, obj); }
        public void SetMoveFromRangeEndArray(int p, CT_MarkupRange obj)            { SetObject(ItemsChoiceType50.moveFromRangeEnd, p, obj); }
        public void SetMoveFromRangeStartArray(int p, CT_MoveBookmark obj)         { SetObject(ItemsChoiceType50.moveFromRangeStart, p, obj); }
        public void SetMoveToArray(int p, CT_RunTrackChange obj)                   { SetObject(ItemsChoiceType50.moveTo, p, obj); }
        public void SetMoveToRangeEndArray(int p, CT_MarkupRange obj)              { SetObject(ItemsChoiceType50.moveToRangeEnd, p, obj); }
        public void SetMoveToRangeStartArray(int p, CT_MoveBookmark obj)           { SetObject(ItemsChoiceType50.moveToRangeStart, p, obj); }
        public void SetOMathArray(int p, CT_OMath obj)                             { SetObject(ItemsChoiceType50.oMath, p, obj); }
        public void SetOMathParaArray(int p, CT_OMathPara obj)                     { SetObject(ItemsChoiceType50.oMathPara, p, obj); }
        public void SetPArray(int p, CT_P obj)                                     { SetObject(ItemsChoiceType50.p, p, obj); }
        public void SetPermEndArray(int p, CT_Perm obj)                            { SetObject(ItemsChoiceType50.permEnd, p, obj); }
        public void SetPermStartArray(int p, CT_PermStart obj)                     { SetObject(ItemsChoiceType50.permStart, p, obj); }
        public void SetProofErrArray(int p, CT_ProofErr obj)                       { SetObject(ItemsChoiceType50.proofErr, p, obj); }
        public void SetSdtArray(int p, CT_SdtBlock obj)                            { SetObject(ItemsChoiceType50.sdt, p, obj); }
        public void SetTblArray(int p, CT_Tbl obj)                                 { SetObject(ItemsChoiceType50.tbl, p, obj); }

        public int SizeOfAltChunkArray()                    { return SizeOfArray(ItemsChoiceType50.altChunk); }
        public int SizeOfBookmarkEndArray()                 { return SizeOfArray(ItemsChoiceType50.bookmarkEnd); }
        public int SizeOfBookmarkStartArray()               { return SizeOfArray(ItemsChoiceType50.bookmarkStart); }
        public int SizeOfCommentRangeEndArray()             { return SizeOfArray(ItemsChoiceType50.commentRangeEnd); }
        public int SizeOfCommentRangeStartArray()           { return SizeOfArray(ItemsChoiceType50.commentRangeStart); }
        public int SizeOfCustomXmlArray()                   { return SizeOfArray(ItemsChoiceType50.customXml); }
        public int SizeOfCustomXmlDelRangeEndArray()        { return SizeOfArray(ItemsChoiceType50.customXmlDelRangeEnd); }
        public int SizeOfCustomXmlDelRangeStartArray()      { return SizeOfArray(ItemsChoiceType50.customXmlDelRangeStart); }
        public int SizeOfCustomXmlInsRangeEndArray()        { return SizeOfArray(ItemsChoiceType50.customXmlInsRangeEnd); }
        public int SizeOfCustomXmlInsRangeStartArray()      { return SizeOfArray(ItemsChoiceType50.customXmlInsRangeStart); }
        public int SizeOfCustomXmlMoveFromRangeEndArray()   { return SizeOfArray(ItemsChoiceType50.customXmlMoveFromRangeEnd); }
        public int SizeOfCustomXmlMoveFromRangeStartArray() { return SizeOfArray(ItemsChoiceType50.customXmlMoveFromRangeStart); }
        public int SizeOfCustomXmlMoveToRangeEndArray()     { return SizeOfArray(ItemsChoiceType50.customXmlMoveToRangeEnd); }
        public int SizeOfCustomXmlMoveToRangeStartArray()   { return SizeOfArray(ItemsChoiceType50.customXmlMoveToRangeStart); }
        public int SizeOfDelArray()                         { return SizeOfArray(ItemsChoiceType50.del); }
        public int SizeOfInsArray()                         { return SizeOfArray(ItemsChoiceType50.ins); }
        public int SizeOfMoveFromArray()                    { return SizeOfArray(ItemsChoiceType50.moveFrom); }
        public int SizeOfMoveFromRangeEndArray()            { return SizeOfArray(ItemsChoiceType50.moveFromRangeEnd); }
        public int SizeOfMoveFromRangeStartArray()          { return SizeOfArray(ItemsChoiceType50.moveFromRangeStart); }
        public int SizeOfMoveToArray()                      { return SizeOfArray(ItemsChoiceType50.moveTo); }
        public int SizeOfMoveToRangeEndArray()              { return SizeOfArray(ItemsChoiceType50.moveToRangeEnd); }
        public int SizeOfMoveToRangeStartArray()            { return SizeOfArray(ItemsChoiceType50.moveToRangeStart); }
        public int SizeOfOMathArray()                       { return SizeOfArray(ItemsChoiceType50.oMath); }
        public int SizeOfOMathParaArray()                   { return SizeOfArray(ItemsChoiceType50.oMathPara); }
        public int SizeOfPArray()                           { return SizeOfArray(ItemsChoiceType50.p); }
        public int SizeOfPermEndArray()                     { return SizeOfArray(ItemsChoiceType50.permEnd); }
        public int SizeOfPermStartArray()                   { return SizeOfArray(ItemsChoiceType50.permStart); }
        public int SizeOfProofErrArray()                    { return SizeOfArray(ItemsChoiceType50.proofErr); }
        public int SizeOfSdtArray()                         { return SizeOfArray(ItemsChoiceType50.sdt); }
        public int SizeOfTblArray()                         { return SizeOfArray(ItemsChoiceType50.tbl); }
        #pragma warning restore format
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
