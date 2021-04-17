using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Serialization;
using NPOI.OpenXmlFormats.Shared;
using System.IO;
using NPOI.OpenXml4Net.Util;
using System.Collections;

namespace NPOI.OpenXmlFormats.Wordprocessing
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_P
    {

        private CT_PPr pPrField;

        private ArrayList itemsField;

        private List<ParagraphItemsChoiceType> itemsElementNameField;

        private byte[] rsidRPrField;

        private byte[] rsidRField;

        private byte[] rsidDelField;

        private byte[] rsidPField;

        private byte[] rsidRDefaultField;

        public CT_P()
        {
            this.itemsElementNameField = new List<ParagraphItemsChoiceType>();
            this.itemsField = new ArrayList();
            this.pPrField = new CT_PPr();
        }

        [XmlElement(Order = 0)]
        public CT_PPr pPr
        {
            get
            {
                return this.pPrField;
            }
            set
            {
                this.pPrField = value;
            }
        }
        public static CT_P Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_P ctObj = new CT_P();
            ctObj.rsidRPr = XmlHelper.ReadBytes(node.Attributes["w:rsidRPr"]);
            ctObj.rsidR = XmlHelper.ReadBytes(node.Attributes["w:rsidR"]);
            ctObj.rsidDel = XmlHelper.ReadBytes(node.Attributes["w:rsidDel"]);
            ctObj.rsidP = XmlHelper.ReadBytes(node.Attributes["w:rsidP"]);
            ctObj.rsidRDefault = XmlHelper.ReadBytes(node.Attributes["w:rsidRDefault"]);

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "pPr")
                {
                    ctObj.pPr = CT_PPr.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "bookmarkEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.bookmarkEnd);
                }
                else if (childNode.LocalName == "moveFromRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.moveFromRangeStart);
                }
                else if (childNode.LocalName == "moveTo")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.moveTo);
                }
                else if (childNode.LocalName == "oMathPara")
                {
                    ctObj.Items.Add(CT_OMathPara.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.oMathPara);
                }
                else if (childNode.LocalName == "oMath")
                {
                    ctObj.Items.Add(CT_OMath.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.oMath);
                }
                else if (childNode.LocalName == "bookmarkStart")
                {
                    ctObj.Items.Add(CT_Bookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.bookmarkStart);
                }
                else if (childNode.LocalName == "commentRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.commentRangeEnd);
                }
                else if (childNode.LocalName == "commentRangeStart")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.commentRangeStart);
                }
                else if (childNode.LocalName == "customXml")
                {
                    ctObj.Items.Add(CT_CustomXmlRun.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.customXml);
                }
                else if (childNode.LocalName == "customXmlDelRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.customXmlDelRangeEnd);
                }
                else if (childNode.LocalName == "customXmlDelRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.customXmlDelRangeStart);
                }
                else if (childNode.LocalName == "customXmlInsRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.customXmlInsRangeEnd);
                }
                else if (childNode.LocalName == "customXmlInsRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.customXmlInsRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.customXmlMoveFromRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.customXmlMoveFromRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.customXmlMoveToRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.customXmlMoveToRangeStart);
                }
                else if (childNode.LocalName == "del")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.del);
                }
                else if (childNode.LocalName == "fldSimple")
                {
                    ctObj.Items.Add(CT_SimpleField.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.fldSimple);
                }
                else if (childNode.LocalName == "hyperlink")
                {
                    ctObj.Items.Add(CT_Hyperlink1.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.hyperlink);
                }
                else if (childNode.LocalName == "ins")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.ins);
                }
                else if (childNode.LocalName == "moveFrom")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.moveFrom);
                }
                else if (childNode.LocalName == "moveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.moveFromRangeEnd);
                }
                else if (childNode.LocalName == "moveToRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.moveToRangeEnd);
                }
                else if (childNode.LocalName == "moveToRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.moveToRangeStart);
                }
                else if (childNode.LocalName == "permEnd")
                {
                    ctObj.Items.Add(CT_Perm.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.permEnd);
                }
                else if (childNode.LocalName == "permStart")
                {
                    ctObj.Items.Add(CT_PermStart.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.permStart);
                }
                else if (childNode.LocalName == "proofErr")
                {
                    ctObj.Items.Add(CT_ProofErr.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.proofErr);
                }
                else if (childNode.LocalName == "r")
                {
                    ctObj.Items.Add(CT_R.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.r);
                }
                else if (childNode.LocalName == "sdt")
                {
                    ctObj.Items.Add(CT_SdtRun.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.sdt);
                }
                else if (childNode.LocalName == "smartTag")
                {
                    ctObj.Items.Add(CT_SmartTagRun.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.smartTag);
                }
                else if (childNode.LocalName == "subDoc")
                {
                    ctObj.Items.Add(CT_Rel.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ParagraphItemsChoiceType.subDoc);
                }
            }
            return ctObj;
        }

        public bool IsSetRsidR()
        {
            return this.rsidRField != null && rsidRField.Length > 0;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:rsidR", this.rsidR);
            XmlHelper.WriteAttribute(sw, "w:rsidRPr", this.rsidRPr);
            XmlHelper.WriteAttribute(sw, "w:rsidRDefault", this.rsidRDefault);
            XmlHelper.WriteAttribute(sw, "w:rsidP", this.rsidP);
            XmlHelper.WriteAttribute(sw, "w:rsidDel", this.rsidDel);
            sw.Write(">");
            if (this.pPr != null)
                this.pPr.Write(sw, "pPr");

            int i = 0;
            foreach (object o in this.Items)
            {
                if (o is CT_MarkupRange && this.itemsElementNameField[i] == ParagraphItemsChoiceType.bookmarkEnd)
                    ((CT_MarkupRange)o).Write(sw, "bookmarkEnd");
                else if (o is CT_MoveBookmark && this.itemsElementNameField[i] == ParagraphItemsChoiceType.moveFromRangeStart)
                    ((CT_MoveBookmark)o).Write(sw, "moveFromRangeStart");
                else if (o is CT_RunTrackChange && this.itemsElementNameField[i] == ParagraphItemsChoiceType.moveTo)
                    ((CT_RunTrackChange)o).Write(sw, "moveTo");
                else if (o is CT_OMathPara)
                    ((CT_OMathPara)o).Write(sw, "oMathPara");
                else if (o is CT_OMath)
                    ((CT_OMath)o).Write(sw, "oMath");
                else if (o is CT_Bookmark && this.itemsElementNameField[i] == ParagraphItemsChoiceType.bookmarkStart)
                    ((CT_Bookmark)o).Write(sw, "bookmarkStart");
                else if (o is CT_MarkupRange&& this.itemsElementNameField[i] == ParagraphItemsChoiceType.commentRangeEnd)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeEnd");
                else if (o is CT_MarkupRange && this.itemsElementNameField[i] == ParagraphItemsChoiceType.commentRangeStart)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeStart");
                else if (o is CT_CustomXmlRun)
                    ((CT_CustomXmlRun)o).Write(sw, "customXml");
                else if (o is CT_Markup && this.itemsElementNameField[i] == ParagraphItemsChoiceType.customXmlDelRangeEnd)
                    ((CT_Markup)o).Write(sw, "customXmlDelRangeEnd");
                else if (o is CT_TrackChange && this.itemsElementNameField[i] == ParagraphItemsChoiceType.customXmlDelRangeStart)
                    ((CT_TrackChange)o).Write(sw, "customXmlDelRangeStart");
                else if (o is CT_Markup && this.itemsElementNameField[i] == ParagraphItemsChoiceType.customXmlInsRangeEnd)
                    ((CT_Markup)o).Write(sw, "customXmlInsRangeEnd");
                else if (o is CT_TrackChange && this.itemsElementNameField[i] == ParagraphItemsChoiceType.customXmlInsRangeStart)
                    ((CT_TrackChange)o).Write(sw, "customXmlInsRangeStart");
                else if (o is CT_Markup && this.itemsElementNameField[i] == ParagraphItemsChoiceType.customXmlMoveFromRangeEnd)
                    ((CT_Markup)o).Write(sw, "customXmlMoveFromRangeEnd");
                else if (o is CT_TrackChange && this.itemsElementNameField[i] == ParagraphItemsChoiceType.customXmlMoveFromRangeStart)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveFromRangeStart");
                else if (o is CT_Markup && this.itemsElementNameField[i] == ParagraphItemsChoiceType.customXmlMoveToRangeEnd)
                    ((CT_Markup)o).Write(sw, "customXmlMoveToRangeEnd");
                else if (o is CT_TrackChange && this.itemsElementNameField[i] == ParagraphItemsChoiceType.customXmlMoveToRangeStart)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveToRangeStart");
                else if (o is CT_RunTrackChange && this.itemsElementNameField[i] == ParagraphItemsChoiceType.del)
                    ((CT_RunTrackChange)o).Write(sw, "del");
                else if (o is CT_SimpleField)
                    ((CT_SimpleField)o).Write(sw, "fldSimple");
                else if (o is CT_Hyperlink1)
                    ((CT_Hyperlink1)o).Write(sw, "hyperlink");
                else if (o is CT_RunTrackChange && this.itemsElementNameField[i] == ParagraphItemsChoiceType.ins)
                    ((CT_RunTrackChange)o).Write(sw, "ins");
                else if (o is CT_RunTrackChange && this.itemsElementNameField[i] == ParagraphItemsChoiceType.moveFrom)
                    ((CT_RunTrackChange)o).Write(sw, "moveFrom");
                else if (o is CT_MarkupRange && this.itemsElementNameField[i] == ParagraphItemsChoiceType.moveFromRangeEnd)
                    ((CT_MarkupRange)o).Write(sw, "moveFromRangeEnd");
                else if (o is CT_MarkupRange && this.itemsElementNameField[i] == ParagraphItemsChoiceType.moveToRangeEnd)
                    ((CT_MarkupRange)o).Write(sw, "moveToRangeEnd");
                else if (o is CT_MoveBookmark && this.itemsElementNameField[i] == ParagraphItemsChoiceType.moveToRangeStart)
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
                i++;
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
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

        public CT_R AddNewR()
        {
            CT_R r = new CT_R();
            lock (this)
            {
                itemsField.Add(r);
                itemsElementNameField.Add(ParagraphItemsChoiceType.r);
            }
            return r;
        }

        public CT_OMath AddNewOMath()
        {
            CT_OMath oMath = new CT_OMath();
            lock (this)
            {
                itemsField.Add(oMath);
                itemsElementNameField.Add(ParagraphItemsChoiceType.oMath);
            }
            return oMath;
        }

        [XmlElement("ItemsElementName", Order = 2)]
        [XmlIgnore]
        public List<ParagraphItemsChoiceType> ItemsElementName
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

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] rsidRPr
        {
            get
            {
                return this.rsidRPrField;
            }
            set
            {
                this.rsidRPrField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] rsidR
        {
            get
            {
                return this.rsidRField;
            }
            set
            {
                this.rsidRField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] rsidDel
        {
            get
            {
                return this.rsidDelField;
            }
            set
            {
                this.rsidDelField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] rsidP
        {
            get
            {
                return this.rsidPField;
            }
            set
            {
                this.rsidPField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] rsidRDefault
        {
            get
            {
                return this.rsidRDefaultField;
            }
            set
            {
                this.rsidRDefaultField = value;
            }
        }

        public CT_PPr AddNewPPr()
        {
            if (this.pPrField == null)
                this.pPrField = new CT_PPr();
            return this.pPrField;
        }

        public void SetRArray(int pos, CT_R Run)
        {
            SetObject<CT_R>(ParagraphItemsChoiceType.r, pos, Run);
        }

        public CT_R InsertNewR(int pos)
        {
            return InsertNewObject<CT_R>(ParagraphItemsChoiceType.r, pos);
        }

        public int SizeOfRArray()
        {
            return SizeOfArray(ParagraphItemsChoiceType.r);
        }

        public void RemoveR(int pos)
        {
            RemoveObject(ParagraphItemsChoiceType.r, pos);
        }

        public IEnumerable<CT_MarkupRange> GetCommentRangeStartList()
        {
            return GetObjectList<CT_MarkupRange>(ParagraphItemsChoiceType.commentRangeStart);
        }
        public CT_Hyperlink1 AddNewHyperlink()
        {
            return AddNewObject<CT_Hyperlink1>(ParagraphItemsChoiceType.hyperlink);
        }
        public IEnumerable<CT_Hyperlink1> GetHyperlinkList()
        {
            return GetObjectList<CT_Hyperlink1>(ParagraphItemsChoiceType.hyperlink);
        }
        #region Generic methods for object operation

        private List<T> GetObjectList<T>(ParagraphItemsChoiceType type) where T : class
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
        private int SizeOfArray(ParagraphItemsChoiceType type)
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
        private T GetObjectArray<T>(int p, ParagraphItemsChoiceType type) where T : class
        {
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                if (pos < 0 || pos >= this.itemsField.Count)
                    return null;
                return itemsField[pos] as T;
            }
        }
        private T InsertNewObject<T>(ParagraphItemsChoiceType type, int p) where T : class, new()
        {
            T t = new T();
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                if (pos == -1)
                    pos = 0;
                this.itemsElementNameField.Insert(pos, type);
                this.itemsField.Insert(pos, t);
            }
            return t;
        }
        private T AddNewObject<T>(ParagraphItemsChoiceType type) where T : class, new()
        {
            T t = new T();
            lock (this)
            {
                this.itemsElementNameField.Add(type);
                this.itemsField.Add(t);
            }
            return t;
        }
        private void SetObject<T>(ParagraphItemsChoiceType type, int p, T obj) where T : class
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
        private int GetObjectIndex(ParagraphItemsChoiceType type, int p)
        {
            int index = -1;
            int pos = 0;
            for (int i = 0; i < itemsElementNameField.Count; i++)
            {
                if (itemsElementNameField[i] == type)
                {
                    if (pos == p)
                    {
                        //return itemsField[p] as T;
                        index = i;
                        break;
                    }
                    else
                        pos++;
                }
            }
            return index;
        }
        private void RemoveObject(ParagraphItemsChoiceType type, int p)
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


        public IList<CT_R> GetRList()
        {
            return GetObjectList<CT_R>(ParagraphItemsChoiceType.r);
        }

        public int SizeOfBookmarkStartArray()
        {
            return SizeOfArray(ParagraphItemsChoiceType.bookmarkStart);
        }

        public int SizeOfBookmarkEndArray()
        {
            return SizeOfArray(ParagraphItemsChoiceType.bookmarkEnd);
        }

        public CT_Bookmark GetBookmarkStartArray(int p)
        {
            return GetObjectArray<CT_Bookmark>(p, ParagraphItemsChoiceType.bookmarkStart);
        }

        public IEnumerable<CT_Bookmark> GetBookmarkStartList()
        {
            return GetObjectList<CT_Bookmark>(ParagraphItemsChoiceType.bookmarkStart);
        }
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ParagraphItemsChoiceType
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
    public class CT_PPr : CT_PPrBase
    {

        private CT_ParaRPr rPrField;

        private CT_SectPr sectPrField;

        private CT_PPrChange pPrChangeField;

        public CT_PPr()
        {

        }
        public CT_SectPr createSectPr()
        {
            this.sectPrField = new CT_SectPr();
            return this.sectPrField;
        }
        public override bool IsEmpty
        {
            get
            {
                return base.IsEmpty && 
                    rPrField == null && sectPrField == null && pPrChangeField == null;
            }
        }
        public static new CT_PPr Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PPr ctObj = new CT_PPr();
            ctObj.tabs = new List<CT_TabStop>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "rPr")
                    ctObj.rPr = CT_ParaRPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sectPr")
                    ctObj.sectPr = CT_SectPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pPrChange")
                    ctObj.pPrChange = CT_PPrChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pStyle")
                    ctObj.pStyle = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "keepNext")
                    ctObj.keepNext = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "keepLines")
                    ctObj.keepLines = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pageBreakBefore")
                    ctObj.pageBreakBefore = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "framePr")
                    ctObj.framePr = CT_FramePr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "widowControl")
                    ctObj.widowControl = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "numPr")
                    ctObj.numPr = CT_NumPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "suppressLineNumbers")
                    ctObj.suppressLineNumbers = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pBdr")
                    ctObj.pBdr = CT_PBdr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "shd")
                    ctObj.shd = CT_Shd.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "suppressAutoHyphens")
                    ctObj.suppressAutoHyphens = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "kinsoku")
                    ctObj.kinsoku = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "wordWrap")
                    ctObj.wordWrap = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "overflowPunct")
                    ctObj.overflowPunct = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "topLinePunct")
                    ctObj.topLinePunct = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "autoSpaceDE")
                    ctObj.autoSpaceDE = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "autoSpaceDN")
                    ctObj.autoSpaceDN = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bidi")
                    ctObj.bidi = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "adjustRightInd")
                    ctObj.adjustRightInd = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "snapToGrid")
                    ctObj.snapToGrid = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "spacing")
                    ctObj.spacing = CT_Spacing.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "ind")
                    ctObj.ind = CT_Ind.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "contextualSpacing")
                    ctObj.contextualSpacing = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "mirrorIndents")
                    ctObj.mirrorIndents = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "suppressOverlap")
                    ctObj.suppressOverlap = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "jc")
                    ctObj.jc = CT_Jc.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "textDirection")
                    ctObj.textDirection = CT_TextDirection.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "textAlignment")
                    ctObj.textAlignment = CT_TextAlignment.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "textboxTightWrap")
                    ctObj.textboxTightWrap = CT_TextboxTightWrap.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "outlineLvl")
                    ctObj.outlineLvl = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "divId")
                    ctObj.divId = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cnfStyle")
                    ctObj.cnfStyle = CT_Cnf.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tabs")
                {
                    foreach (XmlNode snode in childNode.ChildNodes)
                    {
                        ctObj.tabs.Add(CT_TabStop.Parse(snode, namespaceManager));
                    }
                }
            }
            return ctObj;
        }



        internal new void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.sectPr != null)
                this.sectPr.Write(sw, "sectPr");
            if (this.pPrChange != null)
                this.pPrChange.Write(sw, "pPrChange");
            if (this.pStyle != null)
                this.pStyle.Write(sw, "pStyle");
            if (this.keepNext != null)
                this.keepNext.Write(sw, "keepNext");
            if (this.keepLines != null)
                this.keepLines.Write(sw, "keepLines");
            if (this.pageBreakBefore != null)
                this.pageBreakBefore.Write(sw, "pageBreakBefore");
            if (this.framePr != null)
                this.framePr.Write(sw, "framePr");
            if (this.widowControl != null)
                this.widowControl.Write(sw, "widowControl");
            if (this.numPr != null)
                this.numPr.Write(sw, "numPr");
            if (this.pBdr != null)
                this.pBdr.Write(sw, "pBdr");
            if (this.tabs != null && this.tabs.Count > 0)
            {
                sw.Write("<w:tabs>");
                foreach (CT_TabStop x in this.tabs)
                {
                    x.Write(sw, "tab");
                }
                sw.Write("</w:tabs>");
            }
            if (this.suppressLineNumbers != null)
                this.suppressLineNumbers.Write(sw, "suppressLineNumbers");
            if (this.shd != null)
                this.shd.Write(sw, "shd");
            if (this.suppressAutoHyphens != null)
                this.suppressAutoHyphens.Write(sw, "suppressAutoHyphens");
            if (this.kinsoku != null)
                this.kinsoku.Write(sw, "kinsoku");
            if (this.wordWrap != null)
                this.wordWrap.Write(sw, "wordWrap");
            if (this.overflowPunct != null)
                this.overflowPunct.Write(sw, "overflowPunct");
            if (this.topLinePunct != null)
                this.topLinePunct.Write(sw, "topLinePunct");
            if (this.autoSpaceDE != null)
                this.autoSpaceDE.Write(sw, "autoSpaceDE");
            if (this.autoSpaceDN != null)
                this.autoSpaceDN.Write(sw, "autoSpaceDN");
            if (this.bidi != null)
                this.bidi.Write(sw, "bidi");
            if (this.adjustRightInd != null)
                this.adjustRightInd.Write(sw, "adjustRightInd");
            if (this.snapToGrid != null)
                this.snapToGrid.Write(sw, "snapToGrid");
            if (this.spacing != null)
                this.spacing.Write(sw, "spacing");
            if (this.ind != null)
                this.ind.Write(sw, "ind");

            if (this.contextualSpacing != null)
                this.contextualSpacing.Write(sw, "contextualSpacing");
            if (this.mirrorIndents != null)
                this.mirrorIndents.Write(sw, "mirrorIndents");
            if (this.suppressOverlap != null)
                this.suppressOverlap.Write(sw, "suppressOverlap");
            if (this.jc != null)
                this.jc.Write(sw, "jc");
            if (this.outlineLvl != null)
                this.outlineLvl.Write(sw, "outlineLvl");
            if (this.rPr != null)
                this.rPr.Write(sw, "rPr");
            if (this.textDirection != null)
                this.textDirection.Write(sw, "textDirection");
            if (this.textAlignment != null)
                this.textAlignment.Write(sw, "textAlignment");
            if (this.textboxTightWrap != null)
                this.textboxTightWrap.Write(sw, "textboxTightWrap");
            if (this.divId != null)
                this.divId.Write(sw, "divId");
            if (this.cnfStyle != null)
                this.cnfStyle.Write(sw, "cnfStyle");

            sw.Write(string.Format("</w:{0}>", nodeName));
        }



        [XmlElement(Order = 0)]
        public CT_ParaRPr rPr
        {
            get
            {
                return this.rPrField;
            }
            set
            {
                this.rPrField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_SectPr sectPr
        {
            get
            {
                return this.sectPrField;
            }
            set
            {
                this.sectPrField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_PPrChange pPrChange
        {
            get
            {
                return this.pPrChangeField;
            }
            set
            {
                this.pPrChangeField = value;
            }
        }

        public CT_ParaRPr AddNewRPr()
        {
            if (this.rPrField == null)
                this.rPrField = new CT_ParaRPr();
            return this.rPrField;
        }

        public CT_NumPr AddNewNumPr()
        {
            if (this.numPr == null)
                this.numPr = new CT_NumPr();
            return this.numPr;
        }

        public bool IsSetSpacing()
        {
            return this.spacing != null;
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_ParaRPr
    {

        private CT_TrackChange insField;

        private CT_TrackChange delField;

        private CT_TrackChange moveFromField;

        private CT_TrackChange moveToField;

        private CT_String rStyleField;

        private CT_Fonts rFontsField;

        private CT_OnOff bField;

        private CT_OnOff bCsField;

        private CT_OnOff iField;

        private CT_OnOff iCsField;

        private CT_OnOff capsField;

        private CT_OnOff smallCapsField;

        private CT_OnOff strikeField;

        private CT_OnOff dstrikeField;

        private CT_OnOff outlineField;

        private CT_OnOff shadowField;

        private CT_OnOff embossField;

        private CT_OnOff imprintField;

        private CT_OnOff noProofField;

        private CT_OnOff snapToGridField;

        private CT_OnOff vanishField;

        private CT_OnOff webHiddenField;

        private CT_Color colorField;

        private CT_SignedTwipsMeasure spacingField;

        private CT_TextScale wField;

        private CT_HpsMeasure kernField;

        private CT_SignedHpsMeasure positionField;

        private CT_HpsMeasure szField;

        private CT_HpsMeasure szCsField;

        private CT_Highlight highlightField;

        private CT_Underline uField;

        private CT_TextEffect effectField;

        private CT_Border bdrField;

        private CT_Shd shdField;

        private CT_FitText fitTextField;

        private CT_VerticalAlignRun vertAlignField;

        private CT_OnOff rtlField;

        private CT_OnOff csField;

        private CT_Em emField;

        private CT_Language langField;

        private CT_EastAsianLayout eastAsianLayoutField;

        private CT_OnOff specVanishField;

        private CT_OnOff oMathField;

        private CT_ParaRPrChange rPrChangeField;

        public CT_ParaRPr()
        {

        }

        [XmlElement(Order = 0)]
        public CT_TrackChange ins
        {
            get
            {
                return this.insField;
            }
            set
            {
                this.insField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_TrackChange del
        {
            get
            {
                return this.delField;
            }
            set
            {
                this.delField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_TrackChange moveFrom
        {
            get
            {
                return this.moveFromField;
            }
            set
            {
                this.moveFromField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_TrackChange moveTo
        {
            get
            {
                return this.moveToField;
            }
            set
            {
                this.moveToField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_String rStyle
        {
            get
            {
                return this.rStyleField;
            }
            set
            {
                this.rStyleField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_Fonts rFonts
        {
            get
            {
                return this.rFontsField;
            }
            set
            {
                this.rFontsField = value;
            }
        }

        [XmlElement(Order = 6)]
        public CT_OnOff b
        {
            get
            {
                return this.bField;
            }
            set
            {
                this.bField = value;
            }
        }

        [XmlElement(Order = 7)]
        public CT_OnOff bCs
        {
            get
            {
                return this.bCsField;
            }
            set
            {
                this.bCsField = value;
            }
        }

        [XmlElement(Order = 8)]
        public CT_OnOff i
        {
            get
            {
                return this.iField;
            }
            set
            {
                this.iField = value;
            }
        }

        [XmlElement(Order = 9)]
        public CT_OnOff iCs
        {
            get
            {
                return this.iCsField;
            }
            set
            {
                this.iCsField = value;
            }
        }

        [XmlElement(Order = 10)]
        public CT_OnOff caps
        {
            get
            {
                return this.capsField;
            }
            set
            {
                this.capsField = value;
            }
        }

        [XmlElement(Order = 11)]
        public CT_OnOff smallCaps
        {
            get
            {
                return this.smallCapsField;
            }
            set
            {
                this.smallCapsField = value;
            }
        }

        [XmlElement(Order = 12)]
        public CT_OnOff strike
        {
            get
            {
                return this.strikeField;
            }
            set
            {
                this.strikeField = value;
            }
        }

        [XmlElement(Order = 13)]
        public CT_OnOff dstrike
        {
            get
            {
                return this.dstrikeField;
            }
            set
            {
                this.dstrikeField = value;
            }
        }

        [XmlElement(Order = 14)]
        public CT_OnOff outline
        {
            get
            {
                return this.outlineField;
            }
            set
            {
                this.outlineField = value;
            }
        }

        [XmlElement(Order = 15)]
        public CT_OnOff shadow
        {
            get
            {
                return this.shadowField;
            }
            set
            {
                this.shadowField = value;
            }
        }

        [XmlElement(Order = 16)]
        public CT_OnOff emboss
        {
            get
            {
                return this.embossField;
            }
            set
            {
                this.embossField = value;
            }
        }

        [XmlElement(Order = 17)]
        public CT_OnOff imprint
        {
            get
            {
                return this.imprintField;
            }
            set
            {
                this.imprintField = value;
            }
        }

        [XmlElement(Order = 18)]
        public CT_OnOff noProof
        {
            get
            {
                return this.noProofField;
            }
            set
            {
                this.noProofField = value;
            }
        }

        [XmlElement(Order = 19)]
        public CT_OnOff snapToGrid
        {
            get
            {
                return this.snapToGridField;
            }
            set
            {
                this.snapToGridField = value;
            }
        }

        [XmlElement(Order = 20)]
        public CT_OnOff vanish
        {
            get
            {
                return this.vanishField;
            }
            set
            {
                this.vanishField = value;
            }
        }

        [XmlElement(Order = 21)]
        public CT_OnOff webHidden
        {
            get
            {
                return this.webHiddenField;
            }
            set
            {
                this.webHiddenField = value;
            }
        }

        [XmlElement(Order = 22)]
        public CT_Color color
        {
            get
            {
                return this.colorField;
            }
            set
            {
                this.colorField = value;
            }
        }

        [XmlElement(Order = 23)]
        public CT_SignedTwipsMeasure spacing
        {
            get
            {
                return this.spacingField;
            }
            set
            {
                this.spacingField = value;
            }
        }

        [XmlElement(Order = 24)]
        public CT_TextScale w
        {
            get
            {
                return this.wField;
            }
            set
            {
                this.wField = value;
            }
        }

        [XmlElement(Order = 25)]
        public CT_HpsMeasure kern
        {
            get
            {
                return this.kernField;
            }
            set
            {
                this.kernField = value;
            }
        }

        [XmlElement(Order = 26)]
        public CT_SignedHpsMeasure position
        {
            get
            {
                return this.positionField;
            }
            set
            {
                this.positionField = value;
            }
        }

        [XmlElement(Order = 27)]
        public CT_HpsMeasure sz
        {
            get
            {
                return this.szField;
            }
            set
            {
                this.szField = value;
            }
        }

        [XmlElement(Order = 28)]
        public CT_HpsMeasure szCs
        {
            get
            {
                return this.szCsField;
            }
            set
            {
                this.szCsField = value;
            }
        }

        [XmlElement(Order = 29)]
        public CT_Highlight highlight
        {
            get
            {
                return this.highlightField;
            }
            set
            {
                this.highlightField = value;
            }
        }

        [XmlElement(Order = 30)]
        public CT_Underline u
        {
            get
            {
                return this.uField;
            }
            set
            {
                this.uField = value;
            }
        }

        [XmlElement(Order = 31)]
        public CT_TextEffect effect
        {
            get
            {
                return this.effectField;
            }
            set
            {
                this.effectField = value;
            }
        }

        [XmlElement(Order = 32)]
        public CT_Border bdr
        {
            get
            {
                return this.bdrField;
            }
            set
            {
                this.bdrField = value;
            }
        }

        [XmlElement(Order = 33)]
        public CT_Shd shd
        {
            get
            {
                return this.shdField;
            }
            set
            {
                this.shdField = value;
            }
        }

        [XmlElement(Order = 34)]
        public CT_FitText fitText
        {
            get
            {
                return this.fitTextField;
            }
            set
            {
                this.fitTextField = value;
            }
        }

        [XmlElement(Order = 35)]
        public CT_VerticalAlignRun vertAlign
        {
            get
            {
                return this.vertAlignField;
            }
            set
            {
                this.vertAlignField = value;
            }
        }

        [XmlElement(Order = 36)]
        public CT_OnOff rtl
        {
            get
            {
                return this.rtlField;
            }
            set
            {
                this.rtlField = value;
            }
        }

        [XmlElement(Order = 37)]
        public CT_OnOff cs
        {
            get
            {
                return this.csField;
            }
            set
            {
                this.csField = value;
            }
        }

        [XmlElement(Order = 38)]
        public CT_Em em
        {
            get
            {
                return this.emField;
            }
            set
            {
                this.emField = value;
            }
        }

        [XmlElement(Order = 39)]
        public CT_Language lang
        {
            get
            {
                return this.langField;
            }
            set
            {
                this.langField = value;
            }
        }

        [XmlElement(Order = 40)]
        public CT_EastAsianLayout eastAsianLayout
        {
            get
            {
                return this.eastAsianLayoutField;
            }
            set
            {
                this.eastAsianLayoutField = value;
            }
        }

        [XmlElement(Order = 41)]
        public CT_OnOff specVanish
        {
            get
            {
                return this.specVanishField;
            }
            set
            {
                this.specVanishField = value;
            }
        }

        [XmlElement(Order = 42)]
        public CT_OnOff oMath
        {
            get
            {
                return this.oMathField;
            }
            set
            {
                this.oMathField = value;
            }
        }

        [XmlElement(Order = 43)]
        public CT_ParaRPrChange rPrChange
        {
            get
            {
                return this.rPrChangeField;
            }
            set
            {
                this.rPrChangeField = value;
            }
        }

        public CT_OnOff AddNewNoProof()
        {
            if (this.noProofField == null)
                this.noProofField = new CT_OnOff();
            return this.noProofField;
        }

        public static CT_ParaRPr Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ParaRPr ctObj = new CT_ParaRPr();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "ins")
                    ctObj.ins = CT_TrackChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "del")
                    ctObj.del = CT_TrackChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "moveFrom")
                    ctObj.moveFrom = CT_TrackChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "moveTo")
                    ctObj.moveTo = CT_TrackChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "rStyle")
                    ctObj.rStyle = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "rFonts")
                    ctObj.rFonts = CT_Fonts.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "b")
                    ctObj.b = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bCs")
                    ctObj.bCs = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "i")
                    ctObj.i = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "iCs")
                    ctObj.iCs = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "caps")
                    ctObj.caps = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "smallCaps")
                    ctObj.smallCaps = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "strike")
                    ctObj.strike = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dstrike")
                    ctObj.dstrike = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "outline")
                    ctObj.outline = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "shadow")
                    ctObj.shadow = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "emboss")
                    ctObj.emboss = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "imprint")
                    ctObj.imprint = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "noProof")
                    ctObj.noProof = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "snapToGrid")
                    ctObj.snapToGrid = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "vanish")
                    ctObj.vanish = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "webHidden")
                    ctObj.webHidden = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "color")
                    ctObj.color = CT_Color.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "spacing")
                    ctObj.spacing = CT_SignedTwipsMeasure.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "w")
                    ctObj.w = CT_TextScale.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "kern")
                    ctObj.kern = CT_HpsMeasure.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "position")
                    ctObj.position = CT_SignedHpsMeasure.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sz")
                    ctObj.sz = CT_HpsMeasure.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "szCs")
                    ctObj.szCs = CT_HpsMeasure.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "highlight")
                    ctObj.highlight = CT_Highlight.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "u")
                    ctObj.u = CT_Underline.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "effect")
                    ctObj.effect = CT_TextEffect.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bdr")
                    ctObj.bdr = CT_Border.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "shd")
                    ctObj.shd = CT_Shd.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "fitText")
                    ctObj.fitText = CT_FitText.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "vertAlign")
                    ctObj.vertAlign = CT_VerticalAlignRun.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "rtl")
                    ctObj.rtl = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cs")
                    ctObj.cs = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "em")
                    ctObj.em = CT_Em.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lang")
                    ctObj.lang = CT_Language.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "eastAsianLayout")
                    ctObj.eastAsianLayout = CT_EastAsianLayout.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "specVanish")
                    ctObj.specVanish = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "oMath")
                    ctObj.oMath = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "rPrChange")
                    ctObj.rPrChange = CT_ParaRPrChange.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.ins != null)
                this.ins.Write(sw, "ins");
            if (this.del != null)
                this.del.Write(sw, "del");
            if (this.moveFrom != null)
                this.moveFrom.Write(sw, "moveFrom");
            if (this.moveTo != null)
                this.moveTo.Write(sw, "moveTo");
            if (this.rStyle != null)
                this.rStyle.Write(sw, "rStyle");
            if (this.rFonts != null)
                this.rFonts.Write(sw, "rFonts");
            if (this.b != null)
                this.b.Write(sw, "b");
            if (this.bCs != null)
                this.bCs.Write(sw, "bCs");
            if (this.i != null)
                this.i.Write(sw, "i");
            if (this.iCs != null)
                this.iCs.Write(sw, "iCs");
            if (this.caps != null)
                this.caps.Write(sw, "caps");
            if (this.smallCaps != null)
                this.smallCaps.Write(sw, "smallCaps");
            if (this.strike != null)
                this.strike.Write(sw, "strike");
            if (this.dstrike != null)
                this.dstrike.Write(sw, "dstrike");
            if (this.outline != null)
                this.outline.Write(sw, "outline");
            if (this.shadow != null)
                this.shadow.Write(sw, "shadow");
            if (this.emboss != null)
                this.emboss.Write(sw, "emboss");
            if (this.imprint != null)
                this.imprint.Write(sw, "imprint");
            if (this.noProof != null)
                this.noProof.Write(sw, "noProof");
            if (this.snapToGrid != null)
                this.snapToGrid.Write(sw, "snapToGrid");
            if (this.vanish != null)
                this.vanish.Write(sw, "vanish");
            if (this.webHidden != null)
                this.webHidden.Write(sw, "webHidden");
            if (this.color != null)
                this.color.Write(sw, "color");
            if (this.spacing != null)
                this.spacing.Write(sw, "spacing");
            if (this.w != null)
                this.w.Write(sw, "w");
            if (this.kern != null)
                this.kern.Write(sw, "kern");
            if (this.position != null)
                this.position.Write(sw, "position");
            if (this.sz != null)
                this.sz.Write(sw, "sz");
            if (this.szCs != null)
                this.szCs.Write(sw, "szCs");
            if (this.highlight != null)
                this.highlight.Write(sw, "highlight");
            if (this.u != null)
                this.u.Write(sw, "u");
            if (this.effect != null)
                this.effect.Write(sw, "effect");
            if (this.bdr != null)
                this.bdr.Write(sw, "bdr");
            if (this.shd != null)
                this.shd.Write(sw, "shd");
            if (this.fitText != null)
                this.fitText.Write(sw, "fitText");
            if (this.vertAlign != null)
                this.vertAlign.Write(sw, "vertAlign");
            if (this.rtl != null)
                this.rtl.Write(sw, "rtl");
            if (this.cs != null)
                this.cs.Write(sw, "cs");
            if (this.em != null)
                this.em.Write(sw, "em");
            if (this.lang != null)
                this.lang.Write(sw, "lang");
            if (this.eastAsianLayout != null)
                this.eastAsianLayout.Write(sw, "eastAsianLayout");
            if (this.specVanish != null)
                this.specVanish.Write(sw, "specVanish");
            if (this.oMath != null)
                this.oMath.Write(sw, "oMath");
            if (this.rPrChange != null)
                this.rPrChange.Write(sw, "rPrChange");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }


    }




    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SectPrChange : CT_TrackChange
    {

        private CT_SectPrBase sectPrField;

        public CT_SectPrChange()
        {
            this.sectPrField = new CT_SectPrBase();
        }
        public static new CT_SectPrChange Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SectPrChange ctObj = new CT_SectPrChange();
            ctObj.author = XmlHelper.ReadString(node.Attributes["w:author"]);
            ctObj.date = XmlHelper.ReadString(node.Attributes["w:date"]);
            ctObj.id = XmlHelper.ReadString(node.Attributes["r:id"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "sectPr")
                    ctObj.sectPr = CT_SectPrBase.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal new void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:author", this.author);
            XmlHelper.WriteAttribute(sw, "w:date", this.date);
            XmlHelper.WriteAttribute(sw, "r:id", this.id);
            sw.Write(">");
            if (this.sectPr != null)
                this.sectPr.Write(sw, "sectPr");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_SectPrBase sectPr
        {
            get
            {
                return this.sectPrField;
            }
            set
            {
                this.sectPrField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SectPrBase
    {

        private CT_FtnProps footnotePrField;

        private CT_EdnProps endnotePrField;

        private CT_SectType typeField;

        private CT_PageSz pgSzField;

        private CT_PageMar pgMarField;

        private CT_PaperSource paperSrcField;

        private CT_PageBorders pgBordersField;

        private CT_LineNumber lnNumTypeField;

        private CT_PageNumber pgNumTypeField;

        private CT_Columns colsField;

        private CT_OnOff formProtField;

        private CT_VerticalJc vAlignField;

        private CT_OnOff noEndnoteField;

        private CT_OnOff titlePgField;

        private CT_TextDirection textDirectionField;

        private CT_OnOff bidiField;

        private CT_OnOff rtlGutterField;

        private CT_DocGrid docGridField;

        private CT_Rel printerSettingsField;

        private byte[] rsidRPrField;

        private byte[] rsidDelField;

        private byte[] rsidRField;

        private byte[] rsidSectField;

        public CT_SectPrBase()
        {
            //this.printerSettingsField = new CT_Rel();
            //this.docGridField = new CT_DocGrid();
            //this.rtlGutterField = new CT_OnOff();
            //this.bidiField = new CT_OnOff();
            this.textDirectionField = new CT_TextDirection();
            //this.titlePgField = new CT_OnOff();
            //this.noEndnoteField = new CT_OnOff();
            //this.vAlignField = new CT_VerticalJc();
            //this.formProtField = new CT_OnOff();
            this.colsField = new CT_Columns();
            //this.pgNumTypeField = new CT_PageNumber();
            //this.lnNumTypeField = new CT_LineNumber();
            //this.pgBordersField = new CT_PageBorders();
            //this.paperSrcField = new CT_PaperSource();
            //this.pgMarField = new CT_PageMar();
            //this.pgSzField = new CT_PageSz();
            //this.typeField = new CT_SectType();
            //this.endnotePrField = new CT_EdnProps();
            //this.footnotePrField = new CT_FtnProps();
        }
        public static CT_SectPrBase Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SectPrBase ctObj = new CT_SectPrBase();
            ctObj.rsidRPr = XmlHelper.ReadBytes(node.Attributes["w:rsidRPr"]);
            ctObj.rsidDel = XmlHelper.ReadBytes(node.Attributes["w:rsidDel"]);
            ctObj.rsidR = XmlHelper.ReadBytes(node.Attributes["w:rsidR"]);
            ctObj.rsidSect = XmlHelper.ReadBytes(node.Attributes["w:rsidSect"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "footnotePr")
                    ctObj.footnotePr = CT_FtnProps.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "endnotePr")
                    ctObj.endnotePr = CT_EdnProps.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "type")
                    ctObj.type = CT_SectType.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pgSz")
                    ctObj.pgSz = CT_PageSz.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pgMar")
                    ctObj.pgMar = CT_PageMar.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "paperSrc")
                    ctObj.paperSrc = CT_PaperSource.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pgBorders")
                    ctObj.pgBorders = CT_PageBorders.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lnNumType")
                    ctObj.lnNumType = CT_LineNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pgNumType")
                    ctObj.pgNumType = CT_PageNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cols")
                    ctObj.cols = CT_Columns.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "formProt")
                    ctObj.formProt = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "vAlign")
                    ctObj.vAlign = CT_VerticalJc.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "noEndnote")
                    ctObj.noEndnote = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "titlePg")
                    ctObj.titlePg = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "textDirection")
                    ctObj.textDirection = CT_TextDirection.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bidi")
                    ctObj.bidi = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "rtlGutter")
                    ctObj.rtlGutter = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "docGrid")
                    ctObj.docGrid = CT_DocGrid.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "printerSettings")
                    ctObj.printerSettings = CT_Rel.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:rsidRPr", this.rsidRPr);
            XmlHelper.WriteAttribute(sw, "w:rsidDel", this.rsidDel);
            XmlHelper.WriteAttribute(sw, "w:rsidR", this.rsidR);
            XmlHelper.WriteAttribute(sw, "w:rsidSect", this.rsidSect);
            sw.Write(">");
            if (this.footnotePr != null)
                this.footnotePr.Write(sw, "footnotePr");
            if (this.endnotePr != null)
                this.endnotePr.Write(sw, "endnotePr");
            if (this.type != null)
                this.type.Write(sw, "type");
            if (this.pgSz != null)
                this.pgSz.Write(sw, "pgSz");
            if (this.pgMar != null)
                this.pgMar.Write(sw, "pgMar");
            if (this.paperSrc != null)
                this.paperSrc.Write(sw, "paperSrc");
            if (this.pgBorders != null)
                this.pgBorders.Write(sw, "pgBorders");
            if (this.lnNumType != null)
                this.lnNumType.Write(sw, "lnNumType");
            if (this.pgNumType != null)
                this.pgNumType.Write(sw, "pgNumType");
            if (this.cols != null)
                this.cols.Write(sw, "cols");
            if (this.formProt != null)
                this.formProt.Write(sw, "formProt");
            if (this.vAlign != null)
                this.vAlign.Write(sw, "vAlign");
            if (this.noEndnote != null)
                this.noEndnote.Write(sw, "noEndnote");
            if (this.titlePg != null)
                this.titlePg.Write(sw, "titlePg");
            if (this.textDirection != null)
                this.textDirection.Write(sw, "textDirection");
            if (this.bidi != null)
                this.bidi.Write(sw, "bidi");
            if (this.rtlGutter != null)
                this.rtlGutter.Write(sw, "rtlGutter");
            if (this.docGrid != null)
                this.docGrid.Write(sw, "docGrid");
            if (this.printerSettings != null)
                this.printerSettings.Write(sw, "printerSettings");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_FtnProps footnotePr
        {
            get
            {
                return this.footnotePrField;
            }
            set
            {
                this.footnotePrField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_EdnProps endnotePr
        {
            get
            {
                return this.endnotePrField;
            }
            set
            {
                this.endnotePrField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_SectType type
        {
            get
            {
                return this.typeField;
            }
            set
            {
                this.typeField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_PageSz pgSz
        {
            get
            {
                return this.pgSzField;
            }
            set
            {
                this.pgSzField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_PageMar pgMar
        {
            get
            {
                return this.pgMarField;
            }
            set
            {
                this.pgMarField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_PaperSource paperSrc
        {
            get
            {
                return this.paperSrcField;
            }
            set
            {
                this.paperSrcField = value;
            }
        }

        [XmlElement(Order = 6)]
        public CT_PageBorders pgBorders
        {
            get
            {
                return this.pgBordersField;
            }
            set
            {
                this.pgBordersField = value;
            }
        }

        [XmlElement(Order = 7)]
        public CT_LineNumber lnNumType
        {
            get
            {
                return this.lnNumTypeField;
            }
            set
            {
                this.lnNumTypeField = value;
            }
        }

        [XmlElement(Order = 8)]
        public CT_PageNumber pgNumType
        {
            get
            {
                return this.pgNumTypeField;
            }
            set
            {
                this.pgNumTypeField = value;
            }
        }

        [XmlElement(Order = 9)]
        public CT_Columns cols
        {
            get
            {
                return this.colsField;
            }
            set
            {
                this.colsField = value;
            }
        }

        [XmlElement(Order = 10)]
        public CT_OnOff formProt
        {
            get
            {
                return this.formProtField;
            }
            set
            {
                this.formProtField = value;
            }
        }

        [XmlElement(Order = 11)]
        public CT_VerticalJc vAlign
        {
            get
            {
                return this.vAlignField;
            }
            set
            {
                this.vAlignField = value;
            }
        }

        [XmlElement(Order = 12)]
        public CT_OnOff noEndnote
        {
            get
            {
                return this.noEndnoteField;
            }
            set
            {
                this.noEndnoteField = value;
            }
        }

        [XmlElement(Order = 13)]
        public CT_OnOff titlePg
        {
            get
            {
                return this.titlePgField;
            }
            set
            {
                this.titlePgField = value;
            }
        }

        [XmlElement(Order = 14)]
        public CT_TextDirection textDirection
        {
            get
            {
                return this.textDirectionField;
            }
            set
            {
                this.textDirectionField = value;
            }
        }

        [XmlElement(Order = 15)]
        public CT_OnOff bidi
        {
            get
            {
                return this.bidiField;
            }
            set
            {
                this.bidiField = value;
            }
        }

        [XmlElement(Order = 16)]
        public CT_OnOff rtlGutter
        {
            get
            {
                return this.rtlGutterField;
            }
            set
            {
                this.rtlGutterField = value;
            }
        }

        [XmlElement(Order = 17)]
        public CT_DocGrid docGrid
        {
            get
            {
                return this.docGridField;
            }
            set
            {
                this.docGridField = value;
            }
        }

        [XmlElement(Order = 18)]
        public CT_Rel printerSettings
        {
            get
            {
                return this.printerSettingsField;
            }
            set
            {
                this.printerSettingsField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] rsidRPr
        {
            get
            {
                return this.rsidRPrField;
            }
            set
            {
                this.rsidRPrField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] rsidDel
        {
            get
            {
                return this.rsidDelField;
            }
            set
            {
                this.rsidDelField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] rsidR
        {
            get
            {
                return this.rsidRField;
            }
            set
            {
                this.rsidRField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] rsidSect
        {
            get
            {
                return this.rsidSectField;
            }
            set
            {
                this.rsidSectField = value;
            }
        }
    }

    
    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SectPr
    {
                private CT_FtnProps footnotePrField;

        private CT_EdnProps endnotePrField;

        private CT_SectType typeField;

        private CT_PageSz pgSzField;

        private CT_PageMar pgMarField;

        private CT_PaperSource paperSrcField;

        private CT_PageBorders pgBordersField;

        private CT_LineNumber lnNumTypeField;

        private CT_PageNumber pgNumTypeField;

        private CT_Columns colsField;

        private CT_OnOff formProtField;

        private CT_VerticalJc vAlignField;

        private CT_OnOff noEndnoteField;

        private CT_OnOff titlePgField;

        private CT_TextDirection textDirectionField;

        private CT_OnOff bidiField;

        private CT_OnOff rtlGutterField;

        private CT_DocGrid docGridField;

        private CT_Rel printerSettingsField;

        private CT_SectPrChange sectPrChangeField;

        private byte[] rsidRPrField;

        private byte[] rsidDelField;

        private byte[] rsidRField;

        private byte[] rsidSectField;

        public CT_SectPr()
        {
            this.sectPrChangeField = new CT_SectPrChange();
            //this.printerSettingsField = new CT_Rel();
            this.docGridField = new CT_DocGrid();
            this.docGrid.type = ST_DocGrid.lines;
            this.docGrid.typeSpecified = true;
            this.docGrid.linePitch = "312";
            
            //this.rtlGutterField = new CT_OnOff();
            //this.bidiField = new CT_OnOff();
            this.textDirectionField = new CT_TextDirection();
            //this.titlePgField = new CT_OnOff();
            //this.noEndnoteField = new CT_OnOff();
            //this.vAlignField = new CT_VerticalJc();
            //this.formProtField = new CT_OnOff();
            this.colsField = new CT_Columns();
            this.cols.space = 425;
            this.cols.spaceSpecified = true;
            this.pgNumTypeField = new CT_PageNumber();
            //this.lnNumTypeField = new CT_LineNumber();
            //this.pgBordersField = new CT_PageBorders();
            //this.paperSrcField = new CT_PaperSource();
            this.pgMarField = new CT_PageMar();
            this.pgMar.top = "1440";
            this.pgMar.right = 1800;
            this.pgMar.bottom = "1440";
            this.pgMar.left = 1800;
            this.pgMar.header = 851;
            this.pgMar.footer = 992;
            this.pgMar.gutter = 0;

            this.pgSzField = new CT_PageSz();
            this.pgSz.w = 11906;
            this.pgSz.wSpecified = true;
            this.pgSz.h = 16838;
            this.pgSz.hSpecified = true;

            //this.typeField = new CT_SectType();
            //this.endnotePrField = new CT_EdnProps();
            //this.footnotePrField = new CT_FtnProps();
            //this.itemsElementNameField = new List<ItemsChoiceHdrFtrRefType>();
            //this.itemsField = new List<CT_HdrFtrRef>();
        }
        List<CT_HdrFtrRef> footerReferenceField;
        public List<CT_HdrFtrRef> footerReference
        {
            get { return this.footerReferenceField; }
            set { this.footerReferenceField = value; }
        }

        List<CT_HdrFtrRef> headerReferenceField;
        public List<CT_HdrFtrRef> headerReference
        {
            get { return this.headerReferenceField; }
            set { this.headerReferenceField = value; }
        }

        [XmlElement(Order = 2)]
        public CT_FtnProps footnotePr
        {
            get
            {
                return this.footnotePrField;
            }
            set
            {
                this.footnotePrField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_EdnProps endnotePr
        {
            get
            {
                return this.endnotePrField;
            }
            set
            {
                this.endnotePrField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_SectType type
        {
            get
            {
                return this.typeField;
            }
            set
            {
                this.typeField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_PageSz pgSz
        {
            get
            {
                return this.pgSzField;
            }
            set
            {
                this.pgSzField = value;
            }
        }

        [XmlElement(Order = 6)]
        public CT_PageMar pgMar
        {
            get
            {
                return this.pgMarField;
            }
            set
            {
                this.pgMarField = value;
            }
        }

        [XmlElement(Order = 7)]
        public CT_PaperSource paperSrc
        {
            get
            {
                return this.paperSrcField;
            }
            set
            {
                this.paperSrcField = value;
            }
        }

        [XmlElement(Order = 8)]
        public CT_PageBorders pgBorders
        {
            get
            {
                return this.pgBordersField;
            }
            set
            {
                this.pgBordersField = value;
            }
        }

        [XmlElement(Order = 9)]
        public CT_LineNumber lnNumType
        {
            get
            {
                return this.lnNumTypeField;
            }
            set
            {
                this.lnNumTypeField = value;
            }
        }

        [XmlElement(Order = 10)]
        public CT_PageNumber pgNumType
        {
            get
            {
                return this.pgNumTypeField;
            }
            set
            {
                this.pgNumTypeField = value;
            }
        }

        [XmlElement(Order = 11)]
        public CT_Columns cols
        {
            get
            {
                return this.colsField;
            }
            set
            {
                this.colsField = value;
            }
        }

        [XmlElement(Order = 12)]
        public CT_OnOff formProt
        {
            get
            {
                return this.formProtField;
            }
            set
            {
                this.formProtField = value;
            }
        }

        [XmlElement(Order = 13)]
        public CT_VerticalJc vAlign
        {
            get
            {
                return this.vAlignField;
            }
            set
            {
                this.vAlignField = value;
            }
        }

        [XmlElement(Order = 14)]
        public CT_OnOff noEndnote
        {
            get
            {
                return this.noEndnoteField;
            }
            set
            {
                this.noEndnoteField = value;
            }
        }

        [XmlElement(Order = 15)]
        public CT_OnOff titlePg
        {
            get
            {
                return this.titlePgField;
            }
            set
            {
                this.titlePgField = value;
            }
        }

        [XmlElement(Order = 16)]
        public CT_TextDirection textDirection
        {
            get
            {
                return this.textDirectionField;
            }
            set
            {
                this.textDirectionField = value;
            }
        }

        [XmlElement(Order = 17)]
        public CT_OnOff bidi
        {
            get
            {
                return this.bidiField;
            }
            set
            {
                this.bidiField = value;
            }
        }

        [XmlElement(Order = 18)]
        public CT_OnOff rtlGutter
        {
            get
            {
                return this.rtlGutterField;
            }
            set
            {
                this.rtlGutterField = value;
            }
        }

        [XmlElement(Order = 19)]
        public CT_DocGrid docGrid
        {
            get
            {
                return this.docGridField;
            }
            set
            {
                this.docGridField = value;
            }
        }

        [XmlElement(Order = 20)]
        public CT_Rel printerSettings
        {
            get
            {
                return this.printerSettingsField;
            }
            set
            {
                this.printerSettingsField = value;
            }
        }

        [XmlElement(Order = 21)]
        public CT_SectPrChange sectPrChange
        {
            get
            {
                return this.sectPrChangeField;
            }
            set
            {
                this.sectPrChangeField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] rsidRPr
        {
            get
            {
                return this.rsidRPrField;
            }
            set
            {
                this.rsidRPrField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] rsidDel
        {
            get
            {
                return this.rsidDelField;
            }
            set
            {
                this.rsidDelField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] rsidR
        {
            get
            {
                return this.rsidRField;
            }
            set
            {
                this.rsidRField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] rsidSect
        {
            get
            {
                return this.rsidSectField;
            }
            set
            {
                this.rsidSectField = value;
            }
        }
        public CT_HdrFtrRef AddNewHeaderReference()
        {
            CT_HdrFtrRef ref1 = new CT_HdrFtrRef();
            if (this.headerReferenceField == null)
                this.headerReferenceField = new List<CT_HdrFtrRef>();

            this.headerReferenceField.Add(ref1);
            return ref1;
        }

        public CT_HdrFtrRef AddNewFooterReference()
        {
            CT_HdrFtrRef ref1 = new CT_HdrFtrRef();
            if (this.footerReferenceField == null)
                this.footerReferenceField = new List<CT_HdrFtrRef>();
            this.footerReferenceField.Add(ref1);
            return ref1;
        }
        public int SizeOfHeaderReferenceArray()
        {
            if (headerReferenceField == null)
                return 0;
            return headerReferenceField.Count;
        }
        public CT_HdrFtrRef GetHeaderReferenceArray(int i)
        {
            if (headerReferenceField == null)
                return null;
            return headerReferenceField[i];
        }

        public int SizeOfFooterReferenceArray()
        {
            if (footerReferenceField == null)
                return 0;
            return footerReferenceField.Count;
        }

        public CT_HdrFtrRef GetFooterReferenceArray(int i)
        {
            if (footerReferenceField == null)
                return null;
            return footerReferenceField[i];
        }

        public static CT_SectPr Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SectPr ctObj = new CT_SectPr();
            ctObj.rsidRPr = XmlHelper.ReadBytes(node.Attributes["w:rsidRPr"]);
            ctObj.rsidDel = XmlHelper.ReadBytes(node.Attributes["w:rsidDel"]);
            ctObj.rsidR = XmlHelper.ReadBytes(node.Attributes["w:rsidR"]);
            ctObj.rsidSect = XmlHelper.ReadBytes(node.Attributes["w:rsidSect"]);
            ctObj.footerReference = new List<CT_HdrFtrRef>();
            ctObj.headerReference = new List<CT_HdrFtrRef>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "footnotePr")
                    ctObj.footnotePr = CT_FtnProps.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "endnotePr")
                    ctObj.endnotePr = CT_EdnProps.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "type")
                    ctObj.type = CT_SectType.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pgSz")
                    ctObj.pgSz = CT_PageSz.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pgMar")
                    ctObj.pgMar = CT_PageMar.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "paperSrc")
                    ctObj.paperSrc = CT_PaperSource.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pgBorders")
                    ctObj.pgBorders = CT_PageBorders.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lnNumType")
                    ctObj.lnNumType = CT_LineNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pgNumType")
                    ctObj.pgNumType = CT_PageNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cols")
                    ctObj.cols = CT_Columns.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "formProt")
                    ctObj.formProt = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "vAlign")
                    ctObj.vAlign = CT_VerticalJc.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "noEndnote")
                    ctObj.noEndnote = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "titlePg")
                    ctObj.titlePg = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "textDirection")
                    ctObj.textDirection = CT_TextDirection.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bidi")
                    ctObj.bidi = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "rtlGutter")
                    ctObj.rtlGutter = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "docGrid")
                    ctObj.docGrid = CT_DocGrid.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "printerSettings")
                    ctObj.printerSettings = CT_Rel.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sectPrChange")
                    ctObj.sectPrChange = CT_SectPrChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "footerReference")
                    ctObj.footerReference.Add(CT_HdrFtrRef.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "headerReference")
                    ctObj.headerReference.Add(CT_HdrFtrRef.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:rsidR", this.rsidR);
            XmlHelper.WriteAttribute(sw, "w:rsidRPr", this.rsidRPr);
            XmlHelper.WriteAttribute(sw, "w:rsidSect", this.rsidSect);
            XmlHelper.WriteAttribute(sw, "w:rsidDel", this.rsidDel);
            sw.Write(">");
            if (this.headerReference != null)
            {
                foreach (CT_HdrFtrRef x in this.headerReference)
                {
                    x.Write(sw, "headerReference");
                }
            }
            if (this.footerReference != null)
            {
                foreach (CT_HdrFtrRef x in this.footerReference)
                {
                    x.Write(sw, "footerReference");
                }
            }
            if (this.footnotePr != null)
                this.footnotePr.Write(sw, "footnotePr");
            if (this.endnotePr != null)
                this.endnotePr.Write(sw, "endnotePr");
            if (this.type != null)
                this.type.Write(sw, "type");
            if (this.pgSz != null)
                this.pgSz.Write(sw, "pgSz");
            if (this.pgMar != null)
                this.pgMar.Write(sw, "pgMar");
            if (this.paperSrc != null)
                this.paperSrc.Write(sw, "paperSrc");
            if (this.pgBorders != null)
                this.pgBorders.Write(sw, "pgBorders");
            if (this.lnNumType != null)
                this.lnNumType.Write(sw, "lnNumType");
            if (this.pgNumType != null)
                this.pgNumType.Write(sw, "pgNumType");
            if (this.cols != null)
                this.cols.Write(sw, "cols");
            if (this.formProt != null)
                this.formProt.Write(sw, "formProt");
            if (this.vAlign != null)
                this.vAlign.Write(sw, "vAlign");
            if (this.noEndnote != null)
                this.noEndnote.Write(sw, "noEndnote");
            if (this.titlePg != null)
                this.titlePg.Write(sw, "titlePg");
            if (this.textDirection != null)
                this.textDirection.Write(sw, "textDirection");
            if (this.bidi != null)
                this.bidi.Write(sw, "bidi");
            if (this.rtlGutter != null)
                this.rtlGutter.Write(sw, "rtlGutter");
            if (this.docGrid != null)
                this.docGrid.Write(sw, "docGrid");
            if (this.printerSettings != null)
                this.printerSettings.Write(sw, "printerSettings");
            if (this.sectPrChange != null)
                this.sectPrChange.Write(sw, "sectPrChange");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }


    }
    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_PBdr
    {

        private CT_Border topField;

        private CT_Border leftField;

        private CT_Border bottomField;

        private CT_Border rightField;

        private CT_Border betweenField;

        private CT_Border barField;

        public CT_PBdr()
        {

        }
        public static CT_PBdr Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PBdr ctObj = new CT_PBdr();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "top")
                    ctObj.top = CT_Border.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "left")
                    ctObj.left = CT_Border.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bottom")
                    ctObj.bottom = CT_Border.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "right")
                    ctObj.right = CT_Border.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "between")
                    ctObj.between = CT_Border.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bar")
                    ctObj.bar = CT_Border.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.top != null)
                this.top.Write(sw, "top");
            if (this.left != null)
                this.left.Write(sw, "left");
            if (this.bottom != null)
                this.bottom.Write(sw, "bottom");
            if (this.right != null)
                this.right.Write(sw, "right");
            if (this.between != null)
                this.between.Write(sw, "between");
            if (this.bar != null)
                this.bar.Write(sw, "bar");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }



        [XmlElement(Order = 0)]
        public CT_Border top
        {
            get
            {
                return this.topField;
            }
            set
            {
                this.topField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_Border left
        {
            get
            {
                return this.leftField;
            }
            set
            {
                this.leftField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_Border bottom
        {
            get
            {
                return this.bottomField;
            }
            set
            {
                this.bottomField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_Border right
        {
            get
            {
                return this.rightField;
            }
            set
            {
                this.rightField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_Border between
        {
            get
            {
                return this.betweenField;
            }
            set
            {
                this.betweenField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_Border bar
        {
            get
            {
                return this.barField;
            }
            set
            {
                this.barField = value;
            }
        }

        public bool IsSetTop()
        {
            return this.topField != null && this.topField.val != ST_Border.none && this.topField.val != ST_Border.nil;
        }

        public CT_Border AddNewTop()
        {
            if (this.topField == null)
                this.topField = new CT_Border();
            return this.topField;
        }

        public void UnsetTop()
        {
            this.topField = new CT_Border();//?? set a new border or set null;
        }

        public bool IsSetBottom()
        {
            return this.bottomField != null && this.bottomField.val != ST_Border.none && this.bottomField.val != ST_Border.nil;
        }

        public CT_Border AddNewBottom()
        {
            if (this.bottomField == null)
                this.bottomField = new CT_Border();
            return this.bottomField;
        }

        public void UnsetBottom()
        {
            this.bottomField = null;// new CT_Border();
        }

        public bool IsSetRight()
        {
            return this.rightField != null && this.rightField.val != ST_Border.none && this.rightField.val != ST_Border.nil;
        }

        public void UnsetRight()
        {
            this.rightField = null;// new CT_Border();
        }

        public CT_Border AddNewRight()
        {
            if (this.rightField == null)
                this.rightField = new CT_Border();
            return this.rightField;
        }

        public bool IsSetBetween()
        {
            return this.betweenField != null && this.betweenField.val != ST_Border.none && this.betweenField.val != ST_Border.nil;
        }

        public CT_Border AddNewBetween()
        {
            if (this.betweenField == null)
                this.betweenField = new CT_Border();
            return this.betweenField;
        }

        public void UnsetBetween()
        {
            this.betweenField = null;// new CT_Border();
        }

        public bool IsSetLeft()
        {
            return this.leftField != null && this.leftField.val != ST_Border.none && this.leftField.val != ST_Border.nil;
        }

        public CT_Border AddNewLeft()
        {
            if (this.leftField == null)
                this.leftField = new CT_Border();
            return this.leftField;
        }

        public void UnsetLeft()
        {
            this.leftField = null;// new CT_Border();
        }
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Spacing
    {

        private ulong beforeField;

        private bool beforeFieldSpecified;

        private string beforeLinesField;

        private ST_OnOff beforeAutospacingField;

        private bool beforeAutospacingFieldSpecified;

        private ulong afterField;

        private bool afterFieldSpecified;

        private string afterLinesField;

        private ST_OnOff afterAutospacingField;

        private bool afterAutospacingFieldSpecified;

        private string lineField;

        private ST_LineSpacingRule lineRuleField;

        private bool lineRuleFieldSpecified;
        public static CT_Spacing Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Spacing ctObj = new CT_Spacing();
            ctObj.before = XmlHelper.ReadULong(node.Attributes["w:before"]);
            ctObj.beforeLines = XmlHelper.ReadString(node.Attributes["w:beforeLines"]);
            if (node.Attributes["w:beforeAutospacing"] != null)
                ctObj.beforeAutospacing = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:beforeAutospacing"].Value,true);
            ctObj.after = XmlHelper.ReadULong(node.Attributes["w:after"]);
            ctObj.afterLines = XmlHelper.ReadString(node.Attributes["w:afterLines"]);
            if (node.Attributes["w:afterAutospacing"] != null)
                ctObj.afterAutospacing = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:afterAutospacing"].Value,true);
            ctObj.line = XmlHelper.ReadString(node.Attributes["w:line"]);
            if (node.Attributes["w:lineRule"] != null)
                ctObj.lineRule = (ST_LineSpacingRule)Enum.Parse(typeof(ST_LineSpacingRule), node.Attributes["w:lineRule"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:before", this.before);
            XmlHelper.WriteAttribute(sw, "w:beforeLines", this.beforeLines);
            if(this.beforeAutospacing!= ST_OnOff.off)
                XmlHelper.WriteAttribute(sw, "w:beforeAutospacing", this.beforeAutospacing.ToString());
            XmlHelper.WriteAttribute(sw, "w:after", this.after,false);
            XmlHelper.WriteAttribute(sw, "w:afterLines", this.afterLines);
            if (this.afterAutospacing != ST_OnOff.off)
                XmlHelper.WriteAttribute(sw, "w:afterAutospacing", this.afterAutospacing.ToString());
            XmlHelper.WriteAttribute(sw, "w:line", this.line);
            if(this.lineRule!= ST_LineSpacingRule.nil)
                XmlHelper.WriteAttribute(sw, "w:lineRule", this.lineRule.ToString());
            sw.Write("/>");
        }


        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong before
        {
            get
            {
                return this.beforeField;
            }
            set
            {
                this.beforeField = value;
            }
        }

        [XmlIgnore]
        public bool beforeSpecified
        {
            get
            {
                return this.beforeFieldSpecified;
            }
            set
            {
                this.beforeFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string beforeLines
        {
            get
            {
                return this.beforeLinesField;
            }
            set
            {
                this.beforeLinesField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff beforeAutospacing
        {
            get
            {
                return this.beforeAutospacingField;
            }
            set
            {
                this.beforeAutospacingField = value;
            }
        }

        [XmlIgnore]
        public bool beforeAutospacingSpecified
        {
            get
            {
                return this.beforeAutospacingFieldSpecified;
            }
            set
            {
                this.beforeAutospacingFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong after
        {
            get
            {
                return this.afterField;
            }
            set
            {
                this.afterField = value;
            }
        }

        [XmlIgnore]
        public bool afterSpecified
        {
            get
            {
                return this.afterFieldSpecified;
            }
            set
            {
                this.afterFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string afterLines
        {
            get
            {
                return this.afterLinesField;
            }
            set
            {
                this.afterLinesField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff afterAutospacing
        {
            get
            {
                return this.afterAutospacingField;
            }
            set
            {
                this.afterAutospacingField = value;
            }
        }

        [XmlIgnore]
        public bool afterAutospacingSpecified
        {
            get
            {
                return this.afterAutospacingFieldSpecified;
            }
            set
            {
                this.afterAutospacingFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string line
        {
            get
            {
                return this.lineField;
            }
            set
            {
                this.lineField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_LineSpacingRule lineRule
        {
            get
            {
                return this.lineRuleField;
            }
            set
            {
                this.lineRuleField = value;
            }
        }

        [XmlIgnore]
        public bool lineRuleSpecified
        {
            get
            {
                return this.lineRuleFieldSpecified;
            }
            set
            {
                this.lineRuleFieldSpecified = value;
            }
        }

        public bool IsSetBefore()
        {
            return !(this.beforeField == 0);
        }

        public bool IsSetBeforeLines()
        {
            return !string.IsNullOrEmpty(this.beforeLinesField);
        }

        public bool IsSetLineRule()
        {
            return !(this.lineRuleField == ST_LineSpacingRule.nil);
        }
        public bool IsSetBetweenLines()
        {
            return !string.IsNullOrEmpty(this.lineField);
        }
        public bool IsSetAfter()
        {
            return !(this.afterField == 0);
        }

        public bool IsSetAfterLines()
        {
            return !string.IsNullOrEmpty(this.afterLinesField);
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_LineSpacingRule
    {
        nil,
    
        auto,

    
        exact,

    
        atLeast,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Ind
    {

        private string leftField;

        private string leftCharsField;

        private string rightField;

        private string rightCharsField;

        private ulong hangingField;

        private bool hangingFieldSpecified;

        private string hangingCharsField;

        private long firstLineField;

        private bool firstLineFieldSpecified;

        private string firstLineCharsField;
        public CT_Ind()
        {
            firstLineField = -1;
        }

        public static CT_Ind Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Ind ctObj = new CT_Ind();
            ctObj.left = XmlHelper.ReadString(node.Attributes["w:left"]);
            ctObj.leftChars = XmlHelper.ReadString(node.Attributes["w:leftChars"]);
            ctObj.right = XmlHelper.ReadString(node.Attributes["w:right"]);
            ctObj.rightChars = XmlHelper.ReadString(node.Attributes["w:rightChars"]);
            ctObj.hanging = XmlHelper.ReadULong(node.Attributes["w:hanging"]);
            ctObj.hangingChars = XmlHelper.ReadString(node.Attributes["w:hangingChars"]);
            if (node.Attributes["w:firstLine"]!=null)
                ctObj.firstLine = XmlHelper.ReadLong(node.Attributes["w:firstLine"]);
            ctObj.firstLineChars = XmlHelper.ReadString(node.Attributes["w:firstLineChars"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:leftChars", this.leftChars);
            XmlHelper.WriteAttribute(sw, "w:left", this.left);
            XmlHelper.WriteAttribute(sw, "w:rightChars", this.rightChars);
            XmlHelper.WriteAttribute(sw, "w:right", this.right);
            XmlHelper.WriteAttribute(sw, "w:hangingChars", this.hangingChars);
            XmlHelper.WriteAttribute(sw, "w:hanging", this.hanging);
            XmlHelper.WriteAttribute(sw, "w:firstLineChars", this.firstLineChars);
            if(firstLineField>=0)
                XmlHelper.WriteAttribute(sw, "w:firstLine", this.firstLine, true);
            sw.Write("/>");
        }


        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string left
        {
            get
            {
                return this.leftField;
            }
            set
            {
                this.leftField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string leftChars
        {
            get
            {
                return this.leftCharsField;
            }
            set
            {
                this.leftCharsField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string right
        {
            get
            {
                return this.rightField;
            }
            set
            {
                this.rightField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string rightChars
        {
            get
            {
                return this.rightCharsField;
            }
            set
            {
                this.rightCharsField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong hanging
        {
            get
            {
                return this.hangingField;
            }
            set
            {
                this.hangingField = value;
                this.hangingFieldSpecified = true;
            }
        }

        [XmlIgnore]
        public bool hangingSpecified
        {
            get
            {
                return this.hangingFieldSpecified;
            }
            set
            {
                this.hangingFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string hangingChars
        {
            get
            {
                return this.hangingCharsField;
            }
            set
            {
                this.hangingCharsField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public long firstLine
        {
            get
            {
                return this.firstLineField;
            }
            set
            {
                this.firstLineField = value;
            }
        }

        [XmlIgnore]
        public bool firstLineSpecified
        {
            get
            {
                return this.firstLineFieldSpecified;
            }
            set
            {
                this.firstLineFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string firstLineChars
        {
            get
            {
                return this.firstLineCharsField;
            }
            set
            {
                this.firstLineCharsField = value;
            }
        }

        public bool IsSetLeft()
        {
            return !string.IsNullOrEmpty(this.leftField);
        }

        public bool IsSetRight()
        {
            return !string.IsNullOrEmpty(this.rightField);
        }

        public bool IsSetHanging()
        {
            return !(this.hangingField == 0);
        }

        public bool IsSetFirstLine()
        {
            return !(this.firstLineField==0);
        }
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_RPr
    {

        private CT_String rStyleField;

        private CT_Fonts rFontsField;

        private CT_OnOff bField;

        private CT_OnOff bCsField;

        private CT_OnOff iField;

        private CT_OnOff iCsField;

        private CT_OnOff capsField;

        private CT_OnOff smallCapsField;

        private CT_OnOff strikeField;

        private CT_OnOff dstrikeField;

        private CT_OnOff outlineField;

        private CT_OnOff shadowField;

        private CT_OnOff embossField;

        private CT_OnOff imprintField;

        private CT_OnOff noProofField;

        private CT_OnOff snapToGridField;

        private CT_OnOff vanishField;

        private CT_OnOff webHiddenField;

        private CT_Color colorField;

        private CT_SignedTwipsMeasure spacingField;

        private CT_TextScale wField;

        private CT_HpsMeasure kernField;

        private CT_SignedHpsMeasure positionField;

        private CT_HpsMeasure szField;

        private CT_HpsMeasure szCsField;

        private CT_Highlight highlightField;

        private CT_Underline uField;

        private CT_TextEffect effectField;

        private CT_Border bdrField;

        private CT_Shd shdField;

        private CT_FitText fitTextField;

        private CT_VerticalAlignRun vertAlignField;

        private CT_OnOff rtlField;

        private CT_OnOff csField;

        private CT_Em emField;

        private CT_Language langField;

        private CT_EastAsianLayout eastAsianLayoutField;

        private CT_OnOff specVanishField;

        private CT_OnOff oMathField;

        private CT_RPrChange rPrChangeField;

        public CT_RPr()
        {
        }
        public static CT_RPr Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_RPr ctObj = new CT_RPr();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "rStyle")
                    ctObj.rStyle = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "rFonts")
                    ctObj.rFonts = CT_Fonts.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "b")
                    ctObj.b = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bCs")
                    ctObj.bCs = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "i")
                    ctObj.i = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "iCs")
                    ctObj.iCs = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "caps")
                    ctObj.caps = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "smallCaps")
                    ctObj.smallCaps = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "strike")
                    ctObj.strike = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dstrike")
                    ctObj.dstrike = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "outline")
                    ctObj.outline = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "shadow")
                    ctObj.shadow = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "emboss")
                    ctObj.emboss = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "imprint")
                    ctObj.imprint = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "noProof")
                    ctObj.noProof = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "snapToGrid")
                    ctObj.snapToGrid = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "vanish")
                    ctObj.vanish = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "webHidden")
                    ctObj.webHidden = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "color")
                    ctObj.color = CT_Color.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "spacing")
                    ctObj.spacing = CT_SignedTwipsMeasure.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "w")
                    ctObj.w = CT_TextScale.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "kern")
                    ctObj.kern = CT_HpsMeasure.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "position")
                    ctObj.position = CT_SignedHpsMeasure.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sz")
                    ctObj.sz = CT_HpsMeasure.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "szCs")
                    ctObj.szCs = CT_HpsMeasure.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "highlight")
                    ctObj.highlight = CT_Highlight.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "u")
                    ctObj.u = CT_Underline.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "effect")
                    ctObj.effect = CT_TextEffect.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bdr")
                    ctObj.bdr = CT_Border.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "shd")
                    ctObj.shd = CT_Shd.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "fitText")
                    ctObj.fitText = CT_FitText.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "vertAlign")
                    ctObj.vertAlign = CT_VerticalAlignRun.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "rtl")
                    ctObj.rtl = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cs")
                    ctObj.cs = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "em")
                    ctObj.em = CT_Em.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lang")
                    ctObj.lang = CT_Language.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "eastAsianLayout")
                    ctObj.eastAsianLayout = CT_EastAsianLayout.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "specVanish")
                    ctObj.specVanish = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "oMath")
                    ctObj.oMath = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "rPrChange")
                    ctObj.rPrChange = CT_RPrChange.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.rStyle != null)
                this.rStyle.Write(sw, "rStyle");
            if (this.rFonts != null)
                this.rFonts.Write(sw, "rFonts");
            if (this.b != null)
                this.b.Write(sw, "b");
            if (this.bCs != null)
                this.bCs.Write(sw, "bCs");
            if (this.i != null)
                this.i.Write(sw, "i");
            if (this.iCs != null)
                this.iCs.Write(sw, "iCs");
            if (this.caps != null)
                this.caps.Write(sw, "caps");
            if (this.smallCaps != null)
                this.smallCaps.Write(sw, "smallCaps");
            if (this.strike != null)
                this.strike.Write(sw, "strike");
            if (this.dstrike != null)
                this.dstrike.Write(sw, "dstrike");
            if (this.outline != null)
                this.outline.Write(sw, "outline");
            if (this.shadow != null)
                this.shadow.Write(sw, "shadow");
            if (this.emboss != null)
                this.emboss.Write(sw, "emboss");
            if (this.imprint != null)
                this.imprint.Write(sw, "imprint");
            if (this.noProof != null)
                this.noProof.Write(sw, "noProof");
            if (this.snapToGrid != null)
                this.snapToGrid.Write(sw, "snapToGrid");
            if (this.vanish != null)
                this.vanish.Write(sw, "vanish");
            if (this.webHidden != null)
                this.webHidden.Write(sw, "webHidden");
            if (this.color != null)
                this.color.Write(sw, "color");
            if (this.spacing != null)
                this.spacing.Write(sw, "spacing");
            if (this.w != null)
                this.w.Write(sw, "w");
            if (this.kern != null)
                this.kern.Write(sw, "kern");
            if (this.position != null)
                this.position.Write(sw, "position");
            if (this.sz != null)
                this.sz.Write(sw, "sz");
            if (this.szCs != null)
                this.szCs.Write(sw, "szCs");
            if (this.highlight != null)
                this.highlight.Write(sw, "highlight");
            if (this.u != null)
                this.u.Write(sw, "u");
            if (this.effect != null)
                this.effect.Write(sw, "effect");
            if (this.bdr != null)
                this.bdr.Write(sw, "bdr");
            if (this.shd != null)
                this.shd.Write(sw, "shd");
            if (this.fitText != null)
                this.fitText.Write(sw, "fitText");
            if (this.vertAlign != null)
                this.vertAlign.Write(sw, "vertAlign");
            if (this.rtl != null)
                this.rtl.Write(sw, "rtl");
            if (this.cs != null)
                this.cs.Write(sw, "cs");
            if (this.em != null)
                this.em.Write(sw, "em");
            if (this.lang != null)
                this.lang.Write(sw, "lang");
            if (this.eastAsianLayout != null)
                this.eastAsianLayout.Write(sw, "eastAsianLayout");
            if (this.specVanish != null)
                this.specVanish.Write(sw, "specVanish");
            if (this.oMath != null)
                this.oMath.Write(sw, "oMath");
            if (this.rPrChange != null)
                this.rPrChange.Write(sw, "rPrChange");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_String rStyle
        {
            get
            {
                return this.rStyleField;
            }
            set
            {
                this.rStyleField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_Fonts rFonts
        {
            get
            {
                return this.rFontsField;
            }
            set
            {
                this.rFontsField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_OnOff b
        {
            get
            {
                return this.bField;
            }
            set
            {
                this.bField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_OnOff bCs
        {
            get
            {
                return this.bCsField;
            }
            set
            {
                this.bCsField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_OnOff i
        {
            get
            {
                return this.iField;
            }
            set
            {
                this.iField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_OnOff iCs
        {
            get
            {
                return this.iCsField;
            }
            set
            {
                this.iCsField = value;
            }
        }

        [XmlElement(Order = 6)]
        public CT_OnOff caps
        {
            get
            {
                return this.capsField;
            }
            set
            {
                this.capsField = value;
            }
        }

        [XmlElement(Order = 7)]
        public CT_OnOff smallCaps
        {
            get
            {
                return this.smallCapsField;
            }
            set
            {
                this.smallCapsField = value;
            }
        }

        [XmlElement(Order = 8)]
        public CT_OnOff strike
        {
            get
            {
                return this.strikeField;
            }
            set
            {
                this.strikeField = value;
            }
        }

        [XmlElement(Order = 9)]
        public CT_OnOff dstrike
        {
            get
            {
                return this.dstrikeField;
            }
            set
            {
                this.dstrikeField = value;
            }
        }

        [XmlElement(Order = 10)]
        public CT_OnOff outline
        {
            get
            {
                return this.outlineField;
            }
            set
            {
                this.outlineField = value;
            }
        }

        [XmlElement(Order = 11)]
        public CT_OnOff shadow
        {
            get
            {
                return this.shadowField;
            }
            set
            {
                this.shadowField = value;
            }
        }

        [XmlElement(Order = 12)]
        public CT_OnOff emboss
        {
            get
            {
                return this.embossField;
            }
            set
            {
                this.embossField = value;
            }
        }

        [XmlElement(Order = 13)]
        public CT_OnOff imprint
        {
            get
            {
                return this.imprintField;
            }
            set
            {
                this.imprintField = value;
            }
        }

        [XmlElement(Order = 14)]
        public CT_OnOff noProof
        {
            get
            {
                return this.noProofField;
            }
            set
            {
                this.noProofField = value;
            }
        }

        [XmlElement(Order = 15)]
        public CT_OnOff snapToGrid
        {
            get
            {
                return this.snapToGridField;
            }
            set
            {
                this.snapToGridField = value;
            }
        }

        [XmlElement(Order = 16)]
        public CT_OnOff vanish
        {
            get
            {
                return this.vanishField;
            }
            set
            {
                this.vanishField = value;
            }
        }

        [XmlElement(Order = 17)]
        public CT_OnOff webHidden
        {
            get
            {
                return this.webHiddenField;
            }
            set
            {
                this.webHiddenField = value;
            }
        }

        [XmlElement(Order = 18)]
        public CT_Color color
        {
            get
            {
                return this.colorField;
            }
            set
            {
                this.colorField = value;
            }
        }

        [XmlElement(Order = 19)]
        public CT_SignedTwipsMeasure spacing
        {
            get
            {
                return this.spacingField;
            }
            set
            {
                this.spacingField = value;
            }
        }

        [XmlElement(Order = 20)]
        public CT_TextScale w
        {
            get
            {
                return this.wField;
            }
            set
            {
                this.wField = value;
            }
        }

        [XmlElement(Order = 21)]
        public CT_HpsMeasure kern
        {
            get
            {
                return this.kernField;
            }
            set
            {
                this.kernField = value;
            }
        }

        [XmlElement(Order = 22)]
        public CT_SignedHpsMeasure position
        {
            get
            {
                return this.positionField;
            }
            set
            {
                this.positionField = value;
            }
        }

        [XmlElement(Order = 23)]
        public CT_HpsMeasure sz
        {
            get
            {
                return this.szField;
            }
            set
            {
                this.szField = value;
            }
        }

        [XmlElement(Order = 24)]
        public CT_HpsMeasure szCs
        {
            get
            {
                return this.szCsField;
            }
            set
            {
                this.szCsField = value;
            }
        }

        [XmlElement(Order = 25)]
        public CT_Highlight highlight
        {
            get
            {
                return this.highlightField;
            }
            set
            {
                this.highlightField = value;
            }
        }

        [XmlElement(Order = 26)]
        public CT_Underline u
        {
            get
            {
                return this.uField;
            }
            set
            {
                this.uField = value;
            }
        }

        [XmlElement(Order = 27)]
        public CT_TextEffect effect
        {
            get
            {
                return this.effectField;
            }
            set
            {
                this.effectField = value;
            }
        }

        [XmlElement(Order = 28)]
        public CT_Border bdr
        {
            get
            {
                return this.bdrField;
            }
            set
            {
                this.bdrField = value;
            }
        }

        [XmlElement(Order = 29)]
        public CT_Shd shd
        {
            get
            {
                return this.shdField;
            }
            set
            {
                this.shdField = value;
            }
        }

        [XmlElement(Order = 30)]
        public CT_FitText fitText
        {
            get
            {
                return this.fitTextField;
            }
            set
            {
                this.fitTextField = value;
            }
        }

        [XmlElement(Order = 31)]
        public CT_VerticalAlignRun vertAlign
        {
            get
            {
                return this.vertAlignField;
            }
            set
            {
                this.vertAlignField = value;
            }
        }

        [XmlElement(Order = 32)]
        public CT_OnOff rtl
        {
            get
            {
                return this.rtlField;
            }
            set
            {
                this.rtlField = value;
            }
        }

        [XmlElement(Order = 33)]
        public CT_OnOff cs
        {
            get
            {
                return this.csField;
            }
            set
            {
                this.csField = value;
            }
        }

        [XmlElement(Order = 34)]
        public CT_Em em
        {
            get
            {
                return this.emField;
            }
            set
            {
                this.emField = value;
            }
        }

        [XmlElement(Order = 35)]
        public CT_Language lang
        {
            get
            {
                return this.langField;
            }
            set
            {
                this.langField = value;
            }
        }

        [XmlElement(Order = 36)]
        public CT_EastAsianLayout eastAsianLayout
        {
            get
            {
                return this.eastAsianLayoutField;
            }
            set
            {
                this.eastAsianLayoutField = value;
            }
        }

        [XmlElement(Order = 37)]
        public CT_OnOff specVanish
        {
            get
            {
                return this.specVanishField;
            }
            set
            {
                this.specVanishField = value;
            }
        }

        [XmlElement(Order = 38)]
        public CT_OnOff oMath
        {
            get
            {
                return this.oMathField;
            }
            set
            {
                this.oMathField = value;
            }
        }

        [XmlElement(Order = 39)]
        public CT_RPrChange rPrChange
        {
            get
            {
                return this.rPrChangeField;
            }
            set
            {
                this.rPrChangeField = value;
            }
        }

        public bool IsSetLang()
        {
            return this.langField != null;// !string.IsNullOrEmpty(this.langField.val);
        }

        public CT_Language AddNewLang()
        {
            if (this.langField == null)
                this.langField = new CT_Language();
            return this.langField;
        }

        public CT_Fonts AddNewRFonts()
        {
            if (this.rFontsField == null)
                this.rFontsField = new CT_Fonts();
            return this.rFontsField;
        }

        public CT_OnOff AddNewB()
        {
            if (this.bField == null)
                this.bField = new CT_OnOff();
            return this.bField;
        }

        public CT_OnOff AddNewBCs()
        {
            if (this.bCsField == null)
                this.bCsField = new CT_OnOff();
            return this.bCsField;
        }

        public CT_Color AddNewColor()
        {
            if (this.colorField == null)
                this.colorField = new CT_Color();
            return this.colorField;
        }

        public CT_HpsMeasure AddNewSz()
        {
            if (this.szField == null)
                this.szField = new CT_HpsMeasure();
            return this.szField;
        }

        public CT_HpsMeasure AddNewSzCs()
        {
            if (this.szCsField == null)
                this.szCsField = new CT_HpsMeasure();
            return this.szCsField;
        }

        public bool IsSetPosition()
        {
            return this.positionField!=null && !string.IsNullOrEmpty(this.positionField.val);
        }

        public CT_SignedHpsMeasure AddNewPosition()
        {
            if (this.positionField == null)
                this.positionField = new CT_SignedHpsMeasure();
            return this.positionField;
        }

        public bool IsSetB()
        {
            return this.bField != null;
        }

        public bool IsSetI()
        {
            return this.iField != null;
        }

        public CT_OnOff AddNewI()
        {
            if (this.iField == null)
                this.iField = new CT_OnOff();
            return this.iField;
        }

        public void AddNewNoProof()
        {
            if (this.noProofField == null)
                this.noProofField = new CT_OnOff();
        }

        public bool IsSetU()
        {
            return this.uField != null 
                && this.uField.val != ST_Underline.none;
        }

        public CT_Underline AddNewU()
        {
            if (this.uField == null)
                this.uField = new CT_Underline();
            return this.uField;
        }

        public bool IsSetStrike()
        {
            return this.strike != null;
        }

        public CT_OnOff AddNewStrike()
        {
            if (this.strikeField == null)
                this.strikeField = new CT_OnOff();
            return this.strikeField;
        }

        public bool IsSetVertAlign()
        {
            return !(this.vertAlignField == null);
        }

        public CT_VerticalAlignRun AddNewVertAlign()
        {
            if (this.vertAlignField == null)
                this.vertAlignField = new CT_VerticalAlignRun();
            return this.vertAlignField;
        }

        public bool IsSetRFonts()
        {
            return this.rFontsField != null;
        }

        public bool IsSetSz()
        {
            return (this.szField!=null && this.szField.val != 0);
        }

        public bool IsSetColor()
        {
            return colorField != null && !string.IsNullOrEmpty(colorField.val);
        }

        public bool IsSetDstrike()
        {
            return this.dstrikeField != null;
        }

        public CT_OnOff AddNewDstrike()
        {
            if (this.dstrikeField == null)
                this.dstrikeField = new CT_OnOff();
            return this.dstrikeField;
        }

        public bool IsSetEmboss()
        {
            return this.embossField != null;
        }

        public CT_OnOff AddNewEmboss()
        {
            if (this.embossField == null)
                this.embossField = new CT_OnOff();
            return this.embossField;
        }

        public bool IsSetImprint()
        {
            return this.imprintField != null;
        }

        public CT_OnOff AddNewImprint()
        {
            if (this.imprintField == null)
                this.imprintField = new CT_OnOff();
            return this.imprintField;
        }

        public bool IsSetShadow()
        {
            return this.shadowField != null;
        }

        public CT_OnOff AddNewShadow()
        {
            if (this.shadowField == null)
                this.shadowField = new CT_OnOff();
            return this.shadowField;
        }

        public bool IsSetCaps()
        {
            return this.capsField != null;
        }

        public CT_OnOff AddNewCaps()
        {
            if (this.capsField == null)
                this.capsField = new CT_OnOff();
            return this.capsField;
        }

        public bool IsSetSmallCaps()
        {
            return this.smallCapsField != null;
        }

        public CT_OnOff AddNewSmallCaps()
        {
            if (this.smallCapsField == null)
                this.smallCapsField = new CT_OnOff();
            return this.smallCapsField;
        }

        public bool IsSetKern()
        {
            return this.kernField != null;
        }

        public CT_HpsMeasure AddNewKern()
        {
            if (this.kernField == null)
                this.kernField = new CT_HpsMeasure();
            return this.kernField;
        }

        public bool IsSetSpacing()
        {
            return this.spacingField != null;
        }

        public CT_SignedTwipsMeasure AddNewSpacing()
        {
            if (this.spacingField == null)
                this.spacingField = new CT_SignedTwipsMeasure();
            return this.spacingField;
        }

        public bool IsSetHighlight()
        {
            return this.highlightField != null;
        }

        internal CT_Highlight AddNewHighlight()
        {
            this.highlightField = new CT_Highlight();
            return this.highlightField;
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_ParaRPrChange : CT_TrackChange
    {
        public static new CT_ParaRPrChange Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ParaRPrChange ctObj = new CT_ParaRPrChange();
            ctObj.author = XmlHelper.ReadString(node.Attributes["author"]);
            ctObj.date = XmlHelper.ReadString(node.Attributes["date"]);
            ctObj.id = XmlHelper.ReadString(node.Attributes["id"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "rPr")
                    ctObj.rPr = CT_ParaRPrOriginal.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal new void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "author", this.author);
            XmlHelper.WriteAttribute(sw, "date", this.date);
            XmlHelper.WriteAttribute(sw, "id", this.id);
            sw.Write(">");
            if (this.rPr != null)
                this.rPr.Write(sw, "rPr");
            sw.Write(string.Format("</{0}>", nodeName));
        }


        private CT_ParaRPrOriginal rPrField;

        public CT_ParaRPrChange()
        {
            //this.rPrField = new CT_ParaRPrOriginal();
        }

        [XmlElement(Order = 0)]
        public CT_ParaRPrOriginal rPr
        {
            get
            {
                return this.rPrField;
            }
            set
            {
                this.rPrField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_ParaRPrOriginal
    {

        private CT_TrackChange insField;

        private CT_TrackChange delField;

        private CT_TrackChange moveFromField;

        private CT_TrackChange moveToField;


        public CT_ParaRPrOriginal()
        {

        }

        [XmlElement(Order = 0)]
        public CT_TrackChange ins
        {
            get
            {
                return this.insField;
            }
            set
            {
                this.insField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_TrackChange del
        {
            get
            {
                return this.delField;
            }
            set
            {
                this.delField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_TrackChange moveFrom
        {
            get
            {
                return this.moveFromField;
            }
            set
            {
                this.moveFromField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_TrackChange moveTo
        {
            get
            {
                return this.moveToField;
            }
            set
            {
                this.moveToField = value;
            }
        }

        List<CT_OnOff> webHiddenField;
        [XmlElement(Order = 4)]
        public List<CT_OnOff> webHidden
        {
            get { return this.webHiddenField; }
            set { this.webHiddenField = value; }
        }

        List<CT_OnOff> bField;
        [XmlElement(Order = 4)]
        public List<CT_OnOff> b
        {
            get { return this.bField; }
            set { this.bField = value; }
        }

        List<CT_OnOff> bCsField;
        [XmlElement(Order = 4)]
        public List<CT_OnOff> bCs
        {
            get { return this.bCsField; }
            set { this.bCsField = value; }
        }

        List<CT_Border> bdrField;
        [XmlElement(Order = 4)]
        public List<CT_Border> bdr
        {
            get { return this.bdrField; }
            set { this.bdrField = value; }
        }

        List<CT_OnOff> capsField;
        [XmlElement(Order = 4)]
        public List<CT_OnOff> caps
        {
            get { return this.capsField; }
            set { this.capsField = value; }
        }

        List<CT_Color> colorField;
        [XmlElement(Order = 4)]
        public List<CT_Color> color
        {
            get { return this.colorField; }
            set { this.colorField = value; }
        }

        List<CT_OnOff> csField;
        [XmlElement(Order = 4)]
        public List<CT_OnOff> cs
        {
            get { return this.csField; }
            set { this.csField = value; }
        }

        List<CT_OnOff> dstrikeField;
        [XmlElement(Order = 4)]
        public List<CT_OnOff> dstrike
        {
            get { return this.dstrikeField; }
            set { this.dstrikeField = value; }
        }

        List<CT_EastAsianLayout> eastAsianLayoutField;
        [XmlElement(Order = 4)]
        public List<CT_EastAsianLayout> eastAsianLayout
        {
            get { return this.eastAsianLayoutField; }
            set { this.eastAsianLayoutField = value; }
        }

        List<CT_TextEffect> effectField;
        [XmlElement(Order = 4)]
        public List<CT_TextEffect> effect
        {
            get { return this.effectField; }
            set { this.effectField = value; }
        }

        List<CT_Em> emField;
        [XmlElement(Order = 4)]
        public List<CT_Em> em
        {
            get { return this.emField; }
            set { this.emField = value; }
        }

        List<CT_OnOff> embossField;
        [XmlElement(Order = 4)]
        public List<CT_OnOff> emboss
        {
            get { return this.embossField; }
            set { this.embossField = value; }
        }

        List<CT_FitText> fitTextField;
        [XmlElement(Order = 4)]
        public List<CT_FitText> fitText
        {
            get { return this.fitTextField; }
            set { this.fitTextField = value; }
        }

        List<CT_Highlight> highlightField;
        [XmlElement(Order = 4)]
        public List<CT_Highlight> highlight
        {
            get { return this.highlightField; }
            set { this.highlightField = value; }
        }

        List<CT_OnOff> iField;
        [XmlElement(Order = 4)]
        public List<CT_OnOff> i
        {
            get { return this.iField; }
            set { this.iField = value; }
        }

        List<CT_OnOff> iCsField;
        [XmlElement(Order = 4)]
        public List<CT_OnOff> iCs
        {
            get { return this.iCsField; }
            set { this.iCsField = value; }
        }

        List<CT_OnOff> imprintField;
        [XmlElement(Order = 4)]
        public List<CT_OnOff> imprint
        {
            get { return this.imprintField; }
            set { this.imprintField = value; }
        }

        List<CT_HpsMeasure> kernField;
        [XmlElement(Order = 4)]
        public List<CT_HpsMeasure> kern
        {
            get { return this.kernField; }
            set { this.kernField = value; }
        }

        List<CT_Language> langField;
        [XmlElement(Order = 4)]
        public List<CT_Language> lang
        {
            get { return this.langField; }
            set { this.langField = value; }
        }

        List<CT_OnOff> noProofField;
        [XmlElement(Order = 4)]
        public List<CT_OnOff> noProof
        {
            get { return this.noProofField; }
            set { this.noProofField = value; }
        }

        List<CT_OnOff> oMathField;
        [XmlElement(Order = 4)]
        public List<CT_OnOff> oMath
        {
            get { return this.oMathField; }
            set { this.oMathField = value; }
        }

        List<CT_OnOff> outlineField;
        [XmlElement(Order = 4)]
        public List<CT_OnOff> outline
        {
            get { return this.outlineField; }
            set { this.outlineField = value; }
        }

        List<CT_SignedHpsMeasure> positionField;
        [XmlElement(Order = 4)]
        public List<CT_SignedHpsMeasure> position
        {
            get { return this.positionField; }
            set { this.positionField = value; }
        }

        List<CT_Fonts> rFontsField;
        [XmlElement(Order = 4)]
        public List<CT_Fonts> rFonts
        {
            get { return this.rFontsField; }
            set { this.rFontsField = value; }
        }

        List<CT_String> rStyleField;
        [XmlElement(Order = 4)]
        public List<CT_String> rStyle
        {
            get { return this.rStyleField; }
            set { this.rStyleField = value; }
        }

        List<CT_OnOff> rtlField;
        [XmlElement(Order = 4)]
        public List<CT_OnOff> rtl
        {
            get { return this.rtlField; }
            set { this.rtlField = value; }
        }

        List<CT_OnOff> shadowField;
        [XmlElement(Order = 4)]
        public List<CT_OnOff> shadow
        {
            get { return this.shadowField; }
            set { this.shadowField = value; }
        }

        List<CT_Shd> shdField;
        [XmlElement(Order = 4)]
        public List<CT_Shd> shd
        {
            get { return this.shdField; }
            set { this.shdField = value; }
        }

        List<CT_OnOff> smallCapsField;
        [XmlElement(Order = 4)]
        public List<CT_OnOff> smallCaps
        {
            get { return this.smallCapsField; }
            set { this.smallCapsField = value; }
        }

        List<CT_OnOff> snapToGridField;
        [XmlElement(Order = 4)]
        public List<CT_OnOff> snapToGrid
        {
            get { return this.snapToGridField; }
            set { this.snapToGridField = value; }
        }

        List<CT_SignedTwipsMeasure> spacingField;
        [XmlElement(Order = 4)]
        public List<CT_SignedTwipsMeasure> spacing
        {
            get { return this.spacingField; }
            set { this.spacingField = value; }
        }

        List<CT_OnOff> specVanishField;
        [XmlElement(Order = 4)]
        public List<CT_OnOff> specVanish
        {
            get { return this.specVanishField; }
            set { this.specVanishField = value; }
        }

        List<CT_OnOff> strikeField;
        [XmlElement(Order = 4)]
        public List<CT_OnOff> strike
        {
            get { return this.strikeField; }
            set { this.strikeField = value; }
        }

        List<CT_HpsMeasure> szField;
        [XmlElement(Order = 4)]
        public List<CT_HpsMeasure> sz
        {
            get { return this.szField; }
            set { this.szField = value; }
        }

        List<CT_HpsMeasure> szCsField;
        [XmlElement(Order = 4)]
        public List<CT_HpsMeasure> szCs
        {
            get { return this.szCsField; }
            set { this.szCsField = value; }
        }

        List<CT_Underline> uField;
        [XmlElement(Order = 4)]
        public List<CT_Underline> u
        {
            get { return this.uField; }
            set { this.uField = value; }
        }

        List<CT_OnOff> vanishField;
        [XmlElement(Order = 4)]
        public List<CT_OnOff> vanish
        {
            get { return this.vanishField; }
            set { this.vanishField = value; }
        }

        List<CT_VerticalAlignRun> vertAlignField;
        [XmlElement(Order = 4)]
        public List<CT_VerticalAlignRun> vertAlign
        {
            get { return this.vertAlignField; }
            set { this.vertAlignField = value; }
        }

        List<CT_TextScale> wField;
        [XmlElement(Order = 4)]
        public List<CT_TextScale> w
        {
            get { return this.wField; }
            set { this.wField = value; }
        }


        public static CT_ParaRPrOriginal Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ParaRPrOriginal ctObj = new CT_ParaRPrOriginal();
            ctObj.webHidden = new List<CT_OnOff>();
            ctObj.b = new List<CT_OnOff>();
            ctObj.bCs = new List<CT_OnOff>();
            ctObj.bdr = new List<CT_Border>();
            ctObj.caps = new List<CT_OnOff>();
            ctObj.color = new List<CT_Color>();
            ctObj.cs = new List<CT_OnOff>();
            ctObj.dstrike = new List<CT_OnOff>();
            ctObj.eastAsianLayout = new List<CT_EastAsianLayout>();
            ctObj.effect = new List<CT_TextEffect>();
            ctObj.em = new List<CT_Em>();
            ctObj.emboss = new List<CT_OnOff>();
            ctObj.fitText = new List<CT_FitText>();
            ctObj.highlight = new List<CT_Highlight>();
            ctObj.i = new List<CT_OnOff>();
            ctObj.iCs = new List<CT_OnOff>();
            ctObj.imprint = new List<CT_OnOff>();
            ctObj.kern = new List<CT_HpsMeasure>();
            ctObj.lang = new List<CT_Language>();
            ctObj.noProof = new List<CT_OnOff>();
            ctObj.oMath = new List<CT_OnOff>();
            ctObj.outline = new List<CT_OnOff>();
            ctObj.position = new List<CT_SignedHpsMeasure>();
            ctObj.rFonts = new List<CT_Fonts>();
            ctObj.rStyle = new List<CT_String>();
            ctObj.rtl = new List<CT_OnOff>();
            ctObj.shadow = new List<CT_OnOff>();
            ctObj.shd = new List<CT_Shd>();
            ctObj.smallCaps = new List<CT_OnOff>();
            ctObj.snapToGrid = new List<CT_OnOff>();
            ctObj.spacing = new List<CT_SignedTwipsMeasure>();
            ctObj.specVanish = new List<CT_OnOff>();
            ctObj.strike = new List<CT_OnOff>();
            ctObj.sz = new List<CT_HpsMeasure>();
            ctObj.szCs = new List<CT_HpsMeasure>();
            ctObj.u = new List<CT_Underline>();
            ctObj.vanish = new List<CT_OnOff>();
            ctObj.vertAlign = new List<CT_VerticalAlignRun>();
            ctObj.w = new List<CT_TextScale>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "ins")
                    ctObj.ins = CT_TrackChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "del")
                    ctObj.del = CT_TrackChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "moveFrom")
                    ctObj.moveFrom = CT_TrackChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "moveTo")
                    ctObj.moveTo = CT_TrackChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "webHidden")
                    ctObj.webHidden.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "b")
                    ctObj.b.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "bCs")
                    ctObj.bCs.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "bdr")
                    ctObj.bdr.Add(CT_Border.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "caps")
                    ctObj.caps.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "color")
                    ctObj.color.Add(CT_Color.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "cs")
                    ctObj.cs.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "dstrike")
                    ctObj.dstrike.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "eastAsianLayout")
                    ctObj.eastAsianLayout.Add(CT_EastAsianLayout.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "effect")
                    ctObj.effect.Add(CT_TextEffect.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "em")
                    ctObj.em.Add(CT_Em.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "emboss")
                    ctObj.emboss.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "fitText")
                    ctObj.fitText.Add(CT_FitText.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "highlight")
                    ctObj.highlight.Add(CT_Highlight.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "i")
                    ctObj.i.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "iCs")
                    ctObj.iCs.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "imprint")
                    ctObj.imprint.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "kern")
                    ctObj.kern.Add(CT_HpsMeasure.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "lang")
                    ctObj.lang.Add(CT_Language.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "noProof")
                    ctObj.noProof.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "oMath")
                    ctObj.oMath.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "outline")
                    ctObj.outline.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "position")
                    ctObj.position.Add(CT_SignedHpsMeasure.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "rFonts")
                    ctObj.rFonts.Add(CT_Fonts.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "rStyle")
                    ctObj.rStyle.Add(CT_String.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "rtl")
                    ctObj.rtl.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "shadow")
                    ctObj.shadow.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "shd")
                    ctObj.shd.Add(CT_Shd.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "smallCaps")
                    ctObj.smallCaps.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "snapToGrid")
                    ctObj.snapToGrid.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "spacing")
                    ctObj.spacing.Add(CT_SignedTwipsMeasure.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "specVanish")
                    ctObj.specVanish.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "strike")
                    ctObj.strike.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "sz")
                    ctObj.sz.Add(CT_HpsMeasure.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "szCs")
                    ctObj.szCs.Add(CT_HpsMeasure.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "u")
                    ctObj.u.Add(CT_Underline.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "vanish")
                    ctObj.vanish.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "vertAlign")
                    ctObj.vertAlign.Add(CT_VerticalAlignRun.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "w")
                    ctObj.w.Add(CT_TextScale.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}>", nodeName));
            if (this.ins != null)
                this.ins.Write(sw, "ins");
            if (this.del != null)
                this.del.Write(sw, "del");
            if (this.moveFrom != null)
                this.moveFrom.Write(sw, "moveFrom");
            if (this.moveTo != null)
                this.moveTo.Write(sw, "moveTo");
            if (this.webHidden != null)
            {
                foreach (CT_OnOff x in this.webHidden)
                {
                    x.Write(sw, "webHidden");
                }
            }
            if (this.b != null)
            {
                foreach (CT_OnOff x in this.b)
                {
                    x.Write(sw, "b");
                }
            }
            if (this.bCs != null)
            {
                foreach (CT_OnOff x in this.bCs)
                {
                    x.Write(sw, "bCs");
                }
            }
            if (this.bdr != null)
            {
                foreach (CT_Border x in this.bdr)
                {
                    x.Write(sw, "bdr");
                }
            }
            if (this.caps != null)
            {
                foreach (CT_OnOff x in this.caps)
                {
                    x.Write(sw, "caps");
                }
            }
            if (this.color != null)
            {
                foreach (CT_Color x in this.color)
                {
                    x.Write(sw, "color");
                }
            }
            if (this.cs != null)
            {
                foreach (CT_OnOff x in this.cs)
                {
                    x.Write(sw, "cs");
                }
            }
            if (this.dstrike != null)
            {
                foreach (CT_OnOff x in this.dstrike)
                {
                    x.Write(sw, "dstrike");
                }
            }
            if (this.eastAsianLayout != null)
            {
                foreach (CT_EastAsianLayout x in this.eastAsianLayout)
                {
                    x.Write(sw, "eastAsianLayout");
                }
            }
            if (this.effect != null)
            {
                foreach (CT_TextEffect x in this.effect)
                {
                    x.Write(sw, "effect");
                }
            }
            if (this.em != null)
            {
                foreach (CT_Em x in this.em)
                {
                    x.Write(sw, "em");
                }
            }
            if (this.emboss != null)
            {
                foreach (CT_OnOff x in this.emboss)
                {
                    x.Write(sw, "emboss");
                }
            }
            if (this.fitText != null)
            {
                foreach (CT_FitText x in this.fitText)
                {
                    x.Write(sw, "fitText");
                }
            }
            if (this.highlight != null)
            {
                foreach (CT_Highlight x in this.highlight)
                {
                    x.Write(sw, "highlight");
                }
            }
            if (this.i != null)
            {
                foreach (CT_OnOff x in this.i)
                {
                    x.Write(sw, "i");
                }
            }
            if (this.iCs != null)
            {
                foreach (CT_OnOff x in this.iCs)
                {
                    x.Write(sw, "iCs");
                }
            }
            if (this.imprint != null)
            {
                foreach (CT_OnOff x in this.imprint)
                {
                    x.Write(sw, "imprint");
                }
            }
            if (this.kern != null)
            {
                foreach (CT_HpsMeasure x in this.kern)
                {
                    x.Write(sw, "kern");
                }
            }
            if (this.lang != null)
            {
                foreach (CT_Language x in this.lang)
                {
                    x.Write(sw, "lang");
                }
            }
            if (this.noProof != null)
            {
                foreach (CT_OnOff x in this.noProof)
                {
                    x.Write(sw, "noProof");
                }
            }
            if (this.oMath != null)
            {
                foreach (CT_OnOff x in this.oMath)
                {
                    x.Write(sw, "oMath");
                }
            }
            if (this.outline != null)
            {
                foreach (CT_OnOff x in this.outline)
                {
                    x.Write(sw, "outline");
                }
            }
            if (this.position != null)
            {
                foreach (CT_SignedHpsMeasure x in this.position)
                {
                    x.Write(sw, "position");
                }
            }
            if (this.rFonts != null)
            {
                foreach (CT_Fonts x in this.rFonts)
                {
                    x.Write(sw, "rFonts");
                }
            }
            if (this.rStyle != null)
            {
                foreach (CT_String x in this.rStyle)
                {
                    x.Write(sw, "rStyle");
                }
            }
            if (this.rtl != null)
            {
                foreach (CT_OnOff x in this.rtl)
                {
                    x.Write(sw, "rtl");
                }
            }
            if (this.shadow != null)
            {
                foreach (CT_OnOff x in this.shadow)
                {
                    x.Write(sw, "shadow");
                }
            }
            if (this.shd != null)
            {
                foreach (CT_Shd x in this.shd)
                {
                    x.Write(sw, "shd");
                }
            }
            if (this.smallCaps != null)
            {
                foreach (CT_OnOff x in this.smallCaps)
                {
                    x.Write(sw, "smallCaps");
                }
            }
            if (this.snapToGrid != null)
            {
                foreach (CT_OnOff x in this.snapToGrid)
                {
                    x.Write(sw, "snapToGrid");
                }
            }
            if (this.spacing != null)
            {
                foreach (CT_SignedTwipsMeasure x in this.spacing)
                {
                    x.Write(sw, "spacing");
                }
            }
            if (this.specVanish != null)
            {
                foreach (CT_OnOff x in this.specVanish)
                {
                    x.Write(sw, "specVanish");
                }
            }
            if (this.strike != null)
            {
                foreach (CT_OnOff x in this.strike)
                {
                    x.Write(sw, "strike");
                }
            }
            if (this.sz != null)
            {
                foreach (CT_HpsMeasure x in this.sz)
                {
                    x.Write(sw, "sz");
                }
            }
            if (this.szCs != null)
            {
                foreach (CT_HpsMeasure x in this.szCs)
                {
                    x.Write(sw, "szCs");
                }
            }
            if (this.u != null)
            {
                foreach (CT_Underline x in this.u)
                {
                    x.Write(sw, "u");
                }
            }
            if (this.vanish != null)
            {
                foreach (CT_OnOff x in this.vanish)
                {
                    x.Write(sw, "vanish");
                }
            }
            if (this.vertAlign != null)
            {
                foreach (CT_VerticalAlignRun x in this.vertAlign)
                {
                    x.Write(sw, "vertAlign");
                }
            }
            if (this.w != null)
            {
                foreach (CT_TextScale x in this.w)
                {
                    x.Write(sw, "w");
                }
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType3
    {

    
        b,

    
        bCs,

    
        bdr,

    
        caps,

    
        color,

    
        cs,

    
        dstrike,

    
        eastAsianLayout,

    
        effect,

    
        em,

    
        emboss,

    
        fitText,

    
        highlight,

    
        i,

    
        iCs,

    
        imprint,

    
        kern,

    
        lang,

    
        noProof,

    
        oMath,

    
        outline,

    
        position,

    
        rFonts,

    
        rStyle,

    
        rtl,

    
        shadow,

    
        shd,

    
        smallCaps,

    
        snapToGrid,

    
        spacing,

    
        specVanish,

    
        strike,

    
        sz,

    
        szCs,

    
        u,

    
        vanish,

    
        vertAlign,

    
        w,

    
        webHidden,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_PPrChange : CT_TrackChange
    {

        private CT_PPrBase pPrField;

        public CT_PPrChange()
        {
            //this.pPrField = new CT_PPrBase();
        }
        public static new CT_PPrChange Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PPrChange ctObj = new CT_PPrChange();
            ctObj.author = XmlHelper.ReadString(node.Attributes["w:author"]);
            ctObj.date = XmlHelper.ReadString(node.Attributes["w:date"]);
            ctObj.id = XmlHelper.ReadString(node.Attributes["r:id"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "pPr")
                    ctObj.pPr = CT_PPrBase.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal new void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:author", this.author);
            XmlHelper.WriteAttribute(sw, "w:date", this.date);
            XmlHelper.WriteAttribute(sw, "r:id", this.id);
            sw.Write(">");
            if (this.pPr != null)
                this.pPr.Write(sw, "pPr");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }



        [XmlElement(Order = 0)]
        public CT_PPrBase pPr
        {
            get
            {
                return this.pPrField;
            }
            set
            {
                this.pPrField = value;
            }
        }
    }

    [XmlInclude(typeof(CT_PPr))]

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_PPrBase
    {

        private CT_String pStyleField;

        private CT_OnOff keepNextField;

        private CT_OnOff keepLinesField;

        private CT_OnOff pageBreakBeforeField;

        private CT_FramePr framePrField;

        private CT_OnOff widowControlField;

        private CT_NumPr numPrField;

        private CT_OnOff suppressLineNumbersField;

        private CT_PBdr pBdrField;

        private CT_Shd shdField;

        private List<CT_TabStop> tabsField;

        private CT_OnOff suppressAutoHyphensField;

        private CT_OnOff kinsokuField;

        private CT_OnOff wordWrapField;

        private CT_OnOff overflowPunctField;

        private CT_OnOff topLinePunctField;

        private CT_OnOff autoSpaceDEField;

        private CT_OnOff autoSpaceDNField;

        private CT_OnOff bidiField;

        private CT_OnOff adjustRightIndField;

        private CT_OnOff snapToGridField;

        private CT_Spacing spacingField;

        private CT_Ind indField;

        private CT_OnOff contextualSpacingField;

        private CT_OnOff mirrorIndentsField;

        private CT_OnOff suppressOverlapField;

        private CT_Jc jcField;

        private CT_TextDirection textDirectionField;

        private CT_TextAlignment textAlignmentField;

        private CT_TextboxTightWrap textboxTightWrapField;

        private CT_DecimalNumber outlineLvlField;

        private CT_DecimalNumber divIdField;

        private CT_Cnf cnfStyleField;

        public virtual bool IsEmpty
        {
            get
            {
                return pStyleField == null &&
keepNextField == null &&
keepLinesField == null &&
pageBreakBeforeField == null &&
framePrField == null &&
widowControlField == null &&
numPrField == null &&
suppressLineNumbersField == null &&
pBdrField == null &&
shdField == null &&
tabsField == null &&
suppressAutoHyphensField == null &&
kinsokuField == null &&
wordWrapField == null &&
overflowPunctField == null &&
topLinePunctField == null &&
autoSpaceDEField == null &&
autoSpaceDNField == null &&
bidiField == null &&
adjustRightIndField == null &&
snapToGridField == null &&
spacingField == null &&
indField == null &&
contextualSpacingField == null &&
mirrorIndentsField == null &&
suppressOverlapField == null &&
jcField == null &&
textDirectionField == null &&
textAlignmentField == null &&
textboxTightWrapField == null &&
outlineLvlField == null &&
divIdField == null &&
cnfStyleField == null;
            }
        }

        public CT_PPrBase()
        {

        }
        public static CT_PPrBase Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PPrBase ctObj = new CT_PPrBase();
            ctObj.tabs = new List<CT_TabStop>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "pStyle")
                    ctObj.pStyle = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "keepNext")
                    ctObj.keepNext = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "keepLines")
                    ctObj.keepLines = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pageBreakBefore")
                    ctObj.pageBreakBefore = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "framePr")
                    ctObj.framePr = CT_FramePr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "widowControl")
                    ctObj.widowControl = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "numPr")
                    ctObj.numPr = CT_NumPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "suppressLineNumbers")
                    ctObj.suppressLineNumbers = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pBdr")
                    ctObj.pBdr = CT_PBdr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "shd")
                    ctObj.shd = CT_Shd.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "suppressAutoHyphens")
                    ctObj.suppressAutoHyphens = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "kinsoku")
                    ctObj.kinsoku = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "wordWrap")
                    ctObj.wordWrap = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "overflowPunct")
                    ctObj.overflowPunct = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "topLinePunct")
                    ctObj.topLinePunct = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "autoSpaceDE")
                    ctObj.autoSpaceDE = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "autoSpaceDN")
                    ctObj.autoSpaceDN = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bidi")
                    ctObj.bidi = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "adjustRightInd")
                    ctObj.adjustRightInd = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "snapToGrid")
                    ctObj.snapToGrid = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "spacing")
                    ctObj.spacing = CT_Spacing.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "ind")
                    ctObj.ind = CT_Ind.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "contextualSpacing")
                    ctObj.contextualSpacing = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "mirrorIndents")
                    ctObj.mirrorIndents = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "suppressOverlap")
                    ctObj.suppressOverlap = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "jc")
                    ctObj.jc = CT_Jc.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "textDirection")
                    ctObj.textDirection = CT_TextDirection.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "textAlignment")
                    ctObj.textAlignment = CT_TextAlignment.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "textboxTightWrap")
                    ctObj.textboxTightWrap = CT_TextboxTightWrap.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "outlineLvl")
                    ctObj.outlineLvl = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "divId")
                    ctObj.divId = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cnfStyle")
                    ctObj.cnfStyle = CT_Cnf.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tabs")
                {
                    foreach (XmlNode snode in childNode.ChildNodes)
                    {
                        ctObj.tabs.Add(CT_TabStop.Parse(snode, namespaceManager));
                    }
                }
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.pStyle != null)
                this.pStyle.Write(sw, "pStyle");
            if (this.tabs != null && this.tabs.Count > 0)
            {
                sw.Write("<w:tabs>");
                foreach (CT_TabStop x in this.tabs)
                {
                    x.Write(sw, "tab");
                }
                sw.Write("</w:tabs>");
            }
            if (this.keepNext != null)
                this.keepNext.Write(sw, "keepNext");
            if (this.keepLines != null)
                this.keepLines.Write(sw, "keepLines");
            if (this.pageBreakBefore != null)
                this.pageBreakBefore.Write(sw, "pageBreakBefore");
            if (this.framePr != null)
                this.framePr.Write(sw, "framePr");
            if (this.widowControl != null)
                this.widowControl.Write(sw, "widowControl");
            if (this.numPr != null)
                this.numPr.Write(sw, "numPr");
            if (this.suppressLineNumbers != null)
                this.suppressLineNumbers.Write(sw, "suppressLineNumbers");
            if (this.pBdr != null)
                this.pBdr.Write(sw, "pBdr");
            if (this.shd != null)
                this.shd.Write(sw, "shd");
            if (this.suppressAutoHyphens != null)
                this.suppressAutoHyphens.Write(sw, "suppressAutoHyphens");
            if (this.kinsoku != null)
                this.kinsoku.Write(sw, "kinsoku");
            if (this.wordWrap != null)
                this.wordWrap.Write(sw, "wordWrap");
            if (this.overflowPunct != null)
                this.overflowPunct.Write(sw, "overflowPunct");
            if (this.topLinePunct != null)
                this.topLinePunct.Write(sw, "topLinePunct");
            if (this.autoSpaceDE != null)
                this.autoSpaceDE.Write(sw, "autoSpaceDE");
            if (this.autoSpaceDN != null)
                this.autoSpaceDN.Write(sw, "autoSpaceDN");
            if (this.bidi != null)
                this.bidi.Write(sw, "bidi");
            if (this.adjustRightInd != null)
                this.adjustRightInd.Write(sw, "adjustRightInd");
            if (this.snapToGrid != null)
                this.snapToGrid.Write(sw, "snapToGrid");
            if (this.spacing != null)
                this.spacing.Write(sw, "spacing");
            if (this.ind != null)
                this.ind.Write(sw, "ind");
            if (this.contextualSpacing != null)
                this.contextualSpacing.Write(sw, "contextualSpacing");
            if (this.mirrorIndents != null)
                this.mirrorIndents.Write(sw, "mirrorIndents");
            if (this.suppressOverlap != null)
                this.suppressOverlap.Write(sw, "suppressOverlap");
            if (this.jc != null)
                this.jc.Write(sw, "jc");
            if (this.textDirection != null)
                this.textDirection.Write(sw, "textDirection");
            if (this.textAlignment != null)
                this.textAlignment.Write(sw, "textAlignment");
            if (this.textboxTightWrap != null)
                this.textboxTightWrap.Write(sw, "textboxTightWrap");
            if (this.outlineLvl != null)
                this.outlineLvl.Write(sw, "outlineLvl");
            if (this.divId != null)
                this.divId.Write(sw, "divId");
            if (this.cnfStyle != null)
                this.cnfStyle.Write(sw, "cnfStyle");

            sw.Write(string.Format("</w:{0}>", nodeName));
        }


        [XmlElement(Order = 0)]
        public CT_String pStyle
        {
            get
            {
                return this.pStyleField;
            }
            set
            {
                this.pStyleField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_OnOff keepNext
        {
            get
            {
                return this.keepNextField;
            }
            set
            {
                this.keepNextField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_OnOff keepLines
        {
            get
            {
                return this.keepLinesField;
            }
            set
            {
                this.keepLinesField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_OnOff pageBreakBefore
        {
            get
            {
                return this.pageBreakBeforeField;
            }
            set
            {
                this.pageBreakBeforeField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_FramePr framePr
        {
            get
            {
                return this.framePrField;
            }
            set
            {
                this.framePrField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_OnOff widowControl
        {
            get
            {
                return this.widowControlField;
            }
            set
            {
                this.widowControlField = value;
            }
        }

        [XmlElement(Order = 6)]
        public CT_NumPr numPr
        {
            get
            {
                return this.numPrField;
            }
            set
            {
                this.numPrField = value;
            }
        }

        [XmlElement(Order = 7)]
        public CT_OnOff suppressLineNumbers
        {
            get
            {
                return this.suppressLineNumbersField;
            }
            set
            {
                this.suppressLineNumbersField = value;
            }
        }

        [XmlElement(Order = 8)]
        public CT_PBdr pBdr
        {
            get
            {
                return this.pBdrField;
            }
            set
            {
                this.pBdrField = value;
            }
        }

        [XmlElement(Order = 9)]
        public CT_Shd shd
        {
            get
            {
                return this.shdField;
            }
            set
            {
                this.shdField = value;
            }
        }

        [XmlArray(Order = 10)]
        [XmlArrayItem("tab", IsNullable = false)]
        public List<CT_TabStop> tabs
        {
            get
            {
                return this.tabsField;
            }
            set
            {
                this.tabsField = value;
            }
        }

        [XmlElement(Order = 11)]
        public CT_OnOff suppressAutoHyphens
        {
            get
            {
                return this.suppressAutoHyphensField;
            }
            set
            {
                this.suppressAutoHyphensField = value;
            }
        }

        [XmlElement(Order = 12)]
        public CT_OnOff kinsoku
        {
            get
            {
                return this.kinsokuField;
            }
            set
            {
                this.kinsokuField = value;
            }
        }

        [XmlElement(Order = 13)]
        public CT_OnOff wordWrap
        {
            get
            {
                return this.wordWrapField;
            }
            set
            {
                this.wordWrapField = value;
            }
        }

        [XmlElement(Order = 14)]
        public CT_OnOff overflowPunct
        {
            get
            {
                return this.overflowPunctField;
            }
            set
            {
                this.overflowPunctField = value;
            }
        }

        [XmlElement(Order = 15)]
        public CT_OnOff topLinePunct
        {
            get
            {
                return this.topLinePunctField;
            }
            set
            {
                this.topLinePunctField = value;
            }
        }

        [XmlElement(Order = 16)]
        public CT_OnOff autoSpaceDE
        {
            get
            {
                return this.autoSpaceDEField;
            }
            set
            {
                this.autoSpaceDEField = value;
            }
        }

        [XmlElement(Order = 17)]
        public CT_OnOff autoSpaceDN
        {
            get
            {
                return this.autoSpaceDNField;
            }
            set
            {
                this.autoSpaceDNField = value;
            }
        }

        [XmlElement(Order = 18)]
        public CT_OnOff bidi
        {
            get
            {
                return this.bidiField;
            }
            set
            {
                this.bidiField = value;
            }
        }

        [XmlElement(Order = 19)]
        public CT_OnOff adjustRightInd
        {
            get
            {
                return this.adjustRightIndField;
            }
            set
            {
                this.adjustRightIndField = value;
            }
        }

        [XmlElement(Order = 20)]
        public CT_OnOff snapToGrid
        {
            get
            {
                return this.snapToGridField;
            }
            set
            {
                this.snapToGridField = value;
            }
        }

        [XmlElement(Order = 21)]
        public CT_Spacing spacing
        {
            get
            {
                return this.spacingField;
            }
            set
            {
                this.spacingField = value;
            }
        }

        [XmlElement(Order = 22)]
        public CT_Ind ind
        {
            get
            {
                return this.indField;
            }
            set
            {
                this.indField = value;
            }
        }

        [XmlElement(Order = 23)]
        public CT_OnOff contextualSpacing
        {
            get
            {
                return this.contextualSpacingField;
            }
            set
            {
                this.contextualSpacingField = value;
            }
        }

        [XmlElement(Order = 24)]
        public CT_OnOff mirrorIndents
        {
            get
            {
                return this.mirrorIndentsField;
            }
            set
            {
                this.mirrorIndentsField = value;
            }
        }

        [XmlElement(Order = 25)]
        public CT_OnOff suppressOverlap
        {
            get
            {
                return this.suppressOverlapField;
            }
            set
            {
                this.suppressOverlapField = value;
            }
        }

        [XmlElement(Order = 26)]
        public CT_Jc jc
        {
            get
            {
                return this.jcField;
            }
            set
            {
                this.jcField = value;
            }
        }

        [XmlElement(Order = 27)]
        public CT_TextDirection textDirection
        {
            get
            {
                return this.textDirectionField;
            }
            set
            {
                this.textDirectionField = value;
            }
        }

        [XmlElement(Order = 28)]
        public CT_TextAlignment textAlignment
        {
            get
            {
                return this.textAlignmentField;
            }
            set
            {
                this.textAlignmentField = value;
            }
        }

        [XmlElement(Order = 29)]
        public CT_TextboxTightWrap textboxTightWrap
        {
            get
            {
                return this.textboxTightWrapField;
            }
            set
            {
                this.textboxTightWrapField = value;
            }
        }

        [XmlElement(Order = 30)]
        public CT_DecimalNumber outlineLvl
        {
            get
            {
                return this.outlineLvlField;
            }
            set
            {
                this.outlineLvlField = value;
            }
        }

        [XmlElement(Order = 31)]
        public CT_DecimalNumber divId
        {
            get
            {
                return this.divIdField;
            }
            set
            {
                this.divIdField = value;
            }
        }

        [XmlElement(Order = 32)]
        public CT_Cnf cnfStyle
        {
            get
            {
                return this.cnfStyleField;
            }
            set
            {
                this.cnfStyleField = value;
            }
        }
        public bool IsSetTextAlignment()
        {
            if (this.textAlignmentField == null)
                return false;
            return this.textAlignmentField != null;
        }

        public CT_TextAlignment AddNewTextAlignment()
        {
            if (this.textAlignmentField == null)
                this.textAlignmentField = new CT_TextAlignment();
            return this.textAlignmentField;
        }

        public bool IsSetPStyle()
        {
            if (this.pStyleField == null)
                return false;
            return !string.IsNullOrEmpty(this.pStyleField.val);
        }

        public CT_String AddNewPStyle()
        {
            if (this.pStyleField == null)
                this.pStyleField = new CT_String();
            return pStyleField;
        }

        public bool IsSetShd()
        {
            return this.shdField != null;
        }
        public CT_Shd AddNewShd()
        {
            if (this.shdField == null)
            {
                this.shdField = new CT_Shd();
            }
            return this.shdField;
        }
        public bool IsSetJc()
        {
            return this.jcField != null;
        }

        public CT_Jc AddNewJc()
        {
            if (this.jcField == null)
            {
                this.jcField = new CT_Jc();
            }
            return this.jcField;
        }



        public bool IsSetPBdr()
        {
            return this.pBdrField != null;
        }

        public CT_Spacing AddNewSpacing()
        {
            if (this.spacingField == null)
                this.spacingField = new CT_Spacing();
            return this.spacingField;
        }

        public bool IsSetPageBreakBefore()
        {
            if (pageBreakBeforeField == null)
                return false;
            return this.pageBreakBeforeField.val == true;
        }

        public CT_OnOff AddNewPageBreakBefore()
        {
            if (this.pageBreakBeforeField == null)
                this.pageBreakBeforeField = new CT_OnOff();
            return this.pageBreakBeforeField;
        }

        public CT_PBdr AddNewPBdr()
        {
            if (this.pBdrField == null)
                this.pBdrField = new CT_PBdr();
            return this.pBdrField;
        }

        public bool IsSetWordWrap()
        {
            if (wordWrapField == null)
                return false;
            return this.wordWrapField.val == true;
        }

        public CT_OnOff AddNewWordWrap()
        {
            if (this.wordWrapField == null)
                this.wordWrapField = new CT_OnOff();
            return this.wordWrapField;
        }

        public CT_Ind AddNewInd()
        {
            if (this.indField == null)
                this.indField = new CT_Ind();
            return this.indField;
        }

        public CT_Tabs AddNewTabs()
        {
            CT_Tabs tab = new CT_Tabs();
            this.tabsField = tab.tab;

            return tab;
        }
    }
    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_RPrChange : CT_TrackChange
    {

        private CT_RPrOriginal rPrField;

        public CT_RPrChange()
        {
            //this.rPrField = new CT_RPrOriginal();
        }
        public static new CT_RPrChange Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_RPrChange ctObj = new CT_RPrChange();
            ctObj.author = XmlHelper.ReadString(node.Attributes["w:author"]);
            ctObj.date = XmlHelper.ReadString(node.Attributes["w:date"]);
            ctObj.id = XmlHelper.ReadString(node.Attributes["r:id"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "rPr")
                    ctObj.rPr = CT_RPrOriginal.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal new void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:author", this.author);
            XmlHelper.WriteAttribute(sw, "w:date", this.date);
            XmlHelper.WriteAttribute(sw, "r:id", this.id);
            sw.Write(">");
            if (this.rPr != null)
                this.rPr.Write(sw, "rPr");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }


        [XmlElement(Order = 0)]
        public CT_RPrOriginal rPr
        {
            get
            {
                return this.rPrField;
            }
            set
            {
                this.rPrField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_RPrOriginal
    {

        public CT_RPrOriginal()
        {
        }

        List<CT_SignedTwipsMeasure> spacingField;
        public List<CT_SignedTwipsMeasure> spacing
        {
            get { return this.spacingField; }
            set { this.spacingField = value; }
        }

        List<CT_VerticalAlignRun> vertAlignField;
        public List<CT_VerticalAlignRun> vertAlign
        {
            get { return this.vertAlignField; }
            set { this.vertAlignField = value; }
        }

        List<CT_TextScale> wField;
        public List<CT_TextScale> w
        {
            get { return this.wField; }
            set { this.wField = value; }
        }

        List<CT_OnOff> noProofField;
        public List<CT_OnOff> noProof
        {
            get { return this.noProofField; }
            set { this.noProofField = value; }
        }

        List<CT_OnOff> snapToGridField;
        public List<CT_OnOff> snapToGrid
        {
            get { return this.snapToGridField; }
            set { this.snapToGridField = value; }
        }

        List<CT_Language> langField;
        public List<CT_Language> lang
        {
            get { return this.langField; }
            set { this.langField = value; }
        }

        List<CT_HpsMeasure> kernField;
        public List<CT_HpsMeasure> kern
        {
            get { return this.kernField; }
            set { this.kernField = value; }
        }

        List<CT_OnOff> outlineField;
        public List<CT_OnOff> outline
        {
            get { return this.outlineField; }
            set { this.outlineField = value; }
        }

        List<CT_SignedHpsMeasure> positionField;
        public List<CT_SignedHpsMeasure> position
        {
            get { return this.positionField; }
            set { this.positionField = value; }
        }

        List<CT_Fonts> rFontsField;
        public List<CT_Fonts> rFonts
        {
            get { return this.rFontsField; }
            set { this.rFontsField = value; }
        }

        List<CT_String> rStyleField;
        public List<CT_String> rStyle
        {
            get { return this.rStyleField; }
            set { this.rStyleField = value; }
        }

        List<CT_OnOff> rtlField;
        public List<CT_OnOff> rtl
        {
            get { return this.rtlField; }
            set { this.rtlField = value; }
        }

        List<CT_OnOff> shadowField;
        public List<CT_OnOff> shadow
        {
            get { return this.shadowField; }
            set { this.shadowField = value; }
        }

        List<CT_OnOff> strikeField;
        public List<CT_OnOff> strike
        {
            get { return this.strikeField; }
            set { this.strikeField = value; }
        }

        List<CT_Shd> shdField;
        public List<CT_Shd> shd
        {
            get { return this.shdField; }
            set { this.shdField = value; }
        }

        List<CT_HpsMeasure> szField;
        public List<CT_HpsMeasure> sz
        {
            get { return this.szField; }
            set { this.szField = value; }
        }

        List<CT_HpsMeasure> szCsField;
        public List<CT_HpsMeasure> szCs
        {
            get { return this.szCsField; }
            set { this.szCsField = value; }
        }

        List<CT_OnOff> smallCapsField;
        public List<CT_OnOff> smallCaps
        {
            get { return this.smallCapsField; }
            set { this.smallCapsField = value; }
        }

        List<CT_Underline> uField;
        public List<CT_Underline> u
        {
            get { return this.uField; }
            set { this.uField = value; }
        }

        List<CT_OnOff> vanishField;
        public List<CT_OnOff> vanish
        {
            get { return this.vanishField; }
            set { this.vanishField = value; }
        }

        List<CT_OnOff> oMathField;
        public List<CT_OnOff> oMath
        {
            get { return this.oMathField; }
            set { this.oMathField = value; }
        }

        List<CT_OnOff> webHiddenField;
        public List<CT_OnOff> webHidden
        {
            get { return this.webHiddenField; }
            set { this.webHiddenField = value; }
        }

        List<CT_OnOff> specVanishField;
        public List<CT_OnOff> specVanish
        {
            get { return this.specVanishField; }
            set { this.specVanishField = value; }
        }

        List<CT_OnOff> bField;
        public List<CT_OnOff> b
        {
            get { return this.bField; }
            set { this.bField = value; }
        }

        List<CT_OnOff> bCsField;
        public List<CT_OnOff> bCs
        {
            get { return this.bCsField; }
            set { this.bCsField = value; }
        }

        List<CT_Border> bdrField;
        public List<CT_Border> bdr
        {
            get { return this.bdrField; }
            set { this.bdrField = value; }
        }

        List<CT_OnOff> capsField;
        public List<CT_OnOff> caps
        {
            get { return this.capsField; }
            set { this.capsField = value; }
        }

        List<CT_Color> colorField;
        public List<CT_Color> color
        {
            get { return this.colorField; }
            set { this.colorField = value; }
        }

        List<CT_OnOff> csField;
        public List<CT_OnOff> cs
        {
            get { return this.csField; }
            set { this.csField = value; }
        }

        List<CT_OnOff> dstrikeField;
        public List<CT_OnOff> dstrike
        {
            get { return this.dstrikeField; }
            set { this.dstrikeField = value; }
        }

        List<CT_EastAsianLayout> eastAsianLayoutField;
        public List<CT_EastAsianLayout> eastAsianLayout
        {
            get { return this.eastAsianLayoutField; }
            set { this.eastAsianLayoutField = value; }
        }

        List<CT_TextEffect> effectField;
        public List<CT_TextEffect> effect
        {
            get { return this.effectField; }
            set { this.effectField = value; }
        }

        List<CT_Em> emField;
        public List<CT_Em> em
        {
            get { return this.emField; }
            set { this.emField = value; }
        }

        List<CT_OnOff> embossField;
        public List<CT_OnOff> emboss
        {
            get { return this.embossField; }
            set { this.embossField = value; }
        }

        List<CT_FitText> fitTextField;
        public List<CT_FitText> fitText
        {
            get { return this.fitTextField; }
            set { this.fitTextField = value; }
        }

        List<CT_Highlight> highlightField;
        public List<CT_Highlight> highlight
        {
            get { return this.highlightField; }
            set { this.highlightField = value; }
        }

        List<CT_OnOff> iField;
        public List<CT_OnOff> i
        {
            get { return this.iField; }
            set { this.iField = value; }
        }

        List<CT_OnOff> iCsField;
        public List<CT_OnOff> iCs
        {
            get { return this.iCsField; }
            set { this.iCsField = value; }
        }

        List<CT_OnOff> imprintField;
        public List<CT_OnOff> imprint
        {
            get { return this.imprintField; }
            set { this.imprintField = value; }
        }


        public static CT_RPrOriginal Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_RPrOriginal ctObj = new CT_RPrOriginal();
            ctObj.spacing = new List<CT_SignedTwipsMeasure>();
            ctObj.vertAlign = new List<CT_VerticalAlignRun>();
            ctObj.w = new List<CT_TextScale>();
            ctObj.noProof = new List<CT_OnOff>();
            ctObj.snapToGrid = new List<CT_OnOff>();
            ctObj.lang = new List<CT_Language>();
            ctObj.kern = new List<CT_HpsMeasure>();
            ctObj.outline = new List<CT_OnOff>();
            ctObj.position = new List<CT_SignedHpsMeasure>();
            ctObj.rFonts = new List<CT_Fonts>();
            ctObj.rStyle = new List<CT_String>();
            ctObj.rtl = new List<CT_OnOff>();
            ctObj.shadow = new List<CT_OnOff>();
            ctObj.strike = new List<CT_OnOff>();
            ctObj.shd = new List<CT_Shd>();
            ctObj.sz = new List<CT_HpsMeasure>();
            ctObj.szCs = new List<CT_HpsMeasure>();
            ctObj.smallCaps = new List<CT_OnOff>();
            ctObj.u = new List<CT_Underline>();
            ctObj.vanish = new List<CT_OnOff>();
            ctObj.oMath = new List<CT_OnOff>();
            ctObj.webHidden = new List<CT_OnOff>();
            ctObj.specVanish = new List<CT_OnOff>();
            ctObj.b = new List<CT_OnOff>();
            ctObj.bCs = new List<CT_OnOff>();
            ctObj.bdr = new List<CT_Border>();
            ctObj.caps = new List<CT_OnOff>();
            ctObj.color = new List<CT_Color>();
            ctObj.cs = new List<CT_OnOff>();
            ctObj.dstrike = new List<CT_OnOff>();
            ctObj.eastAsianLayout = new List<CT_EastAsianLayout>();
            ctObj.effect = new List<CT_TextEffect>();
            ctObj.em = new List<CT_Em>();
            ctObj.emboss = new List<CT_OnOff>();
            ctObj.fitText = new List<CT_FitText>();
            ctObj.highlight = new List<CT_Highlight>();
            ctObj.i = new List<CT_OnOff>();
            ctObj.iCs = new List<CT_OnOff>();
            ctObj.imprint = new List<CT_OnOff>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "spacing")
                    ctObj.spacing.Add(CT_SignedTwipsMeasure.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "vertAlign")
                    ctObj.vertAlign.Add(CT_VerticalAlignRun.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "w")
                    ctObj.w.Add(CT_TextScale.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "noProof")
                    ctObj.noProof.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "snapToGrid")
                    ctObj.snapToGrid.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "lang")
                    ctObj.lang.Add(CT_Language.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "kern")
                    ctObj.kern.Add(CT_HpsMeasure.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "outline")
                    ctObj.outline.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "position")
                    ctObj.position.Add(CT_SignedHpsMeasure.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "rFonts")
                    ctObj.rFonts.Add(CT_Fonts.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "rStyle")
                    ctObj.rStyle.Add(CT_String.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "rtl")
                    ctObj.rtl.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "shadow")
                    ctObj.shadow.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "strike")
                    ctObj.strike.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "shd")
                    ctObj.shd.Add(CT_Shd.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "sz")
                    ctObj.sz.Add(CT_HpsMeasure.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "szCs")
                    ctObj.szCs.Add(CT_HpsMeasure.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "smallCaps")
                    ctObj.smallCaps.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "u")
                    ctObj.u.Add(CT_Underline.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "vanish")
                    ctObj.vanish.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "oMath")
                    ctObj.oMath.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "webHidden")
                    ctObj.webHidden.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "specVanish")
                    ctObj.specVanish.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "b")
                    ctObj.b.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "bCs")
                    ctObj.bCs.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "bdr")
                    ctObj.bdr.Add(CT_Border.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "caps")
                    ctObj.caps.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "color")
                    ctObj.color.Add(CT_Color.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "cs")
                    ctObj.cs.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "dstrike")
                    ctObj.dstrike.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "eastAsianLayout")
                    ctObj.eastAsianLayout.Add(CT_EastAsianLayout.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "effect")
                    ctObj.effect.Add(CT_TextEffect.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "em")
                    ctObj.em.Add(CT_Em.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "emboss")
                    ctObj.emboss.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "fitText")
                    ctObj.fitText.Add(CT_FitText.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "highlight")
                    ctObj.highlight.Add(CT_Highlight.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "i")
                    ctObj.i.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "iCs")
                    ctObj.iCs.Add(CT_OnOff.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "imprint")
                    ctObj.imprint.Add(CT_OnOff.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.spacing != null)
            {
                foreach (CT_SignedTwipsMeasure x in this.spacing)
                {
                    x.Write(sw, "spacing");
                }
            }
            if (this.vertAlign != null)
            {
                foreach (CT_VerticalAlignRun x in this.vertAlign)
                {
                    x.Write(sw, "vertAlign");
                }
            }
            if (this.w != null)
            {
                foreach (CT_TextScale x in this.w)
                {
                    x.Write(sw, "w");
                }
            }
            if (this.noProof != null)
            {
                foreach (CT_OnOff x in this.noProof)
                {
                    x.Write(sw, "noProof");
                }
            }
            if (this.snapToGrid != null)
            {
                foreach (CT_OnOff x in this.snapToGrid)
                {
                    x.Write(sw, "snapToGrid");
                }
            }
            if (this.lang != null)
            {
                foreach (CT_Language x in this.lang)
                {
                    x.Write(sw, "lang");
                }
            }
            if (this.kern != null)
            {
                foreach (CT_HpsMeasure x in this.kern)
                {
                    x.Write(sw, "kern");
                }
            }
            if (this.outline != null)
            {
                foreach (CT_OnOff x in this.outline)
                {
                    x.Write(sw, "outline");
                }
            }
            if (this.position != null)
            {
                foreach (CT_SignedHpsMeasure x in this.position)
                {
                    x.Write(sw, "position");
                }
            }
            if (this.rFonts != null)
            {
                foreach (CT_Fonts x in this.rFonts)
                {
                    x.Write(sw, "rFonts");
                }
            }
            if (this.rStyle != null)
            {
                foreach (CT_String x in this.rStyle)
                {
                    x.Write(sw, "rStyle");
                }
            }
            if (this.rtl != null)
            {
                foreach (CT_OnOff x in this.rtl)
                {
                    x.Write(sw, "rtl");
                }
            }
            if (this.shadow != null)
            {
                foreach (CT_OnOff x in this.shadow)
                {
                    x.Write(sw, "shadow");
                }
            }
            if (this.strike != null)
            {
                foreach (CT_OnOff x in this.strike)
                {
                    x.Write(sw, "strike");
                }
            }
            if (this.shd != null)
            {
                foreach (CT_Shd x in this.shd)
                {
                    x.Write(sw, "shd");
                }
            }
            if (this.sz != null)
            {
                foreach (CT_HpsMeasure x in this.sz)
                {
                    x.Write(sw, "sz");
                }
            }
            if (this.szCs != null)
            {
                foreach (CT_HpsMeasure x in this.szCs)
                {
                    x.Write(sw, "szCs");
                }
            }
            if (this.smallCaps != null)
            {
                foreach (CT_OnOff x in this.smallCaps)
                {
                    x.Write(sw, "smallCaps");
                }
            }
            if (this.u != null)
            {
                foreach (CT_Underline x in this.u)
                {
                    x.Write(sw, "u");
                }
            }
            if (this.vanish != null)
            {
                foreach (CT_OnOff x in this.vanish)
                {
                    x.Write(sw, "vanish");
                }
            }
            if (this.oMath != null)
            {
                foreach (CT_OnOff x in this.oMath)
                {
                    x.Write(sw, "oMath");
                }
            }
            if (this.webHidden != null)
            {
                foreach (CT_OnOff x in this.webHidden)
                {
                    x.Write(sw, "webHidden");
                }
            }
            if (this.specVanish != null)
            {
                foreach (CT_OnOff x in this.specVanish)
                {
                    x.Write(sw, "specVanish");
                }
            }
            if (this.b != null)
            {
                foreach (CT_OnOff x in this.b)
                {
                    x.Write(sw, "b");
                }
            }
            if (this.bCs != null)
            {
                foreach (CT_OnOff x in this.bCs)
                {
                    x.Write(sw, "bCs");
                }
            }
            if (this.bdr != null)
            {
                foreach (CT_Border x in this.bdr)
                {
                    x.Write(sw, "bdr");
                }
            }
            if (this.caps != null)
            {
                foreach (CT_OnOff x in this.caps)
                {
                    x.Write(sw, "caps");
                }
            }
            if (this.color != null)
            {
                foreach (CT_Color x in this.color)
                {
                    x.Write(sw, "color");
                }
            }
            if (this.cs != null)
            {
                foreach (CT_OnOff x in this.cs)
                {
                    x.Write(sw, "cs");
                }
            }
            if (this.dstrike != null)
            {
                foreach (CT_OnOff x in this.dstrike)
                {
                    x.Write(sw, "dstrike");
                }
            }
            if (this.eastAsianLayout != null)
            {
                foreach (CT_EastAsianLayout x in this.eastAsianLayout)
                {
                    x.Write(sw, "eastAsianLayout");
                }
            }
            if (this.effect != null)
            {
                foreach (CT_TextEffect x in this.effect)
                {
                    x.Write(sw, "effect");
                }
            }
            if (this.em != null)
            {
                foreach (CT_Em x in this.em)
                {
                    x.Write(sw, "em");
                }
            }
            if (this.emboss != null)
            {
                foreach (CT_OnOff x in this.emboss)
                {
                    x.Write(sw, "emboss");
                }
            }
            if (this.fitText != null)
            {
                foreach (CT_FitText x in this.fitText)
                {
                    x.Write(sw, "fitText");
                }
            }
            if (this.highlight != null)
            {
                foreach (CT_Highlight x in this.highlight)
                {
                    x.Write(sw, "highlight");
                }
            }
            if (this.i != null)
            {
                foreach (CT_OnOff x in this.i)
                {
                    x.Write(sw, "i");
                }
            }
            if (this.iCs != null)
            {
                foreach (CT_OnOff x in this.iCs)
                {
                    x.Write(sw, "iCs");
                }
            }
            if (this.imprint != null)
            {
                foreach (CT_OnOff x in this.imprint)
                {
                    x.Write(sw, "imprint");
                }
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

    }

    


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Br
    {

        private ST_BrType typeField;

        private bool typeFieldSpecified;

        private ST_BrClear clearField;

        private bool clearFieldSpecified;
        public static CT_Br Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Br ctObj = new CT_Br();
            if (node.Attributes["w:type"] != null)
                ctObj.type = (ST_BrType)Enum.Parse(typeof(ST_BrType), node.Attributes["w:type"].Value);
            if (node.Attributes["w:clear"] != null)
                ctObj.clear = (ST_BrClear)Enum.Parse(typeof(ST_BrClear), node.Attributes["w:clear"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:type", this.type.ToString());
            XmlHelper.WriteAttribute(sw, "w:clear", this.clear.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_BrType type
        {
            get
            {
                return this.typeField;
            }
            set
            {
                this.typeField = value;
                this.typeFieldSpecified = true;
            }
        }

        [XmlIgnore]
        public bool typeSpecified
        {
            get
            {
                return this.typeFieldSpecified;
            }
            set
            {
                this.typeFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_BrClear clear
        {
            get
            {
                return this.clearField;
            }
            set
            {
                this.clearField = value;
                this.clearFieldSpecified = true;
            }
        }

        [XmlIgnore]
        public bool clearSpecified
        {
            get
            {
                return this.clearFieldSpecified;
            }
            set
            {
                this.clearFieldSpecified = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_BrType
    {

    
        page,

    
        column,

    
        textWrapping,
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_BrClear
    {

    
        none,

    
        left,

    
        right,

    
        all,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TblStylePr
    {

        private CT_PPr pPrField;

        private CT_RPr rPrField;

        private CT_TblPrBase tblPrField;

        private CT_TrPr trPrField;

        private CT_TcPr tcPrField;

        private ST_TblStyleOverrideType typeField;
        public static CT_TblStylePr Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TblStylePr ctObj = new CT_TblStylePr();
            if (node.Attributes["w:type"] != null)
                ctObj.type = (ST_TblStyleOverrideType)Enum.Parse(typeof(ST_TblStyleOverrideType), node.Attributes["w:type"].Value);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "pPr")
                    ctObj.pPr = CT_PPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "rPr")
                    ctObj.rPr = CT_RPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblPr")
                    ctObj.tblPr = CT_TblPrBase.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "trPr")
                    ctObj.trPr = CT_TrPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tcPr")
                    ctObj.tcPr = CT_TcPr.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:type", this.type.ToString());
            sw.Write(">");
            if (this.pPr != null)
                this.pPr.Write(sw, "pPr");
            if (this.rPr != null)
                this.rPr.Write(sw, "rPr");
            if (this.tblPr != null)
                this.tblPr.Write(sw, "tblPr");
            if (this.trPr != null)
                this.trPr.Write(sw, "trPr");
            if (this.tcPr != null)
                this.tcPr.Write(sw, "tcPr");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        public CT_TblStylePr()
        {
            //this.tcPrField = new CT_TcPr();
            //this.trPrField = new CT_TrPr();
            //this.tblPrField = new CT_TblPrBase();
            //this.rPrField = new CT_RPr();
            //this.pPrField = new CT_PPr();
        }

        [XmlElement(Order = 0)]
        public CT_PPr pPr
        {
            get
            {
                return this.pPrField;
            }
            set
            {
                this.pPrField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_RPr rPr
        {
            get
            {
                return this.rPrField;
            }
            set
            {
                this.rPrField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_TblPrBase tblPr
        {
            get
            {
                return this.tblPrField;
            }
            set
            {
                this.tblPrField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_TrPr trPr
        {
            get
            {
                return this.trPrField;
            }
            set
            {
                this.trPrField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_TcPr tcPr
        {
            get
            {
                return this.tcPrField;
            }
            set
            {
                this.tcPrField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_TblStyleOverrideType type
        {
            get
            {
                return this.typeField;
            }
            set
            {
                this.typeField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_TblStyleOverrideType
    {

    
        wholeTable,

    
        firstRow,

    
        lastRow,

    
        firstCol,

    
        lastCol,

    
        band1Vert,

    
        band2Vert,

    
        band1Horz,

    
        band2Horz,

    
        neCell,

    
        nwCell,

    
        seCell,

    
        swCell,
    }
}