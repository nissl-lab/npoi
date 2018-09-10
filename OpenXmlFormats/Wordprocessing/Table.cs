using System;
using System.Collections.Generic;
using System.Xml.Serialization;
using NPOI.OpenXmlFormats.Shared;
using System.IO;
using System.Xml;
using System.Collections;
using NPOI.OpenXml4Net.Util;


namespace NPOI.OpenXmlFormats.Wordprocessing
{


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Tbl
    {
        //EG_RangeMarkupElements
        private ArrayList itemsField;

        private List<ItemsChoiceType30> itemsElementNameField;

        private CT_TblPr tblPrField;

        private CT_TblGrid tblGridField;

        private ArrayList items1Field;

        private List<Items1ChoiceType> items1ElementNameField;

        public CT_Tbl()
        {
            this.items1ElementNameField = new List<Items1ChoiceType>();
            this.items1Field = new ArrayList();
            //this.tblGridField = new CT_TblGrid();
            //this.tblPrField = new CT_TblPr();
            this.itemsElementNameField = new List<ItemsChoiceType30>();
            this.itemsField = new ArrayList();
        }
        public static CT_Tbl Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Tbl ctObj = new CT_Tbl();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "customXmlMoveToRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType30.customXmlMoveToRangeEnd);
                }
                else if (childNode.LocalName == "tblPr")
                    ctObj.tblPr = CT_TblPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblGrid")
                    ctObj.tblGrid = CT_TblGrid.Parse(childNode, namespaceManager);

                else if (childNode.LocalName == "moveToRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType30.moveToRangeEnd);
                }
                else if (childNode.LocalName == "moveToRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType30.moveToRangeStart);
                }
                else if (childNode.LocalName == "customXmlDelRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType30.customXmlDelRangeEnd);
                }
                else if (childNode.LocalName == "moveFromRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType30.moveFromRangeStart);
                }
                else if (childNode.LocalName == "bookmarkEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType30.bookmarkEnd);
                }
                else if (childNode.LocalName == "bookmarkStart")
                {
                    ctObj.Items.Add(CT_Bookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType30.bookmarkStart);
                }
                else if (childNode.LocalName == "commentRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType30.commentRangeEnd);
                }
                else if (childNode.LocalName == "commentRangeStart")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType30.commentRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType30.customXmlMoveToRangeStart);
                }
                else if (childNode.LocalName == "customXmlInsRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType30.customXmlInsRangeEnd);
                }
                else if (childNode.LocalName == "customXmlInsRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType30.customXmlInsRangeStart);
                }
                else if (childNode.LocalName == "customXmlDelRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType30.customXmlDelRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType30.customXmlMoveFromRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType30.customXmlMoveFromRangeStart);
                }
                else if (childNode.LocalName == "moveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType30.moveFromRangeEnd);
                }
                else if (childNode.LocalName == "sdt")
                {
                    ctObj.Items1.Add(CT_SdtRow.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.sdt);
                }
                else if (childNode.LocalName == "customXml")
                {
                    ctObj.Items1.Add(CT_CustomXmlRow.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.customXml);
                }
                else if (childNode.LocalName == "customXmlInsRangeEnd")
                {
                    ctObj.Items1.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.customXmlInsRangeEnd);
                }
                else if (childNode.LocalName == "customXmlInsRangeStart")
                {
                    ctObj.Items1.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.customXmlInsRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeEnd")
                {
                    ctObj.Items1.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.customXmlMoveFromRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeStart")
                {
                    ctObj.Items1.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.customXmlMoveFromRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeEnd")
                {
                    ctObj.Items1.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.customXmlMoveToRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeStart")
                {
                    ctObj.Items1.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.customXmlMoveToRangeStart);
                }
                else if (childNode.LocalName == "del")
                {
                    ctObj.Items1.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.del);
                }
                else if (childNode.LocalName == "ins")
                {
                    ctObj.Items1.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.ins);
                }
                else if (childNode.LocalName == "moveFrom")
                {
                    ctObj.Items1.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.moveFrom);
                }
                else if (childNode.LocalName == "moveFromRangeEnd")
                {
                    ctObj.Items1.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.moveFromRangeEnd);
                }
                else if (childNode.LocalName == "commentRangeStart")
                {
                    ctObj.Items1.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.commentRangeStart);
                }
                else if (childNode.LocalName == "moveTo")
                {
                    ctObj.Items1.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.moveTo);
                }
                else if (childNode.LocalName == "moveToRangeEnd")
                {
                    ctObj.Items1.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.moveToRangeEnd);
                }
                else if (childNode.LocalName == "moveToRangeStart")
                {
                    ctObj.Items1.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.moveToRangeStart);
                }
                else if (childNode.LocalName == "permEnd")
                {
                    ctObj.Items1.Add(CT_Perm.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.permEnd);
                }
                else if (childNode.LocalName == "permStart")
                {
                    ctObj.Items1.Add(CT_PermStart.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.permStart);
                }
                else if (childNode.LocalName == "proofErr")
                {
                    ctObj.Items1.Add(CT_ProofErr.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.proofErr);
                }
                else if (childNode.LocalName == "oMath")
                {
                    ctObj.Items1.Add(CT_OMath.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.oMath);
                }
                else if (childNode.LocalName == "tr")
                {
                    ctObj.Items1.Add(CT_Row.Parse(childNode, namespaceManager, ctObj));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.tr);
                }
                else if (childNode.LocalName == "moveFromRangeStart")
                {
                    ctObj.Items1.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.moveFromRangeStart);
                }
                else if (childNode.LocalName == "customXmlDelRangeEnd")
                {
                    ctObj.Items1.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.customXmlDelRangeEnd);
                }
                else if (childNode.LocalName == "customXmlDelRangeStart")
                {
                    ctObj.Items1.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.customXmlDelRangeStart);
                }
                else if (childNode.LocalName == "oMathPara")
                {
                    ctObj.Items1.Add(CT_OMathPara.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.oMathPara);
                }
                else if (childNode.LocalName == "bookmarkEnd")
                {
                    ctObj.Items1.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.bookmarkEnd);
                }
                else if (childNode.LocalName == "bookmarkStart")
                {
                    ctObj.Items1.Add(CT_Bookmark.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.bookmarkStart);
                }
                else if (childNode.LocalName == "commentRangeEnd")
                {
                    ctObj.Items1.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.Items1ElementName.Add(Items1ChoiceType.commentRangeEnd);
                }

            }
            return ctObj;
        }
        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            foreach (object o in this.Items)
            {
                if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlMoveToRangeEnd");

                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "moveToRangeEnd");
                else if (o is CT_MoveBookmark)
                    ((CT_MoveBookmark)o).Write(sw, "moveToRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlDelRangeEnd");
                else if (o is CT_MoveBookmark)
                    ((CT_MoveBookmark)o).Write(sw, "moveFromRangeStart");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "bookmarkEnd");
                else if (o is CT_Bookmark)
                    ((CT_Bookmark)o).Write(sw, "bookmarkStart");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeEnd");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeStart");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveToRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlInsRangeEnd");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlInsRangeStart");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlDelRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlMoveFromRangeEnd");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveFromRangeStart");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "moveFromRangeEnd");
            }
            if (this.tblPr != null)
                this.tblPr.Write(sw, "tblPr");
            if (this.tblGrid != null)
                this.tblGrid.Write(sw, "tblGrid");
            foreach (object o in this.Items1)
            {
                if (o is CT_SdtRow)
                    ((CT_SdtRow)o).Write(sw, "sdt");
                else if (o is CT_CustomXmlRow)
                    ((CT_CustomXmlRow)o).Write(sw, "customXml");
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
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "ins");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "moveFrom");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "moveFromRangeEnd");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeStart");
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
                else if (o is CT_OMath)
                    ((CT_OMath)o).Write(sw, "oMath");
                else if (o is CT_Row)
                    ((CT_Row)o).Write(sw, "tr");
                else if (o is CT_MoveBookmark)
                    ((CT_MoveBookmark)o).Write(sw, "moveFromRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlDelRangeEnd");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlDelRangeStart");
                else if (o is CT_OMathPara)
                    ((CT_OMathPara)o).Write(sw, "oMathPara");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "bookmarkEnd");
                else if (o is CT_Bookmark)
                    ((CT_Bookmark)o).Write(sw, "bookmarkStart");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeEnd");
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement("bookmarkEnd", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("bookmarkStart", typeof(CT_Bookmark), Order = 0)]
        [XmlElement("commentRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("commentRangeStart", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("customXmlDelRangeEnd", typeof(CT_Markup), Order = 0)]
        [XmlElement("customXmlDelRangeStart", typeof(CT_TrackChange), Order = 0)]
        [XmlElement("customXmlInsRangeEnd", typeof(CT_Markup), Order = 0)]
        [XmlElement("customXmlInsRangeStart", typeof(CT_TrackChange), Order = 0)]
        [XmlElement("customXmlMoveFromRangeEnd", typeof(CT_Markup), Order = 0)]
        [XmlElement("customXmlMoveFromRangeStart", typeof(CT_TrackChange), Order = 0)]
        [XmlElement("customXmlMoveToRangeEnd", typeof(CT_Markup), Order = 0)]
        [XmlElement("customXmlMoveToRangeStart", typeof(CT_TrackChange), Order = 0)]
        [XmlElement("moveFromRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("moveFromRangeStart", typeof(CT_MoveBookmark), Order = 0)]
        [XmlElement("moveToRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("moveToRangeStart", typeof(CT_MoveBookmark), Order = 0)]
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

        [XmlElement("ItemsElementName", Order = 1)]
        [XmlIgnore]
        public List<ItemsChoiceType30> ItemsElementName
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

        [XmlElement(Order = 2)]
        public CT_TblPr tblPr
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
        public CT_TblGrid tblGrid
        {
            get
            {
                return this.tblGridField;
            }
            set
            {
                this.tblGridField = value;
            }
        }

        //EG_ContentRowContent

        [XmlElement("oMath", typeof(CT_OMath), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 4)]
        [XmlElement("oMathPara", typeof(CT_OMathPara), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 4)]
        [XmlElement("bookmarkEnd", typeof(CT_MarkupRange), Order = 4)]
        [XmlElement("bookmarkStart", typeof(CT_Bookmark), Order = 4)]
        [XmlElement("commentRangeEnd", typeof(CT_MarkupRange), Order = 4)]
        [XmlElement("commentRangeStart", typeof(CT_MarkupRange), Order = 4)]
        [XmlElement("customXml", typeof(CT_CustomXmlRow), Order = 4)]
        [XmlElement("customXmlDelRangeEnd", typeof(CT_Markup), Order = 4)]
        [XmlElement("customXmlDelRangeStart", typeof(CT_TrackChange), Order = 4)]
        [XmlElement("customXmlInsRangeEnd", typeof(CT_Markup), Order = 4)]
        [XmlElement("customXmlInsRangeStart", typeof(CT_TrackChange), Order = 4)]
        [XmlElement("customXmlMoveFromRangeEnd", typeof(CT_Markup), Order = 4)]
        [XmlElement("customXmlMoveFromRangeStart", typeof(CT_TrackChange), Order = 4)]
        [XmlElement("customXmlMoveToRangeEnd", typeof(CT_Markup), Order = 4)]
        [XmlElement("customXmlMoveToRangeStart", typeof(CT_TrackChange), Order = 4)]
        [XmlElement("del", typeof(CT_RunTrackChange), Order = 4)]
        [XmlElement("ins", typeof(CT_RunTrackChange), Order = 4)]
        [XmlElement("moveFrom", typeof(CT_RunTrackChange), Order = 4)]
        [XmlElement("moveFromRangeEnd", typeof(CT_MarkupRange), Order = 4)]
        [XmlElement("moveFromRangeStart", typeof(CT_MoveBookmark), Order = 4)]
        [XmlElement("moveTo", typeof(CT_RunTrackChange), Order = 4)]
        [XmlElement("moveToRangeEnd", typeof(CT_MarkupRange), Order = 4)]
        [XmlElement("moveToRangeStart", typeof(CT_MoveBookmark), Order = 4)]
        [XmlElement("permEnd", typeof(CT_Perm), Order = 4)]
        [XmlElement("permStart", typeof(CT_PermStart), Order = 4)]
        [XmlElement("proofErr", typeof(CT_ProofErr), Order = 4)]
        [XmlElement("sdt", typeof(CT_SdtRow), Order = 4)]
        [XmlElement("tr", typeof(CT_Row), Order = 4)]
        [XmlChoiceIdentifier("Items1ElementName")]
        public ArrayList Items1
        {
            get
            {
                return this.items1Field;
            }
            set
            {
                this.items1Field = value;
            }
        }

        [XmlElement("Items1ElementName", Order = 5)]
        [XmlIgnore]
        public List<Items1ChoiceType> Items1ElementName
        {
            get
            {
                return this.items1ElementNameField;
            }
            set
            {
                this.items1ElementNameField = value;
            }
        }

        public void Set(CT_Tbl table)
        {
            this.items1ElementNameField = new List<Items1ChoiceType>(table.Items1ElementName);
            this.items1Field = new ArrayList(table.items1Field);
            this.itemsElementNameField = new List<ItemsChoiceType30>(table.itemsElementNameField);
            this.itemsField = new ArrayList(table.itemsField);
            this.tblGridField = table.tblGridField;
            this.tblPrField = table.tblPrField;
        }

        public void RemoveTr(int pos)
        {
            RemoveItems1(Items1ChoiceType.tr, pos);
        }

        public CT_Row InsertNewTr(int pos)
        {
            return InsertNewItems1<CT_Row>(Items1ChoiceType.tr, pos);
        }

        public void SetTrArray(int pos, CT_Row cT_Row)
        {
            SetItems1Array<CT_Row>(Items1ChoiceType.tr, pos, cT_Row);
        }

        public CT_Row AddNewTr()
        {
            return AddNewItems1<CT_Row>(Items1ChoiceType.tr);
        }

        public CT_TblPr AddNewTblPr()
        {
            if (this.tblPrField == null)
                this.tblPrField = new CT_TblPr();
            return this.tblPrField;
        }

        public int SizeOfTrArray()
        {
            return SizeOfItems1Array(Items1ChoiceType.tr);
        }

        public CT_Row GetTrArray(int p)
        {
            return GetItems1Array<CT_Row>(p, Items1ChoiceType.tr);
        }
        
        public List<CT_Row> GetTrList()
        {
            return GetItems1List<CT_Row>(Items1ChoiceType.tr);
        }
        #region Generic methods for object operation

        private List<T> GetItems1List<T>(Items1ChoiceType type) where T : class
        {
            lock (this)
            {
                List<T> list = new List<T>();
                for (int i = 0; i < items1ElementNameField.Count; i++)
                {
                    if (items1ElementNameField[i] == type)
                        list.Add(items1Field[i] as T);
                }
                return list;
            }
        }
        private int SizeOfItems1Array(Items1ChoiceType type)
        {
            lock (this)
            {
                int size = 0;
                for (int i = 0; i < items1ElementNameField.Count; i++)
                {
                    if (items1ElementNameField[i] == type)
                        size++;
                }
                return size;
            }
        }
        private T GetItems1Array<T>(int p, Items1ChoiceType type) where T : class
        {
            lock (this)
            {
                int pos = GetItems1Index(type, p);
                if (pos < 0 || pos >= this.items1Field.Count)
                    return null;
                return items1Field[pos] as T;
            }
        }
        private T InsertNewItems1<T>(Items1ChoiceType type, int p) where T : class, new()
        {
            T t = new T();
            lock (this)
            {
                int pos = GetItems1Index(type, p);
                this.items1ElementNameField.Insert(pos, type);
                this.items1Field.Insert(pos, t);
            }
            return t;
        }
        private T AddNewItems1<T>(Items1ChoiceType type) where T : class, new()
        {
            T t = new T();
            lock (this)
            {
                this.items1ElementNameField.Add(type);
                this.items1Field.Add(t);
            }
            return t;
        }
        private void SetItems1Array<T>(Items1ChoiceType type, int p, T obj) where T : class
        {
            lock (this)
            {
                int pos = GetItems1Index(type, p);
                if (pos < 0 || pos >= this.items1Field.Count)
                    return;
                if (this.items1Field[pos] is T)
                    this.items1Field[pos] = obj;
                else
                    throw new Exception(string.Format(@"object types are difference, itemsField[{0}] is {1}, and parameter obj is {2}",
                        pos, this.items1Field[pos].GetType().Name, typeof(T).Name));
            }
        }
        private int GetItems1Index(Items1ChoiceType type, int p)
        {
            int index = -1;
            int pos = 0;
            for (int i = 0; i < items1ElementNameField.Count; i++)
            {
                if (items1ElementNameField[i] == type)
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
        private void RemoveItems1(Items1ChoiceType type, int p)
        {
            lock (this)
            {
                int pos = GetItems1Index(type, p);
                if (pos < 0 || pos >= this.items1Field.Count)
                    return;
                items1ElementNameField.RemoveAt(pos);
                items1Field.RemoveAt(pos);
            }
        }
        #endregion

        public CT_TblGrid AddNewTblGrid()
        {
            this.tblGrid = new CT_TblGrid();
            return this.tblGrid;
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType30
    {

    
        bookmarkEnd,

    
        bookmarkStart,

    
        commentRangeEnd,

    
        commentRangeStart,

    
        customXmlDelRangeEnd,

    
        customXmlDelRangeStart,

    
        customXmlInsRangeEnd,

    
        customXmlInsRangeStart,

    
        customXmlMoveFromRangeEnd,

    
        customXmlMoveFromRangeStart,

    
        customXmlMoveToRangeEnd,

    
        customXmlMoveToRangeStart,

    
        moveFromRangeEnd,

    
        moveFromRangeStart,

    
        moveToRangeEnd,

    
        moveToRangeStart,
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TblGridChange : CT_Markup
    {

        private List<CT_TblGridCol> tblGridField;

        public CT_TblGridChange()
        {
            //this.tblGridField = new List<CT_TblGridCol>();
        }
        public static new CT_TblGridChange Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TblGridChange ctObj = new CT_TblGridChange();
            ctObj.id = XmlHelper.ReadString(node.Attributes["r:id"]);
            ctObj.tblGrid = new List<CT_TblGridCol>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "tblGrid")
                    ctObj.tblGrid.Add(CT_TblGridCol.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal new void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "r:id", this.id);
            sw.Write(">");
            if (this.tblGrid != null)
            {
                foreach (CT_TblGridCol x in this.tblGrid)
                {
                    x.Write(sw, "tblGrid");
                }
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlArray(Order = 0)]
        [XmlArrayItem("gridCol", IsNullable = false)]
        public List<CT_TblGridCol> tblGrid
        {
            get
            {
                return this.tblGridField;
            }
            set
            {
                this.tblGridField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TblGridCol
    {

        private ulong wField;

        private bool wFieldSpecified;
        public static CT_TblGridCol Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TblGridCol ctObj = new CT_TblGridCol();
            ctObj.w = XmlHelper.ReadULong(node.Attributes["w:w"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:w", this.w);
            sw.Write("/>");
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong w
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
        public bool wSpecified
        {
            get { return this.wFieldSpecified; }
            set { this.wFieldSpecified = value; }
        }
    }
    [XmlInclude(typeof(CT_TblGrid))]
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TblGridBase
    {

        private List<CT_TblGridCol> gridColField;

        public CT_TblGridBase()
        {
            this.gridColField = new List<CT_TblGridCol>();
        }

        [XmlElement("gridCol", Order = 0)]
        public List<CT_TblGridCol> gridCol
        {
            get
            {
                return this.gridColField;
            }
            set
            {
                this.gridColField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TblGrid : CT_TblGridBase
    {

        private CT_TblGridChange tblGridChangeField;

        public CT_TblGrid()
        {
            //this.tblGridChangeField = new CT_TblGridChange();
        }
        public static CT_TblGrid Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TblGrid ctObj = new CT_TblGrid();
            ctObj.gridCol = new List<CT_TblGridCol>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "tblGridChange")
                    ctObj.tblGridChange = CT_TblGridChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "gridCol")
                    ctObj.gridCol.Add(CT_TblGridCol.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.tblGridChange != null)
                this.tblGridChange.Write(sw, "tblGridChange");
            if (this.gridCol != null)
            {
                foreach (CT_TblGridCol x in this.gridCol)
                {
                    x.Write(sw, "gridCol");
                }
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_TblGridChange tblGridChange
        {
            get
            {
                return this.tblGridChangeField;
            }
            set
            {
                this.tblGridChangeField = value;
            }
        }

        public CT_TblGridCol AddNewGridCol()
        {
            if (this.gridCol == null)
                gridCol = new List<CT_TblGridCol>();

            CT_TblGridCol col=new CT_TblGridCol();
            gridCol.Add(col);
            return col;
        }
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TblOverlap
    {
        public static CT_TblOverlap Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TblOverlap ctObj = new CT_TblOverlap();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_TblOverlap)Enum.Parse(typeof(ST_TblOverlap), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        private ST_TblOverlap valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_TblOverlap val
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_TblOverlap
    {

    
        never,

    
        overlap,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TblWidth
    {

        private string wField;

        private ST_TblWidth typeField = ST_TblWidth.auto;

        private bool typeFieldSpecified = true;
        public static CT_TblWidth Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TblWidth ctObj = new CT_TblWidth();
            ctObj.w = XmlHelper.ReadString(node.Attributes["w:w"]);
            if (node.Attributes["w:type"] != null)
                ctObj.type = (ST_TblWidth)Enum.Parse(typeof(ST_TblWidth), node.Attributes["w:type"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:w", this.w);
            XmlHelper.WriteAttribute(sw, "w:type", this.type.ToString());
            sw.Write("/>");
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string w
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

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_TblWidth type
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
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_TblWidth
    {

    
        nil,

    
        pct,

    
        dxa,

    
        auto,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TblPrChange : CT_TrackChange
    {
        public static new CT_TblPrChange Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TblPrChange ctObj = new CT_TblPrChange();
            ctObj.author = XmlHelper.ReadString(node.Attributes["w:author"]);
            ctObj.date = XmlHelper.ReadString(node.Attributes["w:date"]);
            ctObj.id = XmlHelper.ReadString(node.Attributes["r:id"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "tblPr")
                    ctObj.tblPr = CT_TblPrBase.Parse(childNode, namespaceManager);
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
            if (this.tblPr != null)
                this.tblPr.Write(sw, "tblPr");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        private CT_TblPrBase tblPrField;

        public CT_TblPrChange()
        {
            //this.tblPrField = new CT_TblPrBase();
        }

        [XmlElement(Order = 0)]
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
    }

    [XmlInclude(typeof(CT_TblPr))]

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TblPrBase
    {

        private CT_String tblStyleField;

        private CT_TblPPr tblpPrField;

        private CT_TblOverlap tblOverlapField;

        private CT_OnOff bidiVisualField;

        private CT_DecimalNumber tblStyleRowBandSizeField;

        private CT_DecimalNumber tblStyleColBandSizeField;

        private CT_TblWidth tblWField;

        private CT_Jc jcField;

        private CT_TblWidth tblCellSpacingField;

        private CT_TblWidth tblIndField;

        private CT_TblBorders tblBordersField;

        private CT_Shd shdField;

        private CT_TblLayoutType tblLayoutField;

        private CT_TblCellMar tblCellMarField;

        private CT_ShortHexNumber tblLookField;

        public CT_TblPrBase()
        {
        }
        public static CT_TblPrBase Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TblPrBase ctObj = new CT_TblPrBase();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "tblStyle")
                    ctObj.tblStyle = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblpPr")
                    ctObj.tblpPr = CT_TblPPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblOverlap")
                    ctObj.tblOverlap = CT_TblOverlap.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bidiVisual")
                    ctObj.bidiVisual = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblStyleRowBandSize")
                    ctObj.tblStyleRowBandSize = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblStyleColBandSize")
                    ctObj.tblStyleColBandSize = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblW")
                    ctObj.tblW = CT_TblWidth.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "jc")
                    ctObj.jc = CT_Jc.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblCellSpacing")
                    ctObj.tblCellSpacing = CT_TblWidth.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblInd")
                    ctObj.tblInd = CT_TblWidth.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblBorders")
                    ctObj.tblBorders = CT_TblBorders.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "shd")
                    ctObj.shd = CT_Shd.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblLayout")
                    ctObj.tblLayout = CT_TblLayoutType.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblCellMar")
                    ctObj.tblCellMar = CT_TblCellMar.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblLook")
                    ctObj.tblLook = CT_ShortHexNumber.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.tblStyle != null)
                this.tblStyle.Write(sw, "tblStyle");
            if (this.tblpPr != null)
                this.tblpPr.Write(sw, "tblpPr");
            if (this.tblOverlap != null)
                this.tblOverlap.Write(sw, "tblOverlap");
            if (this.bidiVisual != null)
                this.bidiVisual.Write(sw, "bidiVisual");
            if (this.tblStyleRowBandSize != null)
                this.tblStyleRowBandSize.Write(sw, "tblStyleRowBandSize");
            if (this.tblStyleColBandSize != null)
                this.tblStyleColBandSize.Write(sw, "tblStyleColBandSize");
            if (this.tblW != null)
                this.tblW.Write(sw, "tblW");
            if (this.jc != null)
                this.jc.Write(sw, "jc");
            if (this.tblCellSpacing != null)
                this.tblCellSpacing.Write(sw, "tblCellSpacing");
            if (this.tblInd != null)
                this.tblInd.Write(sw, "tblInd");
            if (this.tblBorders != null)
                this.tblBorders.Write(sw, "tblBorders");
            if (this.shd != null)
                this.shd.Write(sw, "shd");
            if (this.tblLayout != null)
                this.tblLayout.Write(sw, "tblLayout");
            if (this.tblCellMar != null)
                this.tblCellMar.Write(sw, "tblCellMar");
            if (this.tblLook != null)
                this.tblLook.Write(sw, "tblLook");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_String tblStyle
        {
            get
            {
                return this.tblStyleField;
            }
            set
            {
                this.tblStyleField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_TblPPr tblpPr
        {
            get
            {
                return this.tblpPrField;
            }
            set
            {
                this.tblpPrField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_TblOverlap tblOverlap
        {
            get
            {
                return this.tblOverlapField;
            }
            set
            {
                this.tblOverlapField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_OnOff bidiVisual
        {
            get
            {
                return this.bidiVisualField;
            }
            set
            {
                this.bidiVisualField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_DecimalNumber tblStyleRowBandSize
        {
            get
            {
                return this.tblStyleRowBandSizeField;
            }
            set
            {
                this.tblStyleRowBandSizeField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_DecimalNumber tblStyleColBandSize
        {
            get
            {
                return this.tblStyleColBandSizeField;
            }
            set
            {
                this.tblStyleColBandSizeField = value;
            }
        }

        [XmlElement(Order = 6)]
        public CT_TblWidth tblW
        {
            get
            {
                return this.tblWField;
            }
            set
            {
                this.tblWField = value;
            }
        }

        [XmlElement(Order = 7)]
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

        [XmlElement(Order = 8)]
        public CT_TblWidth tblCellSpacing
        {
            get
            {
                return this.tblCellSpacingField;
            }
            set
            {
                this.tblCellSpacingField = value;
            }
        }

        [XmlElement(Order = 9)]
        public CT_TblWidth tblInd
        {
            get
            {
                return this.tblIndField;
            }
            set
            {
                this.tblIndField = value;
            }
        }

        [XmlElement(Order = 10)]
        public CT_TblBorders tblBorders
        {
            get
            {
                return this.tblBordersField;
            }
            set
            {
                this.tblBordersField = value;
            }
        }

        [XmlElement(Order = 11)]
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

        [XmlElement(Order = 12)]
        public CT_TblLayoutType tblLayout
        {
            get
            {
                return this.tblLayoutField;
            }
            set
            {
                this.tblLayoutField = value;
            }
        }

        [XmlElement(Order = 13)]
        public CT_TblCellMar tblCellMar
        {
            get
            {
                return this.tblCellMarField;
            }
            set
            {
                this.tblCellMarField = value;
            }
        }

        [XmlElement(Order = 14)]
        public CT_ShortHexNumber tblLook
        {
            get
            {
                return this.tblLookField;
            }
            set
            {
                this.tblLookField = value;
            }
        }

        public bool IsSetTblW()
        {
            return this.tblW != null;
        }

        public CT_TblWidth AddNewTblW()
        {
            if (this.tblWField == null)
                this.tblWField = new CT_TblWidth();
            return this.tblWField;
        }

        public CT_TblBorders AddNewTblBorders()
        {
            if (tblBordersField == null)
                this.tblBordersField = new CT_TblBorders();
            return this.tblBordersField;
        }
        public CT_String AddNewTblStyle()
        {
            this.tblStyleField = new CT_String();
            return this.tblStyleField;
        }

        public bool IsSetTblBorders()
        {
            return this.tblBordersField != null;
        }

        public bool IsSetTblStyleRowBandSize()
        {
            return this.tblStyleRowBandSizeField != null;
        }

        public CT_DecimalNumber AddNewTblStyleRowBandSize()
        {
            this.tblStyleRowBandSizeField = new CT_DecimalNumber();
            return this.tblStyleRowBandSizeField;
        }

        public bool IsSetTblStyleColBandSize()
        {
            return this.tblStyleColBandSizeField != null;
        }

        public CT_DecimalNumber AddNewTblStyleColBandSize()
        {
            this.tblStyleColBandSizeField = new CT_DecimalNumber();
            return this.tblStyleColBandSizeField;
        }

        public bool IsSetTblCellMar()
        {
            return this.tblCellMarField != null;
        }

        public CT_TblCellMar AddNewTblCellMar()
        {
            this.tblCellMarField = new CT_TblCellMar();
            return this.tblCellMarField;
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TblPPr
    {

        private ulong leftFromTextField;

        private bool leftFromTextFieldSpecified;

        private ulong rightFromTextField;

        private bool rightFromTextFieldSpecified;

        private ulong topFromTextField;

        private bool topFromTextFieldSpecified;

        private ulong bottomFromTextField;

        private bool bottomFromTextFieldSpecified;

        private ST_VAnchor vertAnchorField;

        private bool vertAnchorFieldSpecified;

        private ST_HAnchor horzAnchorField;

        private bool horzAnchorFieldSpecified;

        private ST_XAlign tblpXSpecField;

        private bool tblpXSpecFieldSpecified;

        private string tblpXField;

        private ST_YAlign tblpYSpecField;

        private bool tblpYSpecFieldSpecified;

        private string tblpYField;

        public static CT_TblPPr Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TblPPr ctObj = new CT_TblPPr();
            ctObj.leftFromText = XmlHelper.ReadULong(node.Attributes["w:leftFromText"]);
            ctObj.rightFromText = XmlHelper.ReadULong(node.Attributes["w:rightFromText"]);
            ctObj.topFromText = XmlHelper.ReadULong(node.Attributes["w:topFromText"]);
            ctObj.bottomFromText = XmlHelper.ReadULong(node.Attributes["w:bottomFromText"]);
            if (node.Attributes["w:vertAnchor"] != null)
                ctObj.vertAnchor = (ST_VAnchor)Enum.Parse(typeof(ST_VAnchor), node.Attributes["w:vertAnchor"].Value);
            if (node.Attributes["w:horzAnchor"] != null)
                ctObj.horzAnchor = (ST_HAnchor)Enum.Parse(typeof(ST_HAnchor), node.Attributes["w:horzAnchor"].Value);
            if (node.Attributes["w:tblpXSpec"] != null)
                ctObj.tblpXSpec = (ST_XAlign)Enum.Parse(typeof(ST_XAlign), node.Attributes["w:tblpXSpec"].Value);
            ctObj.tblpX = XmlHelper.ReadString(node.Attributes["w:tblpX"]);
            if (node.Attributes["w:tblpYSpec"] != null)
                ctObj.tblpYSpec = (ST_YAlign)Enum.Parse(typeof(ST_YAlign), node.Attributes["w:tblpYSpec"].Value);
            ctObj.tblpY = XmlHelper.ReadString(node.Attributes["w:tblpY"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:leftFromText", this.leftFromText);
            XmlHelper.WriteAttribute(sw, "w:rightFromText", this.rightFromText);
            XmlHelper.WriteAttribute(sw, "w:topFromText", this.topFromText);
            XmlHelper.WriteAttribute(sw, "w:bottomFromText", this.bottomFromText);
            XmlHelper.WriteAttribute(sw, "w:vertAnchor", this.vertAnchor.ToString());
            XmlHelper.WriteAttribute(sw, "w:horzAnchor", this.horzAnchor.ToString());
            if (this.tblpXSpecFieldSpecified)
                XmlHelper.WriteAttribute(sw, "w:tblpXSpec", this.tblpXSpec.ToString());
            XmlHelper.WriteAttribute(sw, "w:tblpX", this.tblpX);
            if (this.tblpYSpecFieldSpecified)
                XmlHelper.WriteAttribute(sw, "w:tblpYSpec", this.tblpYSpec.ToString());
            XmlHelper.WriteAttribute(sw, "w:tblpY", this.tblpY);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong leftFromText
        {
            get
            {
                return this.leftFromTextField;
            }
            set
            {
                this.leftFromTextField = value;
            }
        }

        [XmlIgnore]
        public bool leftFromTextSpecified
        {
            get
            {
                return this.leftFromTextFieldSpecified;
            }
            set
            {
                this.leftFromTextFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong rightFromText
        {
            get
            {
                return this.rightFromTextField;
            }
            set
            {
                this.rightFromTextField = value;
            }
        }

        [XmlIgnore]
        public bool rightFromTextSpecified
        {
            get
            {
                return this.rightFromTextFieldSpecified;
            }
            set
            {
                this.rightFromTextFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong topFromText
        {
            get
            {
                return this.topFromTextField;
            }
            set
            {
                this.topFromTextField = value;
            }
        }

        [XmlIgnore]
        public bool topFromTextSpecified
        {
            get
            {
                return this.topFromTextFieldSpecified;
            }
            set
            {
                this.topFromTextFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong bottomFromText
        {
            get
            {
                return this.bottomFromTextField;
            }
            set
            {
                this.bottomFromTextField = value;
            }
        }

        [XmlIgnore]
        public bool bottomFromTextSpecified
        {
            get
            {
                return this.bottomFromTextFieldSpecified;
            }
            set
            {
                this.bottomFromTextFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_VAnchor vertAnchor
        {
            get
            {
                return this.vertAnchorField;
            }
            set
            {
                this.vertAnchorField = value;
            }
        }

        [XmlIgnore]
        public bool vertAnchorSpecified
        {
            get
            {
                return this.vertAnchorFieldSpecified;
            }
            set
            {
                this.vertAnchorFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_HAnchor horzAnchor
        {
            get
            {
                return this.horzAnchorField;
            }
            set
            {
                this.horzAnchorField = value;
            }
        }

        [XmlIgnore]
        public bool horzAnchorSpecified
        {
            get
            {
                return this.horzAnchorFieldSpecified;
            }
            set
            {
                this.horzAnchorFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_XAlign tblpXSpec
        {
            get
            {
                return this.tblpXSpecField;
            }
            set
            {
                this.tblpXSpecFieldSpecified = true;
                this.tblpXSpecField = value;
            }
        }

        [XmlIgnore]
        public bool tblpXSpecSpecified
        {
            get
            {
                return this.tblpXSpecFieldSpecified;
            }
            set
            {
                this.tblpXSpecFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string tblpX
        {
            get
            {
                return this.tblpXField;
            }
            set
            {
                this.tblpXField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_YAlign tblpYSpec
        {
            get
            {
                return this.tblpYSpecField;
            }
            set
            {
                this.tblpYSpecFieldSpecified = true;
                this.tblpYSpecField = value;
            }
        }

        [XmlIgnore]
        public bool tblpYSpecSpecified
        {
            get
            {
                return this.tblpYSpecFieldSpecified;
            }
            set
            {
                this.tblpYSpecFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string tblpY
        {
            get
            {
                return this.tblpYField;
            }
            set
            {
                this.tblpYField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Height
    {

        private ulong valField;


        private ST_HeightRule hRuleField= ST_HeightRule.auto;

        public static CT_Height Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Height ctObj = new CT_Height();
            ctObj.val = XmlHelper.ReadULong(node.Attributes["w:val"]);
            if (node.Attributes["w:hRule"] != null)
                ctObj.hRule = (ST_HeightRule)Enum.Parse(typeof(ST_HeightRule), node.Attributes["w:hRule"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val);
            if(this.hRule!= ST_HeightRule.auto)
                XmlHelper.WriteAttribute(sw, "w:hRule", this.hRule.ToString());
            sw.Write("/>");
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong val
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

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_HeightRule hRule
        {
            get
            {
                return this.hRuleField;
            }
            set
            {
                this.hRuleField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType2
    {

    
        cantSplit,

    
        cnfStyle,

    
        divId,

    
        gridAfter,

    
        gridBefore,

    
        hidden,

    
        jc,

    
        tblCellSpacing,

    
        tblHeader,

    
        trHeight,

    
        wAfter,

    
        wBefore,
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TblPrExChange : CT_TrackChange
    {

        private CT_TblPrExBase tblPrExField;

        public CT_TblPrExChange()
        {
            
        }
        public static new CT_TblPrExChange Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TblPrExChange ctObj = new CT_TblPrExChange();
            ctObj.author = XmlHelper.ReadString(node.Attributes["w:author"]);
            ctObj.date = XmlHelper.ReadString(node.Attributes["w:date"]);
            ctObj.id = XmlHelper.ReadString(node.Attributes["w:id"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "tblPrEx")
                    ctObj.tblPrEx = CT_TblPrExBase.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal new void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:author", this.author);
            XmlHelper.WriteAttribute(sw, "w:date", this.date);
            XmlHelper.WriteAttribute(sw, "w:id", this.id);
            sw.Write(">");
            if (this.tblPrEx != null)
                this.tblPrEx.Write(sw, "tblPrEx");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_TblPrExBase tblPrEx
        {
            get
            {
                return this.tblPrExField;
            }
            set
            {
                this.tblPrExField = value;
            }
        }
    }

    [XmlInclude(typeof(CT_TblPrEx))]

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TblPrExBase
    {

        private CT_TblWidth tblWField;

        private CT_Jc jcField;

        private CT_TblWidth tblCellSpacingField;

        private CT_TblWidth tblIndField;

        private CT_TblBorders tblBordersField;

        private CT_Shd shdField;

        private CT_TblLayoutType tblLayoutField;

        private CT_TblCellMar tblCellMarField;

        private CT_ShortHexNumber tblLookField;
        public static CT_TblPrExBase Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TblPrExBase ctObj = new CT_TblPrExBase();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "tblW")
                    ctObj.tblW = CT_TblWidth.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "jc")
                    ctObj.jc = CT_Jc.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblCellSpacing")
                    ctObj.tblCellSpacing = CT_TblWidth.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblInd")
                    ctObj.tblInd = CT_TblWidth.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblBorders")
                    ctObj.tblBorders = CT_TblBorders.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "shd")
                    ctObj.shd = CT_Shd.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblLayout")
                    ctObj.tblLayout = CT_TblLayoutType.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblCellMar")
                    ctObj.tblCellMar = CT_TblCellMar.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblLook")
                    ctObj.tblLook = CT_ShortHexNumber.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.tblW != null)
                this.tblW.Write(sw, "tblW");
            if (this.jc != null)
                this.jc.Write(sw, "jc");
            if (this.tblCellSpacing != null)
                this.tblCellSpacing.Write(sw, "tblCellSpacing");
            if (this.tblInd != null)
                this.tblInd.Write(sw, "tblInd");
            if (this.tblBorders != null)
                this.tblBorders.Write(sw, "tblBorders");
            if (this.shd != null)
                this.shd.Write(sw, "shd");
            if (this.tblLayout != null)
                this.tblLayout.Write(sw, "tblLayout");
            if (this.tblCellMar != null)
                this.tblCellMar.Write(sw, "tblCellMar");
            if (this.tblLook != null)
                this.tblLook.Write(sw, "tblLook");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        public CT_TblPrExBase()
        {
            //this.tblLookField = new CT_ShortHexNumber();
            //this.tblCellMarField = new CT_TblCellMar();
            //this.tblLayoutField = new CT_TblLayoutType();
            //this.shdField = new CT_Shd();
            //this.tblBordersField = new CT_TblBorders();
            //this.tblIndField = new CT_TblWidth();
            //this.tblCellSpacingField = new CT_TblWidth();
            //this.jcField = new CT_Jc();
            //this.tblWField = new CT_TblWidth();
        }

        [XmlElement(Order = 0)]
        public CT_TblWidth tblW
        {
            get
            {
                return this.tblWField;
            }
            set
            {
                this.tblWField = value;
            }
        }

        [XmlElement(Order = 1)]
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

        [XmlElement(Order = 2)]
        public CT_TblWidth tblCellSpacing
        {
            get
            {
                return this.tblCellSpacingField;
            }
            set
            {
                this.tblCellSpacingField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_TblWidth tblInd
        {
            get
            {
                return this.tblIndField;
            }
            set
            {
                this.tblIndField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_TblBorders tblBorders
        {
            get
            {
                return this.tblBordersField;
            }
            set
            {
                this.tblBordersField = value;
            }
        }

        [XmlElement(Order = 5)]
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

        [XmlElement(Order = 6)]
        public CT_TblLayoutType tblLayout
        {
            get
            {
                return this.tblLayoutField;
            }
            set
            {
                this.tblLayoutField = value;
            }
        }

        [XmlElement(Order = 7)]
        public CT_TblCellMar tblCellMar
        {
            get
            {
                return this.tblCellMarField;
            }
            set
            {
                this.tblCellMarField = value;
            }
        }

        [XmlElement(Order = 8)]
        public CT_ShortHexNumber tblLook
        {
            get
            {
                return this.tblLookField;
            }
            set
            {
                this.tblLookField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TblPrEx : CT_TblPrExBase
    {

        private CT_TblPrExChange tblPrExChangeField;
        public static new CT_TblPrEx Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TblPrEx ctObj = new CT_TblPrEx();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "tblPrExChange")
                    ctObj.tblPrExChange = CT_TblPrExChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblW")
                    ctObj.tblW = CT_TblWidth.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "jc")
                    ctObj.jc = CT_Jc.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblCellSpacing")
                    ctObj.tblCellSpacing = CT_TblWidth.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblInd")
                    ctObj.tblInd = CT_TblWidth.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblBorders")
                    ctObj.tblBorders = CT_TblBorders.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "shd")
                    ctObj.shd = CT_Shd.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblLayout")
                    ctObj.tblLayout = CT_TblLayoutType.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblCellMar")
                    ctObj.tblCellMar = CT_TblCellMar.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblLook")
                    ctObj.tblLook = CT_ShortHexNumber.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal new void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.tblPrExChange != null)
                this.tblPrExChange.Write(sw, "tblPrExChange");
            if (this.tblW != null)
                this.tblW.Write(sw, "tblW");
            if (this.jc != null)
                this.jc.Write(sw, "jc");
            if (this.tblCellSpacing != null)
                this.tblCellSpacing.Write(sw, "tblCellSpacing");
            if (this.tblInd != null)
                this.tblInd.Write(sw, "tblInd");
            if (this.tblBorders != null)
                this.tblBorders.Write(sw, "tblBorders");
            if (this.shd != null)
                this.shd.Write(sw, "shd");
            if (this.tblLayout != null)
                this.tblLayout.Write(sw, "tblLayout");
            if (this.tblCellMar != null)
                this.tblCellMar.Write(sw, "tblCellMar");
            if (this.tblLook != null)
                this.tblLook.Write(sw, "tblLook");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        public CT_TblPrEx()
        {
        }

        [XmlElement(Order = 0)]
        public CT_TblPrExChange tblPrExChange
        {
            get
            {
                return this.tblPrExChangeField;
            }
            set
            {
                this.tblPrExChangeField = value;
            }
        }
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TblBorders
    {

        private CT_Border topField;

        private CT_Border leftField;

        private CT_Border bottomField;

        private CT_Border rightField;

        private CT_Border insideHField;

        private CT_Border insideVField;

        public CT_TblBorders()
        {
            //this.insideVField = new CT_Border();
            //this.insideHField = new CT_Border();
            //this.rightField = new CT_Border();
            //this.bottomField = new CT_Border();
            //this.leftField = new CT_Border();
            //this.topField = new CT_Border();
        }
        public static CT_TblBorders Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TblBorders ctObj = new CT_TblBorders();
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
                else if (childNode.LocalName == "insideH")
                    ctObj.insideH = CT_Border.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "insideV")
                    ctObj.insideV = CT_Border.Parse(childNode, namespaceManager);
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
            if (this.insideH != null)
                this.insideH.Write(sw, "insideH");
            if (this.insideV != null)
                this.insideV.Write(sw, "insideV");
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
        public CT_Border insideH
        {
            get
            {
                return this.insideHField;
            }
            set
            {
                this.insideHField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_Border insideV
        {
            get
            {
                return this.insideVField;
            }
            set
            {
                this.insideVField = value;
            }
        }

        public CT_Border AddNewBottom()
        {
            if (this.bottomField == null)
                this.bottomField = new CT_Border();
            return this.bottomField;
        }

        public CT_Border AddNewLeft()
        {
            if (this.leftField == null)
                this.leftField = new CT_Border();
            return this.leftField;
        }

        public CT_Border AddNewRight()
        {
            if (this.rightField == null)
                this.rightField = new CT_Border();
            return this.rightField;
        }

        public CT_Border AddNewTop()
        {
            if (this.topField == null)
                this.topField = new CT_Border();
            return this.topField;
        }

        public CT_Border AddNewInsideH()
        {
            if (this.insideHField == null)
                this.insideHField = new CT_Border();
            return this.insideHField;
        }

        public CT_Border AddNewInsideV()
        {
            if (this.insideVField == null)
                this.insideVField = new CT_Border();
            return this.insideVField;
        }

        public bool IsSetInsideH()
        {
            return this.insideH != null;
        }

        public bool IsSetInsideV()
        {
            return this.insideV != null;
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TblLayoutType
    {

        private ST_TblLayoutType typeField;

        private bool typeFieldSpecified;
        public static CT_TblLayoutType Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TblLayoutType ctObj = new CT_TblLayoutType();
            if (node.Attributes["w:type"] != null)
                ctObj.type = (ST_TblLayoutType)Enum.Parse(typeof(ST_TblLayoutType), node.Attributes["w:type"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:type", this.type.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_TblLayoutType type
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
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_TblLayoutType
    {

    
        @fixed,

    
        autofit,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TblCellMar
    {

        private CT_TblWidth topField;

        private CT_TblWidth leftField;

        private CT_TblWidth bottomField;

        private CT_TblWidth rightField;
        public static CT_TblCellMar Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TblCellMar ctObj = new CT_TblCellMar();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "top")
                    ctObj.top = CT_TblWidth.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "left")
                    ctObj.left = CT_TblWidth.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bottom")
                    ctObj.bottom = CT_TblWidth.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "right")
                    ctObj.right = CT_TblWidth.Parse(childNode, namespaceManager);
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
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        public CT_TblCellMar()
        {
            //this.rightField = new CT_TblWidth();
            //this.bottomField = new CT_TblWidth();
            //this.leftField = new CT_TblWidth();
            //this.topField = new CT_TblWidth();
        }

        [XmlElement(Order = 0)]
        public CT_TblWidth top
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
        public CT_TblWidth left
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
        public CT_TblWidth bottom
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
        public CT_TblWidth right
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

        public bool IsSetLeft()
        {
            return this.leftField != null;
        }

        public bool IsSetTop()
        {
            return this.topField != null;
        }

        public bool IsSetBottom()
        {
            return this.bottomField != null;
        }

        public bool IsSetRight()
        {
            return this.rightField != null;
        }

        public CT_TblWidth AddNewLeft()
        {
            this.leftField = new CT_TblWidth();
            return this.leftField;
        }

        public CT_TblWidth AddNewTop()
        {
            this.topField = new CT_TblWidth();
            return this.topField;
        }

        public CT_TblWidth AddNewBottom()
        {
            this.bottomField = new CT_TblWidth();
            return this.bottomField;
        }

        public CT_TblWidth AddNewRight()
        {
            this.rightField = new CT_TblWidth();
            return this.rightField;
        }
    }




    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TblPr : CT_TblPrBase
    {

        private CT_TblPrChange tblPrChangeField;
        public static new CT_TblPr Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TblPr ctObj = new CT_TblPr();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "tblPrChange")
                    ctObj.tblPrChange = CT_TblPrChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblStyle")
                    ctObj.tblStyle = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblpPr")
                    ctObj.tblpPr = CT_TblPPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblOverlap")
                    ctObj.tblOverlap = CT_TblOverlap.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bidiVisual")
                    ctObj.bidiVisual = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblStyleRowBandSize")
                    ctObj.tblStyleRowBandSize = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblStyleColBandSize")
                    ctObj.tblStyleColBandSize = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblW")
                    ctObj.tblW = CT_TblWidth.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "jc")
                    ctObj.jc = CT_Jc.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblCellSpacing")
                    ctObj.tblCellSpacing = CT_TblWidth.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblInd")
                    ctObj.tblInd = CT_TblWidth.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblBorders")
                    ctObj.tblBorders = CT_TblBorders.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "shd")
                    ctObj.shd = CT_Shd.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblLayout")
                    ctObj.tblLayout = CT_TblLayoutType.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblCellMar")
                    ctObj.tblCellMar = CT_TblCellMar.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblLook")
                    ctObj.tblLook = CT_ShortHexNumber.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal new void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.tblPrChange != null)
                this.tblPrChange.Write(sw, "tblPrChange");
            if (this.tblStyle != null)
                this.tblStyle.Write(sw, "tblStyle");
            if (this.tblpPr != null)
                this.tblpPr.Write(sw, "tblpPr");
            if (this.tblOverlap != null)
                this.tblOverlap.Write(sw, "tblOverlap");
            if (this.bidiVisual != null)
                this.bidiVisual.Write(sw, "bidiVisual");
            if (this.tblStyleRowBandSize != null)
                this.tblStyleRowBandSize.Write(sw, "tblStyleRowBandSize");
            if (this.tblStyleColBandSize != null)
                this.tblStyleColBandSize.Write(sw, "tblStyleColBandSize");
            if (this.tblW != null)
                this.tblW.Write(sw, "tblW");
            if (this.jc != null)
                this.jc.Write(sw, "jc");
            if (this.tblCellSpacing != null)
                this.tblCellSpacing.Write(sw, "tblCellSpacing");
            if (this.tblInd != null)
                this.tblInd.Write(sw, "tblInd");
            if (this.tblBorders != null)
                this.tblBorders.Write(sw, "tblBorders");
            if (this.shd != null)
                this.shd.Write(sw, "shd");
            if (this.tblLayout != null)
                this.tblLayout.Write(sw, "tblLayout");
            if (this.tblCellMar != null)
                this.tblCellMar.Write(sw, "tblCellMar");
            if (this.tblLook != null)
                this.tblLook.Write(sw, "tblLook");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        public CT_TblPr()
        {
            //this.tblPrChangeField = new CT_TblPrChange();
        }

        [XmlElement(Order = 0)]
        public CT_TblPrChange tblPrChange
        {
            get
            {
                return this.tblPrChangeField;
            }
            set
            {
                this.tblPrChangeField = value;
            }
        }

        public CT_TblLayoutType AddNewTblLayout()
        {
            this.tblLayout = new CT_TblLayoutType();
            return this.tblLayout;
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TrPrChange : CT_TrackChange
    {

        private CT_TrPrBase trPrField;

        public CT_TrPrChange()
        {
            
        }
        public static new CT_TrPrChange Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TrPrChange ctObj = new CT_TrPrChange();
            ctObj.author = XmlHelper.ReadString(node.Attributes["w:author"]);
            ctObj.date = XmlHelper.ReadString(node.Attributes["w:date"]);
            ctObj.id = XmlHelper.ReadString(node.Attributes["r:id"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "trPr")
                    ctObj.trPr = CT_TrPrBase.Parse(childNode, namespaceManager);
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
            if (this.trPr != null)
                this.trPr.Write(sw, "trPr");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_TrPrBase trPr
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
    }

    [XmlInclude(typeof(CT_TrPr))]

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TrPrBase
    {

        private ArrayList itemsField;

        private List<ItemsChoiceType2> itemsElementNameField;

        public CT_TrPrBase()
        {
            this.itemsElementNameField = new List<ItemsChoiceType2>();
            this.itemsField = new ArrayList();
        }
        public static CT_TrPrBase Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TrPrBase ctObj = new CT_TrPrBase();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "gridBefore")
                {
                    ctObj.Items.Add(CT_DecimalNumber.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.gridBefore);
                }
                else if (childNode.LocalName == "cantSplit")
                {
                    ctObj.Items.Add(CT_OnOff.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.cantSplit);
                }
                else if (childNode.LocalName == "cnfStyle")
                {
                    ctObj.Items.Add(CT_Cnf.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.cnfStyle);
                }
                else if (childNode.LocalName == "divId")
                {
                    ctObj.Items.Add(CT_DecimalNumber.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.divId);
                }
                else if (childNode.LocalName == "gridAfter")
                {
                    ctObj.Items.Add(CT_DecimalNumber.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.gridAfter);
                }
                else if (childNode.LocalName == "trHeight")
                {
                    ctObj.Items.Add(CT_Height.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.trHeight);
                }
                else if (childNode.LocalName == "hidden")
                {
                    ctObj.Items.Add(CT_OnOff.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.hidden);
                }
                else if (childNode.LocalName == "tblCellSpacing")
                {
                    ctObj.Items.Add(CT_TblWidth.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.tblCellSpacing);
                }
                else if (childNode.LocalName == "tblHeader")
                {
                    ctObj.Items.Add(CT_OnOff.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.tblHeader);
                }
                else if (childNode.LocalName == "jc")
                {
                    ctObj.Items.Add(CT_Jc.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.jc);
                }
                else if (childNode.LocalName == "wAfter")
                {
                    ctObj.Items.Add(CT_TblWidth.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.wAfter);
                }
                else if (childNode.LocalName == "wBefore")
                {
                    ctObj.Items.Add(CT_TblWidth.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.wBefore);
                }
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            for (int i = 0; i < this.Items.Count; i++)
            {
                object o = this.Items[i];
                if (o is CT_DecimalNumber
                    && this.ItemsElementName[i] == ItemsChoiceType2.gridBefore)
                    ((CT_DecimalNumber)o).Write(sw, "gridBefore");
                else if (o is CT_OnOff
                    && this.ItemsElementName[i] == ItemsChoiceType2.cantSplit)
                    ((CT_OnOff)o).Write(sw, "cantSplit");
                else if (o is CT_Cnf
                    && this.ItemsElementName[i] == ItemsChoiceType2.cnfStyle)
                    ((CT_Cnf)o).Write(sw, "cnfStyle");
                else if (o is CT_DecimalNumber
                    && this.ItemsElementName[i] == ItemsChoiceType2.divId)
                    ((CT_DecimalNumber)o).Write(sw, "divId");
                else if (o is CT_DecimalNumber
                    && this.ItemsElementName[i] == ItemsChoiceType2.gridAfter)
                    ((CT_DecimalNumber)o).Write(sw, "gridAfter");
                else if (o is CT_Height
                    && this.ItemsElementName[i] == ItemsChoiceType2.trHeight)
                    ((CT_Height)o).Write(sw, "trHeight");
                else if (o is CT_OnOff
                    && this.ItemsElementName[i] == ItemsChoiceType2.hidden)
                    ((CT_OnOff)o).Write(sw, "hidden");
                else if (o is CT_TblWidth
                    && this.ItemsElementName[i] == ItemsChoiceType2.tblCellSpacing)
                    ((CT_TblWidth)o).Write(sw, "tblCellSpacing");
                else if (o is CT_OnOff
                    && this.ItemsElementName[i] == ItemsChoiceType2.tblHeader)
                    ((CT_OnOff)o).Write(sw, "tblHeader");
                else if (o is CT_Jc
                    && this.ItemsElementName[i] == ItemsChoiceType2.jc)
                    ((CT_Jc)o).Write(sw, "jc");
                else if (o is CT_TblWidth
                    && this.ItemsElementName[i] == ItemsChoiceType2.wAfter)
                    ((CT_TblWidth)o).Write(sw, "wAfter");
                else if (o is CT_TblWidth
                    && this.ItemsElementName[i] == ItemsChoiceType2.wBefore)
                    ((CT_TblWidth)o).Write(sw, "wBefore");
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement("cantSplit", typeof(CT_OnOff), Order = 0)]
        [XmlElement("cnfStyle", typeof(CT_Cnf), Order = 0)]
        [XmlElement("divId", typeof(CT_DecimalNumber), Order = 0)]
        [XmlElement("gridAfter", typeof(CT_DecimalNumber), Order = 0)]
        [XmlElement("gridBefore", typeof(CT_DecimalNumber), Order = 0)]
        [XmlElement("hidden", typeof(CT_OnOff), Order = 0)]
        [XmlElement("jc", typeof(CT_Jc), Order = 0)]
        [XmlElement("tblCellSpacing", typeof(CT_TblWidth), Order = 0)]
        [XmlElement("tblHeader", typeof(CT_OnOff), Order = 0)]
        [XmlElement("trHeight", typeof(CT_Height), Order = 0)]
        [XmlElement("wAfter", typeof(CT_TblWidth), Order = 0)]
        [XmlElement("wBefore", typeof(CT_TblWidth), Order = 0)]
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

        [XmlElement("ItemsElementName", Order = 1)]
        [XmlIgnore]
        public List<ItemsChoiceType2> ItemsElementName
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
        public int SizeOfTrHeightArray()
        {
            return SizeOfArray(ItemsChoiceType2.trHeight);
        }

        public CT_Height GetTrHeightArray(int p)
        {
            return GetObjectArray<CT_Height>(p, ItemsChoiceType2.trHeight);
        }

        public CT_Height AddNewTrHeight()
        {
            return AddNewObject<CT_Height>(ItemsChoiceType2.trHeight);
        }

        public CT_OnOff AddNewCantSplit()
        {
            return AddNewObject<CT_OnOff>(ItemsChoiceType2.cantSplit);
        }

        public List<CT_OnOff> GetCantSplitList()
        {
            return GetObjectList<CT_OnOff>(ItemsChoiceType2.cantSplit);
        }

        public CT_OnOff AddNewTblHeader()
        {
            return AddNewObject<CT_OnOff>(ItemsChoiceType2.tblHeader);
        }

        public List<CT_OnOff> GetTblHeaderList()
        {
            return GetObjectList<CT_OnOff>(ItemsChoiceType2.tblHeader);
        }

        public int SizeOfTblHeaderArray()
        {
            return SizeOfArray(ItemsChoiceType2.tblHeader);
        }

        public int SizeOfCantSplitArray()
        {
            return SizeOfArray(ItemsChoiceType2.cantSplit);
        }
        #region Generic methods for object operation

        private List<T> GetObjectList<T>(ItemsChoiceType2 type) where T : class
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
        private int SizeOfArray(ItemsChoiceType2 type)
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
        private T GetObjectArray<T>(int p, ItemsChoiceType2 type) where T : class
        {
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                if (pos < 0 || pos >= this.itemsField.Count)
                    return null;
                return itemsField[pos] as T;
            }
        }
        private T InsertNewObject<T>(ItemsChoiceType2 type, int p) where T : class, new()
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
        private T AddNewObject<T>(ItemsChoiceType2 type) where T : class, new()
        {
            T t = new T();
            lock (this)
            {
                this.itemsElementNameField.Add(type);
                this.itemsField.Add(t);
            }
            return t;
        }
        private void SetObject<T>(ItemsChoiceType2 type, int p, T obj) where T : class
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
        private int GetObjectIndex(ItemsChoiceType2 type, int p)
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
        private void RemoveObject(ItemsChoiceType2 type, int p)
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


    #region Table Cell
    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Tc
    {

        private CT_TcPr tcPrField;

        private ArrayList itemsField;

        private List<ItemsChoiceTableCellType> itemsElementNameField;

        public CT_Tc()
        {
            this.itemsElementNameField = new List<ItemsChoiceTableCellType>();
            this.itemsField = new ArrayList();
            //this.tcPrField = new CT_TcPr();
        }

        [XmlElement(Order = 0)]
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
        public static CT_Tc Parse(XmlNode node, XmlNamespaceManager namespaceManager, object parent)
        {
            if (node == null)
                return null;
                
            CT_Tc ctObj = new CT_Tc();
            if (parent != null)
            {
                ctObj.parent = parent;
            }
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "commentRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.commentRangeEnd);
                }
                else if (childNode.LocalName == "tcPr")
                {
                    ctObj.tcPr = CT_TcPr.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "oMath")
                {
                    ctObj.Items.Add(CT_OMath.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.oMath);
                }
                else if (childNode.LocalName == "oMathPara")
                {
                    ctObj.Items.Add(CT_OMathPara.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.oMathPara);
                }
                else if (childNode.LocalName == "altChunk")
                {
                    ctObj.Items.Add(CT_AltChunk.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.altChunk);
                }
                else if (childNode.LocalName == "bookmarkEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.bookmarkEnd);
                }
                else if (childNode.LocalName == "bookmarkStart")
                {
                    ctObj.Items.Add(CT_Bookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.bookmarkStart);
                }
                else if (childNode.LocalName == "moveFromRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.moveFromRangeStart);
                }
                else if (childNode.LocalName == "commentRangeStart")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.commentRangeStart);
                }
                else if (childNode.LocalName == "customXml")
                {
                    ctObj.Items.Add(CT_CustomXmlBlock.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.customXml);
                }
                else if (childNode.LocalName == "customXmlDelRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.customXmlDelRangeEnd);
                }
                else if (childNode.LocalName == "customXmlDelRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.customXmlDelRangeStart);
                }
                else if (childNode.LocalName == "customXmlInsRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.customXmlInsRangeEnd);
                }
                else if (childNode.LocalName == "customXmlInsRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.customXmlInsRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.customXmlMoveFromRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.customXmlMoveFromRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.customXmlMoveToRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.customXmlMoveToRangeStart);
                }
                else if (childNode.LocalName == "del")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.del);
                }
                else if (childNode.LocalName == "ins")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.ins);
                }
                else if (childNode.LocalName == "moveFrom")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.moveFrom);
                }
                else if (childNode.LocalName == "moveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.moveFromRangeEnd);
                }
                else if (childNode.LocalName == "sdt")
                {
                    ctObj.Items.Add(CT_SdtBlock.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.sdt);
                }
                else if (childNode.LocalName == "moveTo")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.moveTo);
                }
                else if (childNode.LocalName == "moveToRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.moveToRangeEnd);
                }
                else if (childNode.LocalName == "moveToRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.moveToRangeStart);
                }
                else if (childNode.LocalName == "p")
                {
                    ctObj.Items.Add(CT_P.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.p);
                }
                else if (childNode.LocalName == "permEnd")
                {
                    ctObj.Items.Add(CT_Perm.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.permEnd);
                }
                else if (childNode.LocalName == "proofErr")
                {
                    ctObj.Items.Add(CT_ProofErr.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.proofErr);
                }
                else if (childNode.LocalName == "permStart")
                {
                    ctObj.Items.Add(CT_PermStart.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.permStart);
                }
                else if (childNode.LocalName == "tbl")
                {
                    ctObj.Items.Add(CT_Tbl.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableCellType.tbl);
                }
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if(this.tcPr!=null)
                    this.tcPr.Write(sw, "tcPr");
            foreach (object o in this.Items)
            {
                if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeEnd");
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
                else if (o is CT_MoveBookmark)
                    ((CT_MoveBookmark)o).Write(sw, "moveFromRangeStart");
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
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "del");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "ins");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "moveFrom");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "moveFromRangeEnd");
                else if (o is CT_SdtBlock)
                    ((CT_SdtBlock)o).Write(sw, "sdt");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "moveTo");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "moveToRangeEnd");
                else if (o is CT_MoveBookmark)
                    ((CT_MoveBookmark)o).Write(sw, "moveToRangeStart");
                else if (o is CT_P)
                    ((CT_P)o).Write(sw, "p");
                else if (o is CT_Perm)
                    ((CT_Perm)o).Write(sw, "permEnd");
                else if (o is CT_ProofErr)
                    ((CT_ProofErr)o).Write(sw, "proofErr");
                else if (o is CT_PermStart)
                    ((CT_PermStart)o).Write(sw, "permStart");
                else if (o is CT_Tbl)
                    ((CT_Tbl)o).Write(sw, "tbl");
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement("oMath", typeof(CT_OMath), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 1)]
        [XmlElement("oMathPara", typeof(CT_OMathPara), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 1)]
        [XmlElement("altChunk", typeof(CT_AltChunk), Order = 1)]
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
        public List<ItemsChoiceTableCellType> ItemsElementName
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
        object parent;
        [XmlIgnore]
        public object Parent
        {
            get { return parent; }
        }
        #region Generic methods for object operation
        private List<T> GetObjectList<T>(ItemsChoiceTableCellType type) where T : class
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
        private int SizeOfArray(ItemsChoiceTableCellType type)
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
        private T GetObjectArray<T>(int p, ItemsChoiceTableCellType type) where T : class
        {
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                if (pos < 0 || pos >= this.itemsField.Count)
                    return null;
                return itemsField[pos] as T;
            }
        }
        private int GetObjectIndex(ItemsChoiceTableCellType type, int p)
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
        private void RemoveObject(ItemsChoiceTableCellType type, int p)
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
        private void SetObject<T>(ItemsChoiceTableCellType type, int p, T obj) where T : class
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

        private T AddNewObject<T>(ItemsChoiceTableCellType type) where T : class, new()
        {
            T t = new T();
            lock (this)
            {
                this.itemsElementNameField.Add(type);
                this.itemsField.Add(t);
            }
            return t;
        }
        #endregion
        public CT_P AddNewP()
        {
            return AddNewObject<CT_P>(ItemsChoiceTableCellType.p);
        }

        public IList<CT_P> GetPList()
        {
            return GetObjectList<CT_P>(ItemsChoiceTableCellType.p);
        }

        public int SizeOfPArray()
        {
            return SizeOfArray(ItemsChoiceTableCellType.p);
        }

        public void SetPArray(int p, CT_P cT_P)
        {
            SetObject<CT_P>(ItemsChoiceTableCellType.p, p, cT_P);
        }

        public void RemoveP(int pos)
        {
            RemoveObject(ItemsChoiceTableCellType.p, pos);
        }

        public CT_P GetPArray(int p)
        {
            return GetObjectArray<CT_P>(p, ItemsChoiceTableCellType.p);
        }

        public IList<CT_Tbl> GetTblList()
        {
            return GetObjectList<CT_Tbl>(ItemsChoiceTableCellType.tbl);
        }

        public CT_Tbl GetTblArray(int p)
        {
            return GetObjectArray<CT_Tbl>(p, ItemsChoiceTableCellType.tbl);
        }

        public CT_TcPr AddNewTcPr()
        {
            this.tcPrField = new CT_TcPr();
            return this.tcPrField;
        }

        public bool IsSetTcPr()
        {
            return (this.tcPrField != null);

        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceTableCellType
    {

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:oMath")]
        oMath,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:oMathPara")]
        oMathPara,

    
        altChunk,

    
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
    public class CT_TrPr : CT_TrPrBase
    {

        private CT_TrackChange insField;

        private CT_TrackChange delField;

        private CT_TrPrChange trPrChangeField;

        public CT_TrPr()
        {
            //this.trPrChangeField = new CT_TrPrChange();
            //this.delField = new CT_TrackChange();
            //this.insField = new CT_TrackChange();
        }
        public static new CT_TrPr Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TrPr ctObj = new CT_TrPr();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "ins")
                    ctObj.ins = CT_TrackChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "del")
                    ctObj.del = CT_TrackChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "trPrChange")
                    ctObj.trPrChange = CT_TrPrChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "gridBefore")
                {
                    ctObj.Items.Add(CT_DecimalNumber.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.gridBefore);
                }
                else if (childNode.LocalName == "cantSplit")
                {
                    ctObj.Items.Add(CT_OnOff.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.cantSplit);
                }
                else if (childNode.LocalName == "cnfStyle")
                {
                    ctObj.Items.Add(CT_Cnf.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.cnfStyle);
                }
                else if (childNode.LocalName == "divId")
                {
                    ctObj.Items.Add(CT_DecimalNumber.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.divId);
                }
                else if (childNode.LocalName == "gridAfter")
                {
                    ctObj.Items.Add(CT_DecimalNumber.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.gridAfter);
                }
                else if (childNode.LocalName == "trHeight")
                {
                    ctObj.Items.Add(CT_Height.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.trHeight);
                }
                else if (childNode.LocalName == "hidden")
                {
                    ctObj.Items.Add(CT_OnOff.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.hidden);
                }
                else if (childNode.LocalName == "tblCellSpacing")
                {
                    ctObj.Items.Add(CT_TblWidth.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.tblCellSpacing);
                }
                else if (childNode.LocalName == "tblHeader")
                {
                    ctObj.Items.Add(CT_OnOff.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.tblHeader);
                }
                else if (childNode.LocalName == "jc")
                {
                    ctObj.Items.Add(CT_Jc.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.jc);
                }
                else if (childNode.LocalName == "wAfter")
                {
                    ctObj.Items.Add(CT_TblWidth.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.wAfter);
                }
                else if (childNode.LocalName == "wBefore")
                {
                    ctObj.Items.Add(CT_TblWidth.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType2.wBefore);
                }
            }
            return ctObj;
        }



        internal new void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.ins != null)
                this.ins.Write(sw, "ins");
            if (this.del != null)
                this.del.Write(sw, "del");
            if (this.trPrChange != null)
                this.trPrChange.Write(sw, "trPrChange");
            for (int i=0;i<this.Items.Count;i++)
            {
                object o = this.Items[i];
                if (o is CT_DecimalNumber 
                    &&this.ItemsElementName[i]== ItemsChoiceType2.gridBefore)
                    ((CT_DecimalNumber)o).Write(sw, "gridBefore");
                else if (o is CT_OnOff
                    && this.ItemsElementName[i] == ItemsChoiceType2.cantSplit)
                    ((CT_OnOff)o).Write(sw, "cantSplit");
                else if (o is CT_Cnf
                    && this.ItemsElementName[i] == ItemsChoiceType2.cnfStyle)
                    ((CT_Cnf)o).Write(sw, "cnfStyle");
                else if (o is CT_DecimalNumber
                    && this.ItemsElementName[i] == ItemsChoiceType2.divId)
                    ((CT_DecimalNumber)o).Write(sw, "divId");
                else if (o is CT_DecimalNumber
                    && this.ItemsElementName[i] == ItemsChoiceType2.gridAfter)
                    ((CT_DecimalNumber)o).Write(sw, "gridAfter");
                else if (o is CT_Height
                    && this.ItemsElementName[i] == ItemsChoiceType2.trHeight)
                    ((CT_Height)o).Write(sw, "trHeight");
                else if (o is CT_OnOff
                    && this.ItemsElementName[i] == ItemsChoiceType2.hidden)
                    ((CT_OnOff)o).Write(sw, "hidden");
                else if (o is CT_TblWidth
                    && this.ItemsElementName[i] == ItemsChoiceType2.tblCellSpacing)
                    ((CT_TblWidth)o).Write(sw, "tblCellSpacing");
                else if (o is CT_OnOff
                    && this.ItemsElementName[i] == ItemsChoiceType2.tblHeader)
                    ((CT_OnOff)o).Write(sw, "tblHeader");
                else if (o is CT_Jc
                    && this.ItemsElementName[i] == ItemsChoiceType2.jc)
                    ((CT_Jc)o).Write(sw, "jc");
                else if (o is CT_TblWidth
                    && this.ItemsElementName[i] == ItemsChoiceType2.wAfter)
                    ((CT_TblWidth)o).Write(sw, "wAfter");
                else if (o is CT_TblWidth
                    && this.ItemsElementName[i] == ItemsChoiceType2.wBefore)
                    ((CT_TblWidth)o).Write(sw, "wBefore");
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
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
        public CT_TrPrChange trPrChange
        {
            get
            {
                return this.trPrChangeField;
            }
            set
            {
                this.trPrChangeField = value;
            }
        }

    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TcPrChange : CT_TrackChange
    {
        public static new CT_TcPrChange Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TcPrChange ctObj = new CT_TcPrChange();
            ctObj.author = XmlHelper.ReadString(node.Attributes["w:author"]);
            ctObj.date = XmlHelper.ReadString(node.Attributes["w:date"]);
            ctObj.id = XmlHelper.ReadString(node.Attributes["r:id"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "tcPr")
                    ctObj.tcPr = CT_TcPrInner.Parse(childNode, namespaceManager);
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
            if (this.tcPr != null)
                this.tcPr.Write(sw, "tcPr");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        private CT_TcPrInner tcPrField;

        public CT_TcPrChange()
        {
            
        }

        [XmlElement(Order = 0)]
        public CT_TcPrInner tcPr
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
    }

    [XmlInclude(typeof(CT_TcPr))]
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TcPrInner : CT_TcPrBase
    {

        private CT_TrackChange cellInsField;

        private CT_TrackChange cellDelField;

        private CT_CellMergeTrackChange cellMergeField;

        public CT_TcPrInner()
        {
            //this.cellMergeField = new CT_CellMergeTrackChange();
            //this.cellDelField = new CT_TrackChange();
            //this.cellInsField = new CT_TrackChange();
        }
        public static CT_TcPrInner Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TcPrInner ctObj = new CT_TcPrInner();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "cellIns")
                    ctObj.cellIns = CT_TrackChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cellDel")
                    ctObj.cellDel = CT_TrackChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cellMerge")
                    ctObj.cellMerge = CT_CellMergeTrackChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cnfStyle")
                    ctObj.cnfStyle = CT_Cnf.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tcW")
                    ctObj.tcW = CT_TblWidth.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "gridSpan")
                    ctObj.gridSpan = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "hMerge")
                    ctObj.hMerge = CT_HMerge.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "vMerge")
                    ctObj.vMerge = CT_VMerge.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tcBorders")
                    ctObj.tcBorders = CT_TcBorders.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "shd")
                    ctObj.shd = CT_Shd.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "noWrap")
                    ctObj.noWrap = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tcMar")
                    ctObj.tcMar = CT_TcMar.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "textDirection")
                    ctObj.textDirection = CT_TextDirection.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tcFitText")
                    ctObj.tcFitText = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "vAlign")
                    ctObj.vAlign = CT_VerticalJc.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "hideMark")
                    ctObj.hideMark = CT_OnOff.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.cellIns != null)
                this.cellIns.Write(sw, "cellIns");
            if (this.cellDel != null)
                this.cellDel.Write(sw, "cellDel");
            if (this.cellMerge != null)
                this.cellMerge.Write(sw, "cellMerge");
            if (this.cnfStyle != null)
                this.cnfStyle.Write(sw, "cnfStyle");
            if (this.tcW != null)
                this.tcW.Write(sw, "tcW");
            if (this.gridSpan != null)
                this.gridSpan.Write(sw, "gridSpan");
            if (this.hMerge != null)
                this.hMerge.Write(sw, "hMerge");
            if (this.vMerge != null)
                this.vMerge.Write(sw, "vMerge");
            if (this.tcBorders != null)
                this.tcBorders.Write(sw, "tcBorders");
            if (this.shd != null)
                this.shd.Write(sw, "shd");
            if (this.noWrap != null)
                this.noWrap.Write(sw, "noWrap");
            if (this.tcMar != null)
                this.tcMar.Write(sw, "tcMar");
            if (this.textDirection != null)
                this.textDirection.Write(sw, "textDirection");
            if (this.tcFitText != null)
                this.tcFitText.Write(sw, "tcFitText");
            if (this.vAlign != null)
                this.vAlign.Write(sw, "vAlign");
            if (this.hideMark != null)
                this.hideMark.Write(sw, "hideMark");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_TrackChange cellIns
        {
            get
            {
                return this.cellInsField;
            }
            set
            {
                this.cellInsField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_TrackChange cellDel
        {
            get
            {
                return this.cellDelField;
            }
            set
            {
                this.cellDelField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_CellMergeTrackChange cellMerge
        {
            get
            {
                return this.cellMergeField;
            }
            set
            {
                this.cellMergeField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_CellMergeTrackChange : CT_TrackChange
    {

        private ST_AnnotationVMerge vMergeField;

        private bool vMergeFieldSpecified;

        private ST_AnnotationVMerge vMergeOrigField;

        private bool vMergeOrigFieldSpecified;
        public static new CT_CellMergeTrackChange Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CellMergeTrackChange ctObj = new CT_CellMergeTrackChange();
            if (node.Attributes["w:vMerge"] != null)
                ctObj.vMerge = (ST_AnnotationVMerge)Enum.Parse(typeof(ST_AnnotationVMerge), node.Attributes["w:vMerge"].Value);
            if (node.Attributes["w:vMergeOrig"] != null)
                ctObj.vMergeOrig = (ST_AnnotationVMerge)Enum.Parse(typeof(ST_AnnotationVMerge), node.Attributes["w:vMergeOrig"].Value);
            ctObj.author = XmlHelper.ReadString(node.Attributes["w:author"]);
            ctObj.date = XmlHelper.ReadString(node.Attributes["w:date"]);
            ctObj.id = XmlHelper.ReadString(node.Attributes["r:id"]);
            return ctObj;
        }



        internal new void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:vMerge", this.vMerge.ToString());
            XmlHelper.WriteAttribute(sw, "w:vMergeOrig", this.vMergeOrig.ToString());
            XmlHelper.WriteAttribute(sw, "w:author", this.author);
            XmlHelper.WriteAttribute(sw, "w:date", this.date);
            XmlHelper.WriteAttribute(sw, "r:id", this.id);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_AnnotationVMerge vMerge
        {
            get
            {
                return this.vMergeField;
            }
            set
            {
                this.vMergeField = value;
            }
        }

        [XmlIgnore]
        public bool vMergeSpecified
        {
            get
            {
                return this.vMergeFieldSpecified;
            }
            set
            {
                this.vMergeFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_AnnotationVMerge vMergeOrig
        {
            get
            {
                return this.vMergeOrigField;
            }
            set
            {
                this.vMergeOrigField = value;
            }
        }

        [XmlIgnore]
        public bool vMergeOrigSpecified
        {
            get
            {
                return this.vMergeOrigFieldSpecified;
            }
            set
            {
                this.vMergeOrigFieldSpecified = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_AnnotationVMerge
    {

    
        cont,

    
        rest,
    }



    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TcBorders
    {

        private CT_Border topField;

        private CT_Border leftField;

        private CT_Border bottomField;

        private CT_Border rightField;

        private CT_Border insideHField;

        private CT_Border insideVField;

        private CT_Border tl2brField;

        private CT_Border tr2blField;
        public static CT_TcBorders Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TcBorders ctObj = new CT_TcBorders();
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
                else if (childNode.LocalName == "insideH")
                    ctObj.insideH = CT_Border.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "insideV")
                    ctObj.insideV = CT_Border.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tl2br")
                    ctObj.tl2br = CT_Border.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tr2bl")
                    ctObj.tr2bl = CT_Border.Parse(childNode, namespaceManager);
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
            if (this.insideH != null)
                this.insideH.Write(sw, "insideH");
            if (this.insideV != null)
                this.insideV.Write(sw, "insideV");
            if (this.tl2br != null)
                this.tl2br.Write(sw, "tl2br");
            if (this.tr2bl != null)
                this.tr2bl.Write(sw, "tr2bl");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        public CT_TcBorders()
        {

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
        public CT_Border insideH
        {
            get
            {
                return this.insideHField;
            }
            set
            {
                this.insideHField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_Border insideV
        {
            get
            {
                return this.insideVField;
            }
            set
            {
                this.insideVField = value;
            }
        }

        [XmlElement(Order = 6)]
        public CT_Border tl2br
        {
            get
            {
                return this.tl2brField;
            }
            set
            {
                this.tl2brField = value;
            }
        }

        [XmlElement(Order = 7)]
        public CT_Border tr2bl
        {
            get
            {
                return this.tr2blField;
            }
            set
            {
                this.tr2blField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TcMar
    {

        private CT_TblWidth topField;

        private CT_TblWidth leftField;

        private CT_TblWidth bottomField;

        private CT_TblWidth rightField;

        public CT_TcMar()
        {
            //this.rightField = new CT_TblWidth();
            //this.bottomField = new CT_TblWidth();
            //this.leftField = new CT_TblWidth();
            //this.topField = new CT_TblWidth();
        }
        public static CT_TcMar Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TcMar ctObj = new CT_TcMar();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "top")
                    ctObj.top = CT_TblWidth.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "left")
                    ctObj.left = CT_TblWidth.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bottom")
                    ctObj.bottom = CT_TblWidth.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "right")
                    ctObj.right = CT_TblWidth.Parse(childNode, namespaceManager);
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
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_TblWidth top
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
        public CT_TblWidth left
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
        public CT_TblWidth bottom
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
        public CT_TblWidth right
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
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TcPr : CT_TcPrInner
    {

        private CT_TcPrChange tcPrChangeField;
        public static new CT_TcPr Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TcPr ctObj = new CT_TcPr();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "tcPrChange")
                    ctObj.tcPrChange = CT_TcPrChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cellIns")
                    ctObj.cellIns = CT_TrackChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cellDel")
                    ctObj.cellDel = CT_TrackChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cellMerge")
                    ctObj.cellMerge = CT_CellMergeTrackChange.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cnfStyle")
                    ctObj.cnfStyle = CT_Cnf.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tcW")
                    ctObj.tcW = CT_TblWidth.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "gridSpan")
                    ctObj.gridSpan = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "hMerge")
                    ctObj.hMerge = CT_HMerge.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "vMerge")
                    ctObj.vMerge = CT_VMerge.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tcBorders")
                    ctObj.tcBorders = CT_TcBorders.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "shd")
                    ctObj.shd = CT_Shd.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "noWrap")
                    ctObj.noWrap = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tcMar")
                    ctObj.tcMar = CT_TcMar.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "textDirection")
                    ctObj.textDirection = CT_TextDirection.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tcFitText")
                    ctObj.tcFitText = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "vAlign")
                    ctObj.vAlign = CT_VerticalJc.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "hideMark")
                    ctObj.hideMark = CT_OnOff.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }

        public CT_TblWidth AddNewTcW()
        {
            this.tcW = new CT_TblWidth();
            return this.tcW;
        }

        internal new void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.tcPrChange != null)
                this.tcPrChange.Write(sw, "tcPrChange");
            if (this.cellIns != null)
                this.cellIns.Write(sw, "cellIns");
            if (this.cellDel != null)
                this.cellDel.Write(sw, "cellDel");
            if (this.cellMerge != null)
                this.cellMerge.Write(sw, "cellMerge");
            if (this.cnfStyle != null)
                this.cnfStyle.Write(sw, "cnfStyle");
            if (this.tcW != null)
                this.tcW.Write(sw, "tcW");
            if (this.gridSpan != null)
                this.gridSpan.Write(sw, "gridSpan");
            if (this.hMerge != null)
                this.hMerge.Write(sw, "hMerge");
            if (this.vMerge != null)
                this.vMerge.Write(sw, "vMerge");
            if (this.tcBorders != null)
                this.tcBorders.Write(sw, "tcBorders");
            if (this.shd != null)
                this.shd.Write(sw, "shd");
            if (this.noWrap != null)
                this.noWrap.Write(sw, "noWrap");
            if (this.tcMar != null)
                this.tcMar.Write(sw, "tcMar");
            if (this.textDirection != null)
                this.textDirection.Write(sw, "textDirection");
            if (this.tcFitText != null)
                this.tcFitText.Write(sw, "tcFitText");
            if (this.vAlign != null)
                this.vAlign.Write(sw, "vAlign");
            if (this.hideMark != null)
                this.hideMark.Write(sw, "hideMark");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        public CT_TcPr()
        {
            //this.tcPrChangeField = new CT_TcPrChange();
        }

        [XmlElement(Order = 0)]
        public CT_TcPrChange tcPrChange
        {
            get
            {
                return this.tcPrChangeField;
            }
            set
            {
                this.tcPrChangeField = value;
            }
        }

    }

    

    [XmlInclude(typeof(CT_TcPrInner))]
    [XmlInclude(typeof(CT_TcPr))]
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TcPrBase
    {

        private CT_Cnf cnfStyleField;

        private CT_TblWidth tcWField;

        private CT_DecimalNumber gridSpanField;

        private CT_HMerge hMergeField;

        private CT_VMerge vMergeField;

        private CT_TcBorders tcBordersField;

        private CT_Shd shdField;

        private CT_OnOff noWrapField;

        private CT_TcMar tcMarField;

        private CT_TextDirection textDirectionField;

        private CT_OnOff tcFitTextField;

        private CT_VerticalJc vAlignField;

        private CT_OnOff hideMarkField;

        public CT_TcPrBase()
        {
        }

        [XmlElement(Order = 0)]
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

        [XmlElement(Order = 1)]
        public CT_TblWidth tcW
        {
            get
            {
                return this.tcWField;
            }
            set
            {
                this.tcWField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_DecimalNumber gridSpan
        {
            get
            {
                return this.gridSpanField;
            }
            set
            {
                this.gridSpanField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_HMerge hMerge
        {
            get
            {
                return this.hMergeField;
            }
            set
            {
                this.hMergeField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_VMerge vMerge
        {
            get
            {
                return this.vMergeField;
            }
            set
            {
                this.vMergeField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_TcBorders tcBorders
        {
            get
            {
                return this.tcBordersField;
            }
            set
            {
                this.tcBordersField = value;
            }
        }

        [XmlElement(Order = 6)]
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

        [XmlElement(Order = 7)]
        public CT_OnOff noWrap
        {
            get
            {
                return this.noWrapField;
            }
            set
            {
                this.noWrapField = value;
            }
        }

        [XmlElement(Order = 8)]
        public CT_TcMar tcMar
        {
            get
            {
                return this.tcMarField;
            }
            set
            {
                this.tcMarField = value;
            }
        }

        [XmlElement(Order = 9)]
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

        [XmlElement(Order = 10)]
        public CT_OnOff tcFitText
        {
            get
            {
                return this.tcFitTextField;
            }
            set
            {
                this.tcFitTextField = value;
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
        public CT_OnOff hideMark
        {
            get
            {
                return this.hideMarkField;
            }
            set
            {
                this.hideMarkField = value;
            }
        }

        public CT_Shd AddNewShd()
        {
            this.shdField = new CT_Shd();
            return this.shdField;
        }

        public bool IsSetShd()
        {
            return this.shdField != null;
        }

        public CT_VerticalJc AddNewVAlign()
        {
            this.vAlign = new CT_VerticalJc();
            return this.vAlign;
        }

        public CT_VMerge AddNewVMerge()
        {
            this.vMerge = new CT_VMerge();
            return this.vMerge;
        }

        public CT_TcBorders AddNewTcBorders()
        {
            this.tcBorders = new CT_TcBorders();
            return this.tcBorders;
        }

        public CT_HMerge AddNewHMerge()
        {
            this.hMerge = new CT_HMerge();
            return this.hMerge;
        }
        public CT_DecimalNumber AddNewGridspan()
        {
            this.gridSpanField = new CT_DecimalNumber();
            return this.gridSpanField;
        }
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_HMerge
    {

        private ST_Merge valField;

        private bool valFieldSpecified;
        public static CT_HMerge Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_HMerge ctObj = new CT_HMerge();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_Merge)Enum.Parse(typeof(ST_Merge), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_Merge val
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

        [XmlIgnore]
        public bool valSpecified
        {
            get
            {
                return this.valFieldSpecified;
            }
            set
            {
                this.valFieldSpecified = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_Merge
    {

    
        @continue,

    
        restart,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_VMerge
    {
        public CT_VMerge()
        {
            this.valField = ST_Merge.@continue;
        }
        private ST_Merge valField;

        private bool valFieldSpecified;
        public static CT_VMerge Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_VMerge ctObj = new CT_VMerge();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_Merge)Enum.Parse(typeof(ST_Merge), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            if(valField != ST_Merge.@continue)
                XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write("/>");
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_Merge val
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

        [XmlIgnore]
        public bool valSpecified
        {
            get
            {
                return this.valFieldSpecified;
            }
            set
            {
                this.valFieldSpecified = value;
            }
        }
    }

    
    #endregion
    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Column
    {

        private ulong wField;

        private bool wFieldSpecified;

        private ulong spaceField;

        private bool spaceFieldSpecified;
        public static CT_Column Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Column ctObj = new CT_Column();
            ctObj.w = XmlHelper.ReadULong(node.Attributes["w:w"]);
            ctObj.space = XmlHelper.ReadULong(node.Attributes["w:space"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:w", this.w);
            XmlHelper.WriteAttribute(sw, "w:space", this.space);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong w
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

        [XmlIgnore]
        public bool wSpecified
        {
            get
            {
                return this.wFieldSpecified;
            }
            set
            {
                this.wFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong space
        {
            get
            {
                return this.spaceField;
            }
            set
            {
                this.spaceField = value;
            }
        }

        [XmlIgnore]
        public bool spaceSpecified
        {
            get
            {
                return this.spaceFieldSpecified;
            }
            set
            {
                this.spaceFieldSpecified = value;
            }
        }
    }
    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Columns
    {

        private List<CT_Column> colField;

        private ST_OnOff equalWidthField;

        private bool equalWidthFieldSpecified;

        private ulong spaceField;

        private bool spaceFieldSpecified;

        private string numField;

        private ST_OnOff sepField;

        private bool sepFieldSpecified;

        public CT_Columns()
        {
            //this.colField = new List<CT_Column>();
            this.equalWidthField = ST_OnOff.off;
            this.sepField = ST_OnOff.off;
        }
        public static CT_Columns Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Columns ctObj = new CT_Columns();
            if (node.Attributes["w:equalWidth"] != null)
                ctObj.equalWidth = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:equalWidth"].Value);
            ctObj.space = XmlHelper.ReadULong(node.Attributes["w:space"]);
            ctObj.num = XmlHelper.ReadString(node.Attributes["w:num"]);
            if (node.Attributes["w:sep"] != null)
                ctObj.sep = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:sep"].Value);
            ctObj.col = new List<CT_Column>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "col")
                    ctObj.col.Add(CT_Column.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            if(this.equalWidth!= ST_OnOff.off)
                XmlHelper.WriteAttribute(sw, "w:equalWidth", this.equalWidth.ToString());
            XmlHelper.WriteAttribute(sw, "w:space", this.space);
            XmlHelper.WriteAttribute(sw, "w:num", this.num);
            if(this.sep!= ST_OnOff.off)
                XmlHelper.WriteAttribute(sw, "w:sep", this.sep.ToString());
            sw.Write(">");
            if (this.col != null)
            {
                foreach (CT_Column x in this.col)
                {
                    x.Write(sw, "col");
                }
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }


        [XmlElement("col", Order = 0)]
        public List<CT_Column> col
        {
            get
            {
                return this.colField;
            }
            set
            {
                this.colField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff equalWidth
        {
            get
            {
                return this.equalWidthField;
            }
            set
            {
                this.equalWidthField = value;
            }
        }

        [XmlIgnore]
        public bool equalWidthSpecified
        {
            get
            {
                return this.equalWidthFieldSpecified;
            }
            set
            {
                this.equalWidthFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong space
        {
            get
            {
                return this.spaceField;
            }
            set
            {
                this.spaceField = value;
            }
        }

        [XmlIgnore]
        public bool spaceSpecified
        {
            get
            {
                return this.spaceFieldSpecified;
            }
            set
            {
                this.spaceFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string num
        {
            get
            {
                return this.numField;
            }
            set
            {
                this.numField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff sep
        {
            get
            {
                return this.sepField;
            }
            set
            {
                this.sepField = value;
            }
        }

        [XmlIgnore]
        public bool sepSpecified
        {
            get
            {
                return this.sepFieldSpecified;
            }
            set
            {
                this.sepFieldSpecified = value;
            }
        }
    }


#region Table Row
    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Row
    {

        private CT_TblPrEx tblPrExField;

        private CT_TrPr trPrField;

        private ArrayList itemsField;

        private List<ItemsChoiceTableRowType> itemsElementNameField;

        private byte[] rsidRPrField;

        private byte[] rsidRField;

        private byte[] rsidDelField;

        private byte[] rsidTrField;

        public CT_Row()
        {
            this.itemsElementNameField = new List<ItemsChoiceTableRowType>();
            this.itemsField = new ArrayList();
        }
        object parent;
        [XmlIgnore]
        public object Parent
        {
            get { return parent; }
        }
        public static CT_Row Parse(XmlNode node, XmlNamespaceManager namespaceManager, object parent)
        {
            if (node == null)
                return null;
            CT_Row ctObj = new CT_Row();
            if (parent != null)
                ctObj.parent = parent;
            ctObj.rsidRPr = XmlHelper.ReadBytes(node.Attributes["w:rsidRPr"]);
            ctObj.rsidR = XmlHelper.ReadBytes(node.Attributes["w:rsidR"]);
            ctObj.rsidDel = XmlHelper.ReadBytes(node.Attributes["w:rsidDel"]);
            ctObj.rsidTr = XmlHelper.ReadBytes(node.Attributes["w:rsidTr"]);

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "tblPrEx")
                    ctObj.tblPrEx = CT_TblPrEx.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "trPr")
                    ctObj.trPr = CT_TrPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bookmarkStart")
                {
                    ctObj.Items.Add(CT_Bookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.bookmarkStart);
                }
                else if (childNode.LocalName == "commentRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.commentRangeEnd);
                }
                else if (childNode.LocalName == "commentRangeStart")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.commentRangeStart);
                }
                else if (childNode.LocalName == "customXmlInsRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.customXmlInsRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.customXmlMoveFromRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.customXmlMoveToRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.customXmlMoveToRangeStart);
                }
                else if (childNode.LocalName == "del")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.del);
                }
                else if (childNode.LocalName == "moveTo")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.moveTo);
                }
                else if (childNode.LocalName == "oMath")
                {
                    ctObj.Items.Add(CT_OMath.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.oMath);
                }
                else if (childNode.LocalName == "oMathPara")
                {
                    ctObj.Items.Add(CT_OMathPara.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.oMathPara);
                }
                else if (childNode.LocalName == "bookmarkEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.bookmarkEnd);
                }
                else if (childNode.LocalName == "customXml")
                {
                    ctObj.Items.Add(CT_CustomXmlCell.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.customXml);
                }
                else if (childNode.LocalName == "customXmlDelRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.customXmlDelRangeStart);
                }
                else if (childNode.LocalName == "customXmlInsRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.customXmlInsRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.customXmlMoveFromRangeEnd);
                }
                else if (childNode.LocalName == "ins")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.ins);
                }
                else if (childNode.LocalName == "moveFrom")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.moveFrom);
                }
                else if (childNode.LocalName == "moveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.moveFromRangeEnd);
                }
                else if (childNode.LocalName == "moveFromRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.moveFromRangeStart);
                }
                else if (childNode.LocalName == "customXmlDelRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.customXmlDelRangeEnd);
                }
                else if (childNode.LocalName == "moveToRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.moveToRangeEnd);
                }
                else if (childNode.LocalName == "moveToRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.moveToRangeStart);
                }
                else if (childNode.LocalName == "permEnd")
                {
                    ctObj.Items.Add(CT_Perm.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.permEnd);
                }
                else if (childNode.LocalName == "permStart")
                {
                    ctObj.Items.Add(CT_PermStart.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.permStart);
                }
                else if (childNode.LocalName == "proofErr")
                {
                    ctObj.Items.Add(CT_ProofErr.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.proofErr);
                }
                else if (childNode.LocalName == "sdt")
                {
                    ctObj.Items.Add(CT_SdtCell.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.sdt);
                }
                else if (childNode.LocalName == "tc")
                {
                    ctObj.Items.Add(CT_Tc.Parse(childNode, namespaceManager, ctObj));
                    ctObj.ItemsElementName.Add(ItemsChoiceTableRowType.tc);
                }
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:rsidRPr", this.rsidRPr);
            XmlHelper.WriteAttribute(sw, "w:rsidR", this.rsidR);
            XmlHelper.WriteAttribute(sw, "w:rsidDel", this.rsidDel);
            XmlHelper.WriteAttribute(sw, "w:rsidTr", this.rsidTr);
            sw.Write(">");
            if (this.tblPrEx != null)
                this.tblPrEx.Write(sw, "tblPrEx");
            if (this.trPr != null)
                this.trPr.Write(sw, "trPr");
            foreach (object o in this.Items)
            {
                if (o is CT_Bookmark)
                    ((CT_Bookmark)o).Write(sw, "bookmarkStart");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeEnd");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeStart");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlInsRangeStart");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveFromRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlMoveToRangeEnd");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveToRangeStart");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "del");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "moveTo");
                else if (o is CT_OMath)
                    ((CT_OMath)o).Write(sw, "oMath");
                else if (o is CT_OMathPara)
                    ((CT_OMathPara)o).Write(sw, "oMathPara");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "bookmarkEnd");
                else if (o is CT_CustomXmlCell)
                    ((CT_CustomXmlCell)o).Write(sw, "customXml");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlDelRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlInsRangeEnd");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlMoveFromRangeEnd");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "ins");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "moveFrom");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "moveFromRangeEnd");
                else if (o is CT_MoveBookmark)
                    ((CT_MoveBookmark)o).Write(sw, "moveFromRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlDelRangeEnd");
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
                else if (o is CT_SdtCell)
                    ((CT_SdtCell)o).Write(sw, "sdt");
                else if (o is CT_Tc)
                    ((CT_Tc)o).Write(sw, "tc");
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }
        public void RemoveTc(int pos)
        {
            RemoveObject(ItemsChoiceTableRowType.tc, pos);
        }
        [XmlElement(Order = 0)]
        public CT_TblPrEx tblPrEx
        {
            get
            {
                return this.tblPrExField;
            }
            set
            {
                this.tblPrExField = value;
            }
        }

        [XmlElement(Order = 1)]
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

        [XmlElement("oMath", typeof(CT_OMath), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 2)]
        [XmlElement("oMathPara", typeof(CT_OMathPara), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 2)]
        [XmlElement("bookmarkEnd", typeof(CT_MarkupRange), Order = 2)]
        [XmlElement("bookmarkStart", typeof(CT_Bookmark), Order = 2)]
        [XmlElement("commentRangeEnd", typeof(CT_MarkupRange), Order = 2)]
        [XmlElement("commentRangeStart", typeof(CT_MarkupRange), Order = 2)]
        [XmlElement("customXml", typeof(CT_CustomXmlCell), Order = 2)]
        [XmlElement("customXmlDelRangeEnd", typeof(CT_Markup), Order = 2)]
        [XmlElement("customXmlDelRangeStart", typeof(CT_TrackChange), Order = 2)]
        [XmlElement("customXmlInsRangeEnd", typeof(CT_Markup), Order = 2)]
        [XmlElement("customXmlInsRangeStart", typeof(CT_TrackChange), Order = 2)]
        [XmlElement("customXmlMoveFromRangeEnd", typeof(CT_Markup), Order = 2)]
        [XmlElement("customXmlMoveFromRangeStart", typeof(CT_TrackChange), Order = 2)]
        [XmlElement("customXmlMoveToRangeEnd", typeof(CT_Markup), Order = 2)]
        [XmlElement("customXmlMoveToRangeStart", typeof(CT_TrackChange), Order = 2)]
        [XmlElement("del", typeof(CT_RunTrackChange), Order = 2)]
        [XmlElement("ins", typeof(CT_RunTrackChange), Order = 2)]
        [XmlElement("moveFrom", typeof(CT_RunTrackChange), Order = 2)]
        [XmlElement("moveFromRangeEnd", typeof(CT_MarkupRange), Order = 2)]
        [XmlElement("moveFromRangeStart", typeof(CT_MoveBookmark), Order = 2)]
        [XmlElement("moveTo", typeof(CT_RunTrackChange), Order = 2)]
        [XmlElement("moveToRangeEnd", typeof(CT_MarkupRange), Order = 2)]
        [XmlElement("moveToRangeStart", typeof(CT_MoveBookmark), Order = 2)]
        [XmlElement("permEnd", typeof(CT_Perm), Order = 2)]
        [XmlElement("permStart", typeof(CT_PermStart), Order = 2)]
        [XmlElement("proofErr", typeof(CT_ProofErr), Order = 2)]
        [XmlElement("sdt", typeof(CT_SdtCell), Order = 2)]
        [XmlElement("tc", typeof(CT_Tc), Order = 2)]
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

        [XmlElement("ItemsElementName", Order = 3)]
        [XmlIgnore]
        public List<ItemsChoiceTableRowType> ItemsElementName
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
        public byte[] rsidTr
        {
            get
            {
                return this.rsidTrField;
            }
            set
            {
                this.rsidTrField = value;
            }
        }
        private CT_Tbl _table;
        [XmlIgnore]
        public CT_Tbl Table
        {
            get { return _table; }
            set { _table = value; }
        }

        public IList<CT_Tc> GetTcList()
        {
            return GetObjectList<CT_Tc>(ItemsChoiceTableRowType.tc);
        }

        public bool IsSetTrPr()
        {
            if (this.trPrField == null)
                return false;
            return this.trPrField.Items.Count > 0;
        }

        public CT_TrPr AddNewTrPr()
        {
            if (this.trPrField == null)
                this.trPrField = new CT_TrPr();
            return this.trPrField;
        }

        public CT_Tc AddNewTc()
        {
            return AddNewObject<CT_Tc>(ItemsChoiceTableRowType.tc);
        }

        public int SizeOfTcArray()
        {
            return SizeOfArray(ItemsChoiceTableRowType.tc);
        }

        public CT_Tc GetTcArray(int p)
        {
            return GetObjectArray<CT_Tc>(p, ItemsChoiceTableRowType.tc);
        }
        #region Generic methods for object operation

        private List<T> GetObjectList<T>(ItemsChoiceTableRowType type) where T : class
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
        private int SizeOfArray(ItemsChoiceTableRowType type)
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
        private T GetObjectArray<T>(int p, ItemsChoiceTableRowType type) where T : class
        {
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                if (pos < 0 || pos >= this.itemsField.Count)
                    return null;
                return itemsField[pos] as T;
            }
        }
        private T InsertNewObject<T>(ItemsChoiceTableRowType type, int p) where T : class, new()
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
        private T AddNewObject<T>(ItemsChoiceTableRowType type) where T : class, new()
        {
            T t = new T();
            lock (this)
            {
                this.itemsElementNameField.Add(type);
                this.itemsField.Add(t);
            }
            return t;
        }
        private void SetObject<T>(ItemsChoiceTableRowType type, int p, T obj) where T : class
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
        private int GetObjectIndex(ItemsChoiceTableRowType type, int p)
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
        private void RemoveObject(ItemsChoiceTableRowType type, int p)
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


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceTableRowType
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

#endregion
}