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
    public class CT_SdtContentCell
    {
        public override string ToString()
        {
            string text = string.Empty;
            using (MemoryStream ms = new MemoryStream())
            {
                using (StreamWriter sw = new StreamWriter(ms))
                {
                    this.Write(sw, "sdtContent");
                    sw.Flush();
                    ms.Position = 0;
                    using (StreamReader sr = new StreamReader(ms))
                    {
                        text = sr.ReadToEnd();
                    }
                }
            }
            return text;
        }
        private ArrayList itemsField;

        private List<ItemsChoiceType23> itemsElementNameField;

        public CT_SdtContentCell()
        {
            this.itemsElementNameField = new List<ItemsChoiceType23>();
            this.itemsField = new ArrayList();
        }
        public static CT_SdtContentCell Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SdtContentCell ctObj = new CT_SdtContentCell();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "permStart")
                {
                    ctObj.Items.Add(CT_PermStart.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.permStart);
                }
                else if (childNode.LocalName == "sdt")
                {
                    ctObj.Items.Add(CT_SdtCell.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.sdt);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.customXmlMoveToRangeEnd);
                }
                else if (childNode.LocalName == "proofErr")
                {
                    ctObj.Items.Add(CT_ProofErr.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.proofErr);
                }
                else if (childNode.LocalName == "customXmlDelRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.customXmlDelRangeEnd);
                }
                else if (childNode.LocalName == "moveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.moveFromRangeEnd);
                }
                else if (childNode.LocalName == "moveFromRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.moveFromRangeStart);
                }
                else if (childNode.LocalName == "moveTo")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.moveTo);
                }
                else if (childNode.LocalName == "moveToRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.moveToRangeEnd);
                }
                else if (childNode.LocalName == "moveToRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.moveToRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.customXmlMoveToRangeStart);
                }
                else if (childNode.LocalName == "tc")
                {
                    ctObj.Items.Add(CT_Tc.Parse(childNode, namespaceManager, ctObj));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.tc);
                }
                else if (childNode.LocalName == "del")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.del);
                }
                else if (childNode.LocalName == "permEnd")
                {
                    ctObj.Items.Add(CT_Perm.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.permEnd);
                }
                else if (childNode.LocalName == "commentRangeStart")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.commentRangeStart);
                }
                else if (childNode.LocalName == "moveFrom")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.moveFrom);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.customXmlMoveFromRangeStart);
                }
                else if (childNode.LocalName == "customXml")
                {
                    ctObj.Items.Add(CT_CustomXmlCell.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.customXml);
                }
                else if (childNode.LocalName == "oMath")
                {
                    ctObj.Items.Add(CT_OMath.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.oMath);
                }
                else if (childNode.LocalName == "customXmlDelRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.customXmlDelRangeStart);
                }
                else if (childNode.LocalName == "customXmlInsRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.customXmlInsRangeEnd);
                }
                else if (childNode.LocalName == "ins")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.ins);
                }
                else if (childNode.LocalName == "customXmlInsRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.customXmlInsRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.customXmlMoveFromRangeEnd);
                }
                else if (childNode.LocalName == "oMathPara")
                {
                    ctObj.Items.Add(CT_OMathPara.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.oMathPara);
                }
                else if (childNode.LocalName == "bookmarkEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.bookmarkEnd);
                }
                else if (childNode.LocalName == "bookmarkStart")
                {
                    ctObj.Items.Add(CT_Bookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.bookmarkStart);
                }
                else if (childNode.LocalName == "commentRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType23.commentRangeEnd);
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
                if (o is CT_PermStart)
                    ((CT_PermStart)o).Write(sw, "permStart");
                else if (o is CT_SdtCell)
                    ((CT_SdtCell)o).Write(sw, "sdt");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlMoveToRangeEnd");
                else if (o is CT_ProofErr)
                    ((CT_ProofErr)o).Write(sw, "proofErr");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlDelRangeEnd");
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
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveToRangeStart");
                else if (o is CT_Tc)
                    ((CT_Tc)o).Write(sw, "tc");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "del");
                else if (o is CT_Perm)
                    ((CT_Perm)o).Write(sw, "permEnd");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeStart");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "moveFrom");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveFromRangeStart");
                else if (o is CT_CustomXmlCell)
                    ((CT_CustomXmlCell)o).Write(sw, "customXml");
                else if (o is CT_OMath)
                    ((CT_OMath)o).Write(sw, "oMath");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlDelRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlInsRangeEnd");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "ins");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlInsRangeStart");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "customXmlMoveFromRangeEnd");
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

        [XmlElement("oMath", typeof(CT_OMath), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("oMathPara", typeof(CT_OMathPara), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("bookmarkEnd", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("bookmarkStart", typeof(CT_Bookmark), Order = 0)]
        [XmlElement("commentRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("commentRangeStart", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("customXml", typeof(CT_CustomXmlCell), Order = 0)]
        [XmlElement("customXmlDelRangeEnd", typeof(CT_Markup), Order = 0)]
        [XmlElement("customXmlDelRangeStart", typeof(CT_TrackChange), Order = 0)]
        [XmlElement("customXmlInsRangeEnd", typeof(CT_Markup), Order = 0)]
        [XmlElement("customXmlInsRangeStart", typeof(CT_TrackChange), Order = 0)]
        [XmlElement("customXmlMoveFromRangeEnd", typeof(CT_Markup), Order = 0)]
        [XmlElement("customXmlMoveFromRangeStart", typeof(CT_TrackChange), Order = 0)]
        [XmlElement("customXmlMoveToRangeEnd", typeof(CT_Markup), Order = 0)]
        [XmlElement("customXmlMoveToRangeStart", typeof(CT_TrackChange), Order = 0)]
        [XmlElement("del", typeof(CT_RunTrackChange), Order = 0)]
        [XmlElement("ins", typeof(CT_RunTrackChange), Order = 0)]
        [XmlElement("moveFrom", typeof(CT_RunTrackChange), Order = 0)]
        [XmlElement("moveFromRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("moveFromRangeStart", typeof(CT_MoveBookmark), Order = 0)]
        [XmlElement("moveTo", typeof(CT_RunTrackChange), Order = 0)]
        [XmlElement("moveToRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("moveToRangeStart", typeof(CT_MoveBookmark), Order = 0)]
        [XmlElement("permEnd", typeof(CT_Perm), Order = 0)]
        [XmlElement("permStart", typeof(CT_PermStart), Order = 0)]
        [XmlElement("proofErr", typeof(CT_ProofErr), Order = 0)]
        [XmlElement("sdt", typeof(CT_SdtCell), Order = 0)]
        [XmlElement("tc", typeof(CT_Tc), Order = 0)]
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
        public List<ItemsChoiceType23> ItemsElementName
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
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType23
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
    public class CT_SdtBlock
    {

        private CT_SdtPr sdtPrField;

        private List<CT_RPr> sdtEndPrField;
        private CT_SdtContentBlock sdtContentField;

        public CT_SdtBlock()
        {
            //this.sdtContentField = new CT_SdtContentBlock();
            //this.sdtEndPrField = new List<CT_RPr>();
            //this.sdtPrField = new CT_SdtPr();
        }
        public static CT_SdtBlock Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SdtBlock ctObj = new CT_SdtBlock();
            ctObj.sdtEndPr = new List<CT_RPr>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "sdtPr")
                    ctObj.sdtPr = CT_SdtPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sdtContent")
                    ctObj.sdtContent = CT_SdtContentBlock.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sdtEndPr")
                    ctObj.sdtEndPr.Add(CT_RPr.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.sdtPr != null)
                this.sdtPr.Write(sw, "sdtPr");
            if (this.sdtContent != null)
                this.sdtContent.Write(sw, "sdtContent");
            if (this.sdtEndPr != null)
            {
                foreach (CT_RPr x in this.sdtEndPr)
                {
                    x.Write(sw, "sdtEndPr");
                }
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_SdtPr sdtPr
        {
            get
            {
                return this.sdtPrField;
            }
            set
            {
                this.sdtPrField = value;
            }
        }

        [XmlArray(Order = 1)]
        [XmlArrayItem("rPr", IsNullable = false)]
        public List<CT_RPr> sdtEndPr
        {
            get
            {
                return this.sdtEndPrField;
            }
            set
            {
                this.sdtEndPrField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_SdtContentBlock sdtContent
        {
            get
            {
                return this.sdtContentField;
            }
            set
            {
                this.sdtContentField = value;
            }
        }

        public CT_SdtPr AddNewSdtPr()
        {
            if (this.sdtPrField == null)
                this.sdtPrField = new CT_SdtPr();
            return this.sdtPrField;
        }

        public CT_SdtEndPr AddNewSdtEndPr()
        {
            CT_SdtEndPr endPr = new CT_SdtEndPr();
            this.sdtEndPrField = endPr.Items;
            return endPr;
        }

        public CT_SdtContentBlock AddNewSdtContent()
        {
            if (this.sdtContentField == null)
                this.sdtContentField = new CT_SdtContentBlock();
            return this.sdtContentField;
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SdtRun
    {

        private CT_SdtPr sdtPrField;

        private List<CT_RPr> sdtEndPrField;

        private CT_SdtContentRun sdtContentField;

        public CT_SdtRun()
        {
            //this.sdtContentField = new CT_SdtContentRun();
            //this.sdtEndPrField = new List<CT_RPr>();
            //this.sdtPrField = new CT_SdtPr();
        }
        public static CT_SdtRun Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SdtRun ctObj = new CT_SdtRun();
            ctObj.sdtEndPr = new List<CT_RPr>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "sdtPr")
                    ctObj.sdtPr = CT_SdtPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sdtContent")
                    ctObj.sdtContent = CT_SdtContentRun.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sdtEndPr")
                    ctObj.sdtEndPr.Add(CT_RPr.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.sdtPr != null)
                this.sdtPr.Write(sw, "sdtPr");
            if (this.sdtContent != null)
                this.sdtContent.Write(sw, "sdtContent");
            if (this.sdtEndPr != null)
            {
                foreach (CT_RPr x in this.sdtEndPr)
                {
                    x.Write(sw, "sdtEndPr");
                }
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_SdtPr sdtPr
        {
            get
            {
                return this.sdtPrField;
            }
            set
            {
                this.sdtPrField = value;
            }
        }

        [XmlArray(Order = 1)]
        [XmlArrayItem("rPr", IsNullable = false)]
        public List<CT_RPr> sdtEndPr
        {
            get
            {
                return this.sdtEndPrField;
            }
            set
            {
                this.sdtEndPrField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_SdtContentRun sdtContent
        {
            get
            {
                return this.sdtContentField;
            }
            set
            {
                this.sdtContentField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SdtCell
    {

        private CT_SdtPr sdtPrField;

        private List<CT_RPr> sdtEndPrField;

        private CT_SdtContentCell sdtContentField;

        public CT_SdtCell()
        {
            //this.sdtContentField = new CT_SdtContentCell();
            //this.sdtEndPrField = new List<CT_RPr>();
            //this.sdtPrField = new CT_SdtPr();
        }
        public static CT_SdtCell Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SdtCell ctObj = new CT_SdtCell();
            ctObj.sdtEndPr = new List<CT_RPr>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "sdtPr")
                    ctObj.sdtPr = CT_SdtPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sdtContent")
                    ctObj.sdtContent = CT_SdtContentCell.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sdtEndPr")
                    ctObj.sdtEndPr.Add(CT_RPr.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.sdtPr != null)
                this.sdtPr.Write(sw, "sdtPr");
            if (this.sdtContent != null)
                this.sdtContent.Write(sw, "sdtContent");
            if (this.sdtEndPr != null)
            {
                foreach (CT_RPr x in this.sdtEndPr)
                {
                    x.Write(sw, "sdtEndPr");
                }
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_SdtPr sdtPr
        {
            get
            {
                return this.sdtPrField;
            }
            set
            {
                this.sdtPrField = value;
            }
        }

        [XmlArray(Order = 1)]
        [XmlArrayItem("rPr", IsNullable = false)]
        public List<CT_RPr> sdtEndPr
        {
            get
            {
                return this.sdtEndPrField;
            }
            set
            {
                this.sdtEndPrField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_SdtContentCell sdtContent
        {
            get
            {
                return this.sdtContentField;
            }
            set
            {
                this.sdtContentField = value;
            }
        }
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SdtComboBox
    {
        public static CT_SdtComboBox Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SdtComboBox ctObj = new CT_SdtComboBox();
            ctObj.lastValue = XmlHelper.ReadString(node.Attributes["w:lastValue"]);
            ctObj.listItem = new List<CT_SdtListItem>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "listItem")
                    ctObj.listItem.Add(CT_SdtListItem.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:lastValue", this.lastValue);
            sw.Write(">");
            if (this.listItem != null)
            {
                foreach (CT_SdtListItem x in this.listItem)
                {
                    x.Write(sw, "listItem");
                }
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        private List<CT_SdtListItem> listItemField;

        private string lastValueField;

        public CT_SdtComboBox()
        {
            //this.listItemField = new List<CT_SdtListItem>();
        }

        [XmlElement("listItem", Order = 0)]
        public List<CT_SdtListItem> listItem
        {
            get
            {
                return this.listItemField;
            }
            set
            {
                this.listItemField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string lastValue
        {
            get
            {
                return this.lastValueField;
            }
            set
            {
                this.lastValueField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SdtDocPart
    {

        private CT_String docPartGalleryField;

        private CT_String docPartCategoryField;

        private CT_OnOff docPartUniqueField;

        public CT_SdtDocPart()
        {
            //this.docPartUniqueField = new CT_OnOff();
            //this.docPartCategoryField = new CT_String();
            //this.docPartGalleryField = new CT_String();
        }
        public static CT_SdtDocPart Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SdtDocPart ctObj = new CT_SdtDocPart();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "docPartGallery")
                    ctObj.docPartGallery = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "docPartCategory")
                    ctObj.docPartCategory = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "docPartUnique")
                    ctObj.docPartUnique = CT_OnOff.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.docPartGallery != null)
                this.docPartGallery.Write(sw, "docPartGallery");
            if (this.docPartCategory != null)
                this.docPartCategory.Write(sw, "docPartCategory");
            if (this.docPartUnique != null)
                this.docPartUnique.Write(sw, "docPartUnique");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_String docPartGallery
        {
            get
            {
                return this.docPartGalleryField;
            }
            set
            {
                this.docPartGalleryField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_String docPartCategory
        {
            get
            {
                return this.docPartCategoryField;
            }
            set
            {
                this.docPartCategoryField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_OnOff docPartUnique
        {
            get
            {
                return this.docPartUniqueField;
            }
            set
            {
                this.docPartUniqueField = value;
            }
        }

        public CT_String AddNewDocPartGallery()
        {
            if (this.docPartGalleryField == null)
                this.docPartGalleryField = new CT_String();
            return this.docPartGalleryField;
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SdtDropDownList
    {

        private List<CT_SdtListItem> listItemField;

        private string lastValueField;

        public CT_SdtDropDownList()
        {
            //this.listItemField = new List<CT_SdtListItem>();
        }
        public static CT_SdtDropDownList Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SdtDropDownList ctObj = new CT_SdtDropDownList();
            ctObj.lastValue = XmlHelper.ReadString(node.Attributes["w:lastValue"]);
            ctObj.listItem = new List<CT_SdtListItem>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "listItem")
                    ctObj.listItem.Add(CT_SdtListItem.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:lastValue", this.lastValue);
            sw.Write(">");
            if (this.listItem != null)
            {
                foreach (CT_SdtListItem x in this.listItem)
                {
                    x.Write(sw, "listItem");
                }
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement("listItem", Order = 0)]
        public List<CT_SdtListItem> listItem
        {
            get
            {
                return this.listItemField;
            }
            set
            {
                this.listItemField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string lastValue
        {
            get
            {
                return this.lastValueField;
            }
            set
            {
                this.lastValueField = value;
            }
        }
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SdtContentBlock
    {

        private ArrayList itemsField;

        private List<ItemsChoiceType19> itemsElementNameField;

        public CT_SdtContentBlock()
        {
            this.itemsElementNameField = new List<ItemsChoiceType19>();
            this.itemsField = new ArrayList();
        }
        public static CT_SdtContentBlock Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SdtContentBlock ctObj = new CT_SdtContentBlock();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "bookmarkStart")
                {
                    ctObj.Items.Add(CT_Bookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.bookmarkStart);
                }
                else if (childNode.LocalName == "oMath")
                {
                    ctObj.Items.Add(CT_OMath.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.oMath);
                }
                else if (childNode.LocalName == "oMathPara")
                {
                    ctObj.Items.Add(CT_OMathPara.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.oMathPara);
                }
                else if (childNode.LocalName == "bookmarkEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.bookmarkEnd);
                }
                else if (childNode.LocalName == "commentRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.commentRangeEnd);
                }
                else if (childNode.LocalName == "commentRangeStart")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.commentRangeStart);
                }
                else if (childNode.LocalName == "customXml")
                {
                    ctObj.Items.Add(CT_CustomXmlBlock.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.customXml);
                }
                else if (childNode.LocalName == "customXmlDelRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.customXmlDelRangeEnd);
                }
                else if (childNode.LocalName == "customXmlDelRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.customXmlDelRangeStart);
                }
                else if (childNode.LocalName == "customXmlInsRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.customXmlInsRangeEnd);
                }
                else if (childNode.LocalName == "customXmlInsRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.customXmlInsRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.customXmlMoveFromRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.customXmlMoveFromRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.customXmlMoveToRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.customXmlMoveToRangeStart);
                }
                else if (childNode.LocalName == "del")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.del);
                }
                else if (childNode.LocalName == "ins")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.ins);
                }
                else if (childNode.LocalName == "moveFrom")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.moveFrom);
                }
                else if (childNode.LocalName == "moveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.moveFromRangeEnd);
                }
                else if (childNode.LocalName == "moveFromRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.moveFromRangeStart);
                }
                else if (childNode.LocalName == "moveTo")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.moveTo);
                }
                else if (childNode.LocalName == "moveToRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.moveToRangeEnd);
                }
                else if (childNode.LocalName == "moveToRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.moveToRangeStart);
                }
                else if (childNode.LocalName == "p")
                {
                    ctObj.Items.Add(CT_P.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.p);
                }
                else if (childNode.LocalName == "permEnd")
                {
                    ctObj.Items.Add(CT_Perm.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.permEnd);
                }
                else if (childNode.LocalName == "permStart")
                {
                    ctObj.Items.Add(CT_PermStart.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.permStart);
                }
                else if (childNode.LocalName == "proofErr")
                {
                    ctObj.Items.Add(CT_ProofErr.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.proofErr);
                }
                else if (childNode.LocalName == "sdt")
                {
                    ctObj.Items.Add(CT_SdtBlock.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.sdt);
                }
                else if (childNode.LocalName == "tbl")
                {
                    ctObj.Items.Add(CT_Tbl.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType19.tbl);
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
                if (o is CT_Bookmark)
                    ((CT_Bookmark)o).Write(sw, "bookmarkStart");
                else if (o is CT_OMath)
                    ((CT_OMath)o).Write(sw, "oMath");
                else if (o is CT_OMathPara)
                    ((CT_OMathPara)o).Write(sw, "oMathPara");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "bookmarkEnd");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeEnd");
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
                else if (o is CT_MoveBookmark)
                    ((CT_MoveBookmark)o).Write(sw, "moveFromRangeStart");
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
                else if (o is CT_PermStart)
                    ((CT_PermStart)o).Write(sw, "permStart");
                else if (o is CT_ProofErr)
                    ((CT_ProofErr)o).Write(sw, "proofErr");
                else if (o is CT_SdtBlock)
                    ((CT_SdtBlock)o).Write(sw, "sdt");
                else if (o is CT_Tbl)
                    ((CT_Tbl)o).Write(sw, "tbl");
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement("oMath", typeof(CT_OMath), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("oMathPara", typeof(CT_OMathPara), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("bookmarkEnd", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("bookmarkStart", typeof(CT_Bookmark), Order = 0)]
        [XmlElement("commentRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("commentRangeStart", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("customXml", typeof(CT_CustomXmlBlock), Order = 0)]
        [XmlElement("customXmlDelRangeEnd", typeof(CT_Markup), Order = 0)]
        [XmlElement("customXmlDelRangeStart", typeof(CT_TrackChange), Order = 0)]
        [XmlElement("customXmlInsRangeEnd", typeof(CT_Markup), Order = 0)]
        [XmlElement("customXmlInsRangeStart", typeof(CT_TrackChange), Order = 0)]
        [XmlElement("customXmlMoveFromRangeEnd", typeof(CT_Markup), Order = 0)]
        [XmlElement("customXmlMoveFromRangeStart", typeof(CT_TrackChange), Order = 0)]
        [XmlElement("customXmlMoveToRangeEnd", typeof(CT_Markup), Order = 0)]
        [XmlElement("customXmlMoveToRangeStart", typeof(CT_TrackChange), Order = 0)]
        [XmlElement("del", typeof(CT_RunTrackChange), Order = 0)]
        [XmlElement("ins", typeof(CT_RunTrackChange), Order = 0)]
        [XmlElement("moveFrom", typeof(CT_RunTrackChange), Order = 0)]
        [XmlElement("moveFromRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("moveFromRangeStart", typeof(CT_MoveBookmark), Order = 0)]
        [XmlElement("moveTo", typeof(CT_RunTrackChange), Order = 0)]
        [XmlElement("moveToRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("moveToRangeStart", typeof(CT_MoveBookmark), Order = 0)]
        [XmlElement("p", typeof(CT_P), Order = 0)]
        [XmlElement("permEnd", typeof(CT_Perm), Order = 0)]
        [XmlElement("permStart", typeof(CT_PermStart), Order = 0)]
        [XmlElement("proofErr", typeof(CT_ProofErr), Order = 0)]
        [XmlElement("sdt", typeof(CT_SdtBlock), Order = 0)]
        [XmlElement("tbl", typeof(CT_Tbl), Order = 0)]
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
        public List<ItemsChoiceType19> ItemsElementName
        {
            get
            {
                return this.itemsElementNameField ;
            }
            set
            {
               this.itemsElementNameField = value;
            }
        }

        public CT_P AddNewP()
        {
            return AddNewObject<CT_P>(ItemsChoiceType19.p);
        }
        #region Generic methods for object operation

        private List<T> GetObjectList<T>(ItemsChoiceType19 type) where T : class
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
        private int SizeOfArray(ItemsChoiceType19 type)
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
        private T GetObjectArray<T>(int p, ItemsChoiceType19 type) where T : class
        {
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                if (pos < 0 || pos >= this.itemsField.Count)
                    return null;
                return itemsField[pos] as T;
            }
        }
        private T InsertNewObject<T>(ItemsChoiceType19 type, int p) where T : class, new()
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
        private T AddNewObject<T>(ItemsChoiceType19 type) where T : class, new()
        {
            T t = new T();
            lock (this)
            {
                this.itemsElementNameField.Add(type);
                this.itemsField.Add(t);
            }
            return t;
        }
        private void SetObject<T>(ItemsChoiceType19 type, int p, T obj) where T : class
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
        private int GetObjectIndex(ItemsChoiceType19 type, int p)
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
        private void RemoveObject(ItemsChoiceType19 type, int p)
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
    public enum ItemsChoiceType19
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
    public class CT_SdtContentRow
    {

        private ArrayList itemsField;

        private List<ItemsChoiceType22> itemsElementNameField;

        public CT_SdtContentRow()
        {
            this.itemsElementNameField = new List<ItemsChoiceType22>();
            this.itemsField = new ArrayList();
        }
        public static CT_SdtContentRow Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SdtContentRow ctObj = new CT_SdtContentRow();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "oMath")
                {
                    ctObj.Items.Add(CT_OMath.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.oMath);
                }
                else if (childNode.LocalName == "oMathPara")
                {
                    ctObj.Items.Add(CT_OMathPara.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.oMathPara);
                }
                else if (childNode.LocalName == "bookmarkEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.bookmarkEnd);
                }
                else if (childNode.LocalName == "bookmarkStart")
                {
                    ctObj.Items.Add(CT_Bookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.bookmarkStart);
                }
                else if (childNode.LocalName == "commentRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.commentRangeEnd);
                }
                else if (childNode.LocalName == "commentRangeStart")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.commentRangeStart);
                }
                else if (childNode.LocalName == "customXml")
                {
                    ctObj.Items.Add(CT_CustomXmlRow.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.customXml);
                }
                else if (childNode.LocalName == "customXmlDelRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.customXmlDelRangeEnd);
                }
                else if (childNode.LocalName == "customXmlDelRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.customXmlDelRangeStart);
                }
                else if (childNode.LocalName == "customXmlInsRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.customXmlInsRangeEnd);
                }
                else if (childNode.LocalName == "customXmlInsRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.customXmlInsRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.customXmlMoveFromRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.customXmlMoveFromRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.customXmlMoveToRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.customXmlMoveToRangeStart);
                }
                else if (childNode.LocalName == "del")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.del);
                }
                else if (childNode.LocalName == "ins")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.ins);
                }
                else if (childNode.LocalName == "moveFrom")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.moveFrom);
                }
                else if (childNode.LocalName == "moveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.moveFromRangeEnd);
                }
                else if (childNode.LocalName == "moveFromRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.moveFromRangeStart);
                }
                else if (childNode.LocalName == "moveTo")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.moveTo);
                }
                else if (childNode.LocalName == "moveToRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.moveToRangeEnd);
                }
                else if (childNode.LocalName == "moveToRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.moveToRangeStart);
                }
                else if (childNode.LocalName == "permEnd")
                {
                    ctObj.Items.Add(CT_Perm.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.permEnd);
                }
                else if (childNode.LocalName == "permStart")
                {
                    ctObj.Items.Add(CT_PermStart.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.permStart);
                }
                else if (childNode.LocalName == "proofErr")
                {
                    ctObj.Items.Add(CT_ProofErr.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.proofErr);
                }
                else if (childNode.LocalName == "sdt")
                {
                    ctObj.Items.Add(CT_SdtRow.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.sdt);
                }
                else if (childNode.LocalName == "tr")
                {
                    ctObj.Items.Add(CT_Row.Parse(childNode, namespaceManager, ctObj));
                    ctObj.ItemsElementName.Add(ItemsChoiceType22.tr);
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
                if (o is CT_OMath)
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
                else if (o is CT_CustomXmlRow)
                    ((CT_CustomXmlRow)o).Write(sw, "customXml");
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
                else if (o is CT_SdtRow)
                    ((CT_SdtRow)o).Write(sw, "sdt");
                else if (o is CT_Row)
                    ((CT_Row)o).Write(sw, "tr");
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement("oMath", typeof(CT_OMath), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("oMathPara", typeof(CT_OMathPara), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("bookmarkEnd", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("bookmarkStart", typeof(CT_Bookmark), Order = 0)]
        [XmlElement("commentRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("commentRangeStart", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("customXml", typeof(CT_CustomXmlRow), Order = 0)]
        [XmlElement("customXmlDelRangeEnd", typeof(CT_Markup), Order = 0)]
        [XmlElement("customXmlDelRangeStart", typeof(CT_TrackChange), Order = 0)]
        [XmlElement("customXmlInsRangeEnd", typeof(CT_Markup), Order = 0)]
        [XmlElement("customXmlInsRangeStart", typeof(CT_TrackChange), Order = 0)]
        [XmlElement("customXmlMoveFromRangeEnd", typeof(CT_Markup), Order = 0)]
        [XmlElement("customXmlMoveFromRangeStart", typeof(CT_TrackChange), Order = 0)]
        [XmlElement("customXmlMoveToRangeEnd", typeof(CT_Markup), Order = 0)]
        [XmlElement("customXmlMoveToRangeStart", typeof(CT_TrackChange), Order = 0)]
        [XmlElement("del", typeof(CT_RunTrackChange), Order = 0)]
        [XmlElement("ins", typeof(CT_RunTrackChange), Order = 0)]
        [XmlElement("moveFrom", typeof(CT_RunTrackChange), Order = 0)]
        [XmlElement("moveFromRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("moveFromRangeStart", typeof(CT_MoveBookmark), Order = 0)]
        [XmlElement("moveTo", typeof(CT_RunTrackChange), Order = 0)]
        [XmlElement("moveToRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("moveToRangeStart", typeof(CT_MoveBookmark), Order = 0)]
        [XmlElement("permEnd", typeof(CT_Perm), Order = 0)]
        [XmlElement("permStart", typeof(CT_PermStart), Order = 0)]
        [XmlElement("proofErr", typeof(CT_ProofErr), Order = 0)]
        [XmlElement("sdt", typeof(CT_SdtRow), Order = 0)]
        [XmlElement("tr", typeof(CT_Row), Order = 0)]
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
        public List<ItemsChoiceType22> ItemsElementName
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
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SdtPr
    {

        private ArrayList itemsField;

        private List<SdtPrElementType> itemsElementNameField;

        public CT_SdtPr()
        {
            this.itemsElementNameField = new List<SdtPrElementType>();
            this.itemsField = new ArrayList();
        }

        [XmlElement("alias", typeof(CT_String), Order = 0)]
        [XmlElement("bibliography", typeof(CT_Empty), Order = 0)]
        [XmlElement("citation", typeof(CT_Empty), Order = 0)]
        [XmlElement("comboBox", typeof(CT_SdtComboBox), Order = 0)]
        [XmlElement("dataBinding", typeof(CT_DataBinding), Order = 0)]
        [XmlElement("date", typeof(CT_SdtDate), Order = 0)]
        [XmlElement("docPartList", typeof(CT_SdtDocPart), Order = 0)]
        [XmlElement("docPartObj", typeof(CT_SdtDocPart), Order = 0)]
        [XmlElement("dropDownList", typeof(CT_SdtDropDownList), Order = 0)]
        [XmlElement("equation", typeof(CT_Empty), Order = 0)]
        [XmlElement("group", typeof(CT_Empty), Order = 0)]
        [XmlElement("id", typeof(CT_DecimalNumber), Order = 0)]
        [XmlElement("lock", typeof(CT_Lock), Order = 0)]
        [XmlElement("picture", typeof(CT_Empty), Order = 0)]
        [XmlElement("placeholder", typeof(CT_Placeholder), Order = 0)]
        [XmlElement("rPr", typeof(CT_RPr), Order = 0)]
        [XmlElement("richText", typeof(CT_Empty), Order = 0)]
        [XmlElement("showingPlcHdr", typeof(CT_OnOff), Order = 0)]
        [XmlElement("tag", typeof(CT_String), Order = 0)]
        [XmlElement("temporary", typeof(CT_OnOff), Order = 0)]
        [XmlElement("text", typeof(CT_SdtText), Order = 0)]
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
        public List<SdtPrElementType> ItemsElementName
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
        public static CT_SdtPr Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SdtPr ctObj = new CT_SdtPr();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "richText")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(SdtPrElementType.richText);
                }
                else if (childNode.LocalName == "docPartList")
                {
                    ctObj.Items.Add(CT_SdtDocPart.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(SdtPrElementType.docPartList);
                }
                else if (childNode.LocalName == "docPartObj")
                {
                    ctObj.Items.Add(CT_SdtDocPart.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(SdtPrElementType.docPartObj);
                }
                else if (childNode.LocalName == "dropDownList")
                {
                    ctObj.Items.Add(CT_SdtDropDownList.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(SdtPrElementType.dropDownList);
                }
                else if (childNode.LocalName == "equation")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(SdtPrElementType.equation);
                }
                else if (childNode.LocalName == "group")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(SdtPrElementType.group);
                }
                else if (childNode.LocalName == "id")
                {
                    ctObj.Items.Add(CT_DecimalNumber.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(SdtPrElementType.id);
                }
                else if (childNode.LocalName == "lock")
                {
                    ctObj.Items.Add(CT_Lock.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(SdtPrElementType.@lock);
                }
                else if (childNode.LocalName == "date")
                {
                    ctObj.Items.Add(CT_SdtDate.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(SdtPrElementType.date);
                }
                else if (childNode.LocalName == "placeholder")
                {
                    ctObj.Items.Add(CT_Placeholder.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(SdtPrElementType.placeholder);
                }
                else if (childNode.LocalName == "rPr")
                {
                    ctObj.Items.Add(CT_RPr.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(SdtPrElementType.rPr);
                }
                else if (childNode.LocalName == "showingPlcHdr")
                {
                    ctObj.Items.Add(CT_OnOff.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(SdtPrElementType.showingPlcHdr);
                }
                else if (childNode.LocalName == "tag")
                {
                    ctObj.Items.Add(CT_String.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(SdtPrElementType.tag);
                }
                else if (childNode.LocalName == "temporary")
                {
                    ctObj.Items.Add(CT_OnOff.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(SdtPrElementType.temporary);
                }
                else if (childNode.LocalName == "text")
                {
                    ctObj.Items.Add(CT_SdtText.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(SdtPrElementType.text);
                }
                else if (childNode.LocalName == "picture")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(SdtPrElementType.picture);
                }
                else if (childNode.LocalName == "alias")
                {
                    ctObj.Items.Add(CT_String.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(SdtPrElementType.alias);
                }
                else if (childNode.LocalName == "bibliography")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(SdtPrElementType.bibliography);
                }
                else if (childNode.LocalName == "citation")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(SdtPrElementType.citation);
                }
                else if (childNode.LocalName == "comboBox")
                {
                    ctObj.Items.Add(CT_SdtComboBox.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(SdtPrElementType.comboBox);
                }
                else if (childNode.LocalName == "dataBinding")
                {
                    ctObj.Items.Add(CT_DataBinding.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(SdtPrElementType.dataBinding);
                }
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            
            for (int i=0;i<this.Items.Count;i++)
            {
                object o = Items[i];
                SdtPrElementType t= this.itemsElementNameField[i];
                if (o is CT_Empty && t== SdtPrElementType.richText)
                    sw.Write("<richText/>");
                else if (o is CT_SdtDocPart&& t  == SdtPrElementType.docPartList)
                    ((CT_SdtDocPart)o).Write(sw, "docPartList");
                else if (o is CT_SdtDocPart&& t == SdtPrElementType.docPartObj)
                    ((CT_SdtDocPart)o).Write(sw, "docPartObj");
                else if (o is CT_SdtDropDownList)
                    ((CT_SdtDropDownList)o).Write(sw, "dropDownList");
                else if (o is CT_Empty&& t== SdtPrElementType.equation)
                    sw.Write("<equation/>");
                else if (o is CT_Empty&& t== SdtPrElementType.group)
                    sw.Write("<group/>");
                else if (o is CT_DecimalNumber&& t== SdtPrElementType.id)
                    ((CT_DecimalNumber)o).Write(sw, "id");
                else if (o is CT_Lock)
                    ((CT_Lock)o).Write(sw, "lock");
                else if (o is CT_SdtDate)
                    ((CT_SdtDate)o).Write(sw, "date");
                else if (o is CT_Placeholder)
                    ((CT_Placeholder)o).Write(sw, "placeholder");
                else if (o is CT_RPr)
                    ((CT_RPr)o).Write(sw, "rPr");
                else if (o is CT_OnOff && t== SdtPrElementType.showingPlcHdr)
                    ((CT_OnOff)o).Write(sw, "showingPlcHdr");
                else if (o is CT_String&& t== SdtPrElementType.tag)
                    ((CT_String)o).Write(sw, "tag");
                else if (o is CT_OnOff && t== SdtPrElementType.temporary)
                    ((CT_OnOff)o).Write(sw, "temporary");
                else if (o is CT_SdtText)
                    ((CT_SdtText)o).Write(sw, "text");
                else if (o is CT_Empty && t== SdtPrElementType.picture)
                    sw.Write("<picture/>");
                else if (o is CT_String&& t== SdtPrElementType.alias)
                    ((CT_String)o).Write(sw, "alias");
                else if (o is CT_Empty && t== SdtPrElementType.bibliography)
                    sw.Write("<bibliography/>");
                else if (o is CT_Empty && t== SdtPrElementType.citation)
                    sw.Write("<citation/>");
                else if (o is CT_SdtComboBox)
                    ((CT_SdtComboBox)o).Write(sw, "comboBox");
                else if (o is CT_DataBinding)
                    ((CT_DataBinding)o).Write(sw, "dataBinding");
                
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        public CT_DecimalNumber AddNewId()
        {
            return AddNewObject<CT_DecimalNumber>(SdtPrElementType.id);
        }

        public CT_SdtDocPart AddNewDocPartObj()
        {
            return AddNewObject<CT_SdtDocPart>(SdtPrElementType.docPartObj);
        }

        public CT_String[] GetAliasArray()
        {
            return GetObjectList<CT_String>(SdtPrElementType.alias).ToArray();
        }

        #region Generic methods for object operation

        public List<T> GetObjectList<T>(SdtPrElementType type) where T : class
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
        private int SizeOfArray(SdtPrElementType type)
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
        private T GetObjectArray<T>(int p, SdtPrElementType type) where T : class
        {
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                if (pos < 0 || pos >= this.itemsField.Count)
                    return null;
                return itemsField[pos] as T;
            }
        }
        private T InsertNewObject<T>(SdtPrElementType type, int p) where T : class, new()
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
        private T AddNewObject<T>(SdtPrElementType type) where T : class, new()
        {
            T t = new T();
            lock (this)
            {
                this.itemsElementNameField.Add(type);
                this.itemsField.Add(t);
            }
            return t;
        }
        private void SetObject<T>(SdtPrElementType type, int p, T obj) where T : class
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
        private int GetObjectIndex(SdtPrElementType type, int p)
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
        private void RemoveObject(SdtPrElementType type, int p)
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
    public enum SdtPrElementType
    {

    
        alias,

    
        bibliography,

    
        citation,

    
        comboBox,

    
        dataBinding,

    
        date,

    
        docPartList,

    
        docPartObj,

    
        dropDownList,

    
        equation,

    
        group,

    
        id,

    
        @lock,

    
        picture,

    
        placeholder,

    
        rPr,

    
        richText,

    
        showingPlcHdr,

    
        tag,

    
        temporary,

    
        text,
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SdtEndPr
    {

        private List<CT_RPr> itemsField;

        public CT_SdtEndPr()
        {
            this.itemsField = new List<CT_RPr>();
        }

        [XmlElement("rPr", Order = 0)]
        public List<CT_RPr> Items
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

        public CT_RPr AddNewRPr()
        {
            CT_RPr r = new CT_RPr();
            this.itemsField.Add(r);
            return r;
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SdtContentRun
    {

        private ArrayList itemsField;

        private List<ItemsChoiceType18> itemsElementNameField;

        public CT_SdtContentRun()
        {
            this.itemsElementNameField = new List<ItemsChoiceType18>();
            this.itemsField = new ArrayList();
        }
        public static CT_SdtContentRun Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SdtContentRun ctObj = new CT_SdtContentRun();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "permStart")
                {
                    ctObj.Items.Add(CT_PermStart.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.permStart);
                }
                else if (childNode.LocalName == "moveToRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.moveToRangeEnd);
                }
                else if (childNode.LocalName == "moveToRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.moveToRangeStart);
                }
                else if (childNode.LocalName == "commentRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.commentRangeEnd);
                }
                else if (childNode.LocalName == "commentRangeStart")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.commentRangeStart);
                }
                else if (childNode.LocalName == "customXml")
                {
                    ctObj.Items.Add(CT_CustomXmlRun.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.customXml);
                }
                else if (childNode.LocalName == "customXmlDelRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.customXmlDelRangeEnd);
                }
                else if (childNode.LocalName == "customXmlDelRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.customXmlDelRangeStart);
                }
                else if (childNode.LocalName == "customXmlInsRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.customXmlInsRangeEnd);
                }
                else if (childNode.LocalName == "customXmlInsRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.customXmlInsRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.customXmlMoveFromRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.customXmlMoveFromRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.customXmlMoveToRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.customXmlMoveToRangeStart);
                }
                else if (childNode.LocalName == "del")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.del);
                }
                else if (childNode.LocalName == "fldSimple")
                {
                    ctObj.Items.Add(CT_SimpleField.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.fldSimple);
                }
                else if (childNode.LocalName == "hyperlink")
                {
                    ctObj.Items.Add(CT_Hyperlink1.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.hyperlink);
                }
                else if (childNode.LocalName == "ins")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.ins);
                }
                else if (childNode.LocalName == "moveFrom")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.moveFrom);
                }
                else if (childNode.LocalName == "bookmarkStart")
                {
                    ctObj.Items.Add(CT_Bookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.bookmarkStart);
                }
                else if (childNode.LocalName == "moveFromRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.moveFromRangeStart);
                }
                else if (childNode.LocalName == "moveTo")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.moveTo);
                }
                else if (childNode.LocalName == "permEnd")
                {
                    ctObj.Items.Add(CT_Perm.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.permEnd);
                }
                else if (childNode.LocalName == "oMath")
                {
                    ctObj.Items.Add(CT_OMath.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.oMath);
                }
                else if (childNode.LocalName == "proofErr")
                {
                    ctObj.Items.Add(CT_ProofErr.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.proofErr);
                }
                else if (childNode.LocalName == "r")
                {
                    ctObj.Items.Add(CT_R.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.r);
                }
                else if (childNode.LocalName == "sdt")
                {
                    ctObj.Items.Add(CT_SdtRun.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.sdt);
                }
                else if (childNode.LocalName == "smartTag")
                {
                    ctObj.Items.Add(CT_SmartTagRun.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.smartTag);
                }
                else if (childNode.LocalName == "subDoc")
                {
                    ctObj.Items.Add(CT_Rel.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.subDoc);
                }
                else if (childNode.LocalName == "moveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.moveFromRangeEnd);
                }
                else if (childNode.LocalName == "oMathPara")
                {
                    ctObj.Items.Add(CT_OMathPara.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.oMathPara);
                }
                else if (childNode.LocalName == "bookmarkEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType18.bookmarkEnd);
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
                if (o is CT_PermStart)
                    ((CT_PermStart)o).Write(sw, "permStart");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "moveToRangeEnd");
                else if (o is CT_MoveBookmark)
                    ((CT_MoveBookmark)o).Write(sw, "moveToRangeStart");
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
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveToRangeStart");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "del");
                else if (o is CT_SimpleField)
                    ((CT_SimpleField)o).Write(sw, "fldSimple");
                else if (o is CT_Hyperlink1)
                    ((CT_Hyperlink1)o).Write(sw, "hyperlink");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "ins");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "moveFrom");
                else if (o is CT_Bookmark)
                    ((CT_Bookmark)o).Write(sw, "bookmarkStart");
                else if (o is CT_MoveBookmark)
                    ((CT_MoveBookmark)o).Write(sw, "moveFromRangeStart");
                else if (o is CT_RunTrackChange)
                    ((CT_RunTrackChange)o).Write(sw, "moveTo");
                else if (o is CT_Perm)
                    ((CT_Perm)o).Write(sw, "permEnd");
                else if (o is CT_OMath)
                    ((CT_OMath)o).Write(sw, "oMath");
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
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "moveFromRangeEnd");
                else if (o is CT_OMathPara)
                    ((CT_OMathPara)o).Write(sw, "oMathPara");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "bookmarkEnd");
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement("oMath", typeof(CT_OMath), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("oMathPara", typeof(CT_OMathPara), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("bookmarkEnd", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("bookmarkStart", typeof(CT_Bookmark), Order = 0)]
        [XmlElement("commentRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("commentRangeStart", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("customXml", typeof(CT_CustomXmlRun), Order = 0)]
        [XmlElement("customXmlDelRangeEnd", typeof(CT_Markup), Order = 0)]
        [XmlElement("customXmlDelRangeStart", typeof(CT_TrackChange), Order = 0)]
        [XmlElement("customXmlInsRangeEnd", typeof(CT_Markup), Order = 0)]
        [XmlElement("customXmlInsRangeStart", typeof(CT_TrackChange), Order = 0)]
        [XmlElement("customXmlMoveFromRangeEnd", typeof(CT_Markup), Order = 0)]
        [XmlElement("customXmlMoveFromRangeStart", typeof(CT_TrackChange), Order = 0)]
        [XmlElement("customXmlMoveToRangeEnd", typeof(CT_Markup), Order = 0)]
        [XmlElement("customXmlMoveToRangeStart", typeof(CT_TrackChange), Order = 0)]
        [XmlElement("del", typeof(CT_RunTrackChange), Order = 0)]
        [XmlElement("fldSimple", typeof(CT_SimpleField), Order = 0)]
        [XmlElement("hyperlink", typeof(CT_Hyperlink1), Order = 0)]
        [XmlElement("ins", typeof(CT_RunTrackChange), Order = 0)]
        [XmlElement("moveFrom", typeof(CT_RunTrackChange), Order = 0)]
        [XmlElement("moveFromRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("moveFromRangeStart", typeof(CT_MoveBookmark), Order = 0)]
        [XmlElement("moveTo", typeof(CT_RunTrackChange), Order = 0)]
        [XmlElement("moveToRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [XmlElement("moveToRangeStart", typeof(CT_MoveBookmark), Order = 0)]
        [XmlElement("permEnd", typeof(CT_Perm), Order = 0)]
        [XmlElement("permStart", typeof(CT_PermStart), Order = 0)]
        [XmlElement("proofErr", typeof(CT_ProofErr), Order = 0)]
        [XmlElement("r", typeof(CT_R), Order = 0)]
        [XmlElement("sdt", typeof(CT_SdtRun), Order = 0)]
        [XmlElement("smartTag", typeof(CT_SmartTagRun), Order = 0)]
        [XmlElement("subDoc", typeof(CT_Rel), Order = 0)]
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
        public List<ItemsChoiceType18> ItemsElementName
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

        public IEnumerable<CT_R> GetRList()
        {
            return GetObjectList<CT_R>(ItemsChoiceType18.r);
        }
        #region Generic methods for object operation

        private List<T> GetObjectList<T>(ItemsChoiceType18 type) where T : class
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
        private int SizeOfArray(ItemsChoiceType18 type)
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
        private T GetObjectArray<T>(int p, ItemsChoiceType18 type) where T : class
        {
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                if (pos < 0 || pos >= this.itemsField.Count)
                    return null;
                return itemsField[pos] as T;
            }
        }
        private T InsertNewObject<T>(ItemsChoiceType18 type, int p) where T : class, new()
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
        private T AddNewObject<T>(ItemsChoiceType18 type) where T : class, new()
        {
            T t = new T();
            lock (this)
            {
                this.itemsElementNameField.Add(type);
                this.itemsField.Add(t);
            }
            return t;
        }
        private void SetObject<T>(ItemsChoiceType18 type, int p, T obj) where T : class
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
        private int GetObjectIndex(ItemsChoiceType18 type, int p)
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
        private void RemoveObject(ItemsChoiceType18 type, int p)
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
    public enum ItemsChoiceType18
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
    public class CT_SdtListItem
    {

        private string displayTextField;

        private string valueField;
        public static CT_SdtListItem Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SdtListItem ctObj = new CT_SdtListItem();
            ctObj.displayText = XmlHelper.ReadString(node.Attributes["w:displayText"]);
            ctObj.value = XmlHelper.ReadString(node.Attributes["w:value"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:displayText", this.displayText);
            XmlHelper.WriteAttribute(sw, "w:value", this.value);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string displayText
        {
            get
            {
                return this.displayTextField;
            }
            set
            {
                this.displayTextField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string value
        {
            get
            {
                return this.valueField;
            }
            set
            {
                this.valueField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SdtDateMappingType
    {

        private ST_SdtDateMappingType valField;

        private bool valFieldSpecified;
        public static CT_SdtDateMappingType Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SdtDateMappingType ctObj = new CT_SdtDateMappingType();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_SdtDateMappingType)Enum.Parse(typeof(ST_SdtDateMappingType), node.Attributes["w:val"].Value);
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
        public ST_SdtDateMappingType val
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
    public enum ST_SdtDateMappingType
    {

    
        text,

    
        date,

    
        dateTime,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_CalendarType
    {

        private ST_CalendarType valField;

        private bool valFieldSpecified;

        public static CT_CalendarType Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CalendarType ctObj = new CT_CalendarType();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_CalendarType)Enum.Parse(typeof(ST_CalendarType), node.Attributes["w:val"].Value);
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
        public ST_CalendarType val
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
    public enum ST_CalendarType
    {

    
        gregorian,

    
        hijri,

    
        hebrew,

    
        taiwan,

    
        japan,

    
        thai,

    
        korea,

    
        saka,

    
        gregorianXlitEnglish,

    
        gregorianXlitFrench,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SdtDate
    {

        private CT_String dateFormatField;

        private CT_Lang lidField;

        private CT_SdtDateMappingType storeMappedDataAsField;

        private CT_CalendarType calendarField;

        private string fullDateField;

        private bool fullDateFieldSpecified;

        public CT_SdtDate()
        {
            //this.calendarField = new CT_CalendarType();
            //this.storeMappedDataAsField = new CT_SdtDateMappingType();
            //this.lidField = new CT_Lang();
            //this.dateFormatField = new CT_String();
        }
        public static CT_SdtDate Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SdtDate ctObj = new CT_SdtDate();
            ctObj.fullDateField = XmlHelper.ReadString(node.Attributes["w.fullDate"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "dateFormat")
                    ctObj.dateFormat = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lid")
                    ctObj.lid = CT_Lang.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "storeMappedDataAs")
                    ctObj.storeMappedDataAs = CT_SdtDateMappingType.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "calendar")
                    ctObj.calendar = CT_CalendarType.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:fullDate", this.fullDateField);
            sw.Write(">");
            if (this.dateFormat != null)
                this.dateFormat.Write(sw, "dateFormat");
            if (this.lid != null)
                this.lid.Write(sw, "lid");
            if (this.storeMappedDataAs != null)
                this.storeMappedDataAs.Write(sw, "storeMappedDataAs");
            if (this.calendar != null)
                this.calendar.Write(sw, "calendar");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_String dateFormat
        {
            get
            {
                return this.dateFormatField;
            }
            set
            {
                this.dateFormatField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_Lang lid
        {
            get
            {
                return this.lidField;
            }
            set
            {
                this.lidField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_SdtDateMappingType storeMappedDataAs
        {
            get
            {
                return this.storeMappedDataAsField;
            }
            set
            {
                this.storeMappedDataAsField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_CalendarType calendar
        {
            get
            {
                return this.calendarField;
            }
            set
            {
                this.calendarField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string fullDate
        {
            get
            {
                return this.fullDateField;
            }
            set
            {
                this.fullDateField = value;
            }
        }

        [XmlIgnore]
        public bool fullDateSpecified
        {
            get
            {
                return this.fullDateFieldSpecified;
            }
            set
            {
                this.fullDateFieldSpecified = value;
            }
        }
    }
    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SdtRow
    {

        private CT_SdtPr sdtPrField;

        private List<CT_RPr> sdtEndPrField;

        private CT_SdtContentRow sdtContentField;

        public CT_SdtRow()
        {
            //this.sdtContentField = new CT_SdtContentRow();
            //this.sdtEndPrField = new List<CT_RPr>();
            //this.sdtPrField = new CT_SdtPr();
        }

        public static CT_SdtRow Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SdtRow ctObj = new CT_SdtRow();
            ctObj.sdtEndPr = new List<CT_RPr>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "sdtPr")
                    ctObj.sdtPr = CT_SdtPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sdtContent")
                    ctObj.sdtContent = CT_SdtContentRow.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sdtEndPr")
                    ctObj.sdtEndPr.Add(CT_RPr.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.sdtPr != null)
                this.sdtPr.Write(sw, "sdtPr");
            if (this.sdtContent != null)
                this.sdtContent.Write(sw, "sdtContent");
            if (this.sdtEndPr != null)
            {
                foreach (CT_RPr x in this.sdtEndPr)
                {
                    x.Write(sw, "sdtEndPr");
                }
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_SdtPr sdtPr
        {
            get
            {
                return this.sdtPrField;
            }
            set
            {
                this.sdtPrField = value;
            }
        }

        [XmlArray(Order = 1)]
        [XmlArrayItem("rPr", IsNullable = false)]
        public List<CT_RPr> sdtEndPr
        {
            get
            {
                return this.sdtEndPrField;
            }
            set
            {
                this.sdtEndPrField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_SdtContentRow sdtContent
        {
            get
            {
                return this.sdtContentField;
            }
            set
            {
                this.sdtContentField = value;
            }
        }
    }
    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SdtText
    {

        private ST_OnOff multiLineField;

        private bool multiLineFieldSpecified;
        public static CT_SdtText Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SdtText ctObj = new CT_SdtText();
            if (node.Attributes["w:multiLine"] != null)
                ctObj.multiLine = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:multiLine"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:multiLine", this.multiLine.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff multiLine
        {
            get
            {
                return this.multiLineField;
            }
            set
            {
                this.multiLineField = value;
            }
        }

        [XmlIgnore]
        public bool multiLineSpecified
        {
            get
            {
                return this.multiLineFieldSpecified;
            }
            set
            {
                this.multiLineFieldSpecified = value;
            }
        }
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Lock
    {
        public static CT_Lock Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Lock ctObj = new CT_Lock();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_Lock)Enum.Parse(typeof(ST_Lock), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        private ST_Lock valField;

        private bool valFieldSpecified;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_Lock val
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
    public enum ST_Lock
    {

    
        sdtLocked,

    
        contentLocked,

    
        unlocked,

    
        sdtContentLocked,
    }
}