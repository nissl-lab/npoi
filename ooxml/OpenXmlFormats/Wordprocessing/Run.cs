using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Serialization;
using NPOI.OpenXmlFormats.Shared;
using System.IO;
using NPOI.OpenXml4Net.Util;

namespace NPOI.OpenXmlFormats.Wordprocessing
{


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_R
    {

        private CT_RPr rPrField;

        private List<object> itemsField;

        private List<RunItemsChoiceType> itemsElementNameField;

        private byte[] rsidRPrField;

        private byte[] rsidDelField;

        private byte[] rsidRField;

        public CT_R()
        {
            this.itemsElementNameField = new List<RunItemsChoiceType>();
            this.itemsField = new List<object>();
            //this.rPrField = new CT_RPr();
        }

        [XmlElement(Order = 0, IsNullable = true)]
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

        [XmlElement("annotationRef", typeof(CT_Empty), Order = 1)]
        [XmlElement("br", typeof(CT_Br), Order = 1)]
        [XmlElement("commentReference", typeof(CT_Markup), Order = 1)]
        [XmlElement("continuationSeparator", typeof(CT_Empty), Order = 1)]
        [XmlElement("cr", typeof(CT_Empty), Order = 1)]
        [XmlElement("dayLong", typeof(CT_Empty), Order = 1)]
        [XmlElement("dayShort", typeof(CT_Empty), Order = 1)]
        [XmlElement("delInstrText", typeof(CT_Text), Order = 1)]
        [XmlElement("delText", typeof(CT_Text), Order = 1)]
        [XmlElement("drawing", typeof(CT_Drawing), Order = 1)]
        [XmlElement("endnoteRef", typeof(CT_Empty), Order = 1)]
        [XmlElement("endnoteReference", typeof(CT_FtnEdnRef), Order = 1)]
        [XmlElement("fldChar", typeof(CT_FldChar), Order = 1)]
        [XmlElement("footnoteRef", typeof(CT_Empty), Order = 1)]
        [XmlElement("footnoteReference", typeof(CT_FtnEdnRef), Order = 1)]
        [XmlElement("instrText", typeof(CT_Text), Order = 1)]
        [XmlElement("lastRenderedPageBreak", typeof(CT_Empty), Order = 1)]
        [XmlElement("monthLong", typeof(CT_Empty), Order = 1)]
        [XmlElement("monthShort", typeof(CT_Empty), Order = 1)]
        [XmlElement("noBreakHyphen", typeof(CT_Empty), Order = 1)]
        [XmlElement("object", typeof(CT_Object), Order = 1)]
        [XmlElement("pgNum", typeof(CT_Empty), Order = 1)]
        [XmlElement("pict", typeof(CT_Picture), Order = 1)]
        [XmlElement("ptab", typeof(CT_PTab), Order = 1)]
        [XmlElement("ruby", typeof(CT_Ruby), Order = 1)]
        [XmlElement("separator", typeof(CT_Empty), Order = 1)]
        [XmlElement("softHyphen", typeof(CT_Empty), Order = 1)]
        [XmlElement("sym", typeof(CT_Sym), Order = 1)]
        [XmlElement("t", typeof(CT_Text), Order = 1)]
        [XmlElement("tab", typeof(CT_Empty), Order = 1)]
        [XmlElement("yearLong", typeof(CT_Empty), Order = 1)]
        [XmlElement("yearShort", typeof(CT_Empty), Order = 1)]
        [XmlChoiceIdentifier("ItemsElementName")]
        public object[] Items
        {
            get
            {
                return this.itemsField.ToArray();
            }
            set
            {
                if (value != null && value.Length != 0)
                    this.itemsField = new List<object>(value);
                else
                    this.itemsField = new List<object>();
            }
        }

        [XmlElement("ItemsElementName", Order = 2)]
        [XmlIgnore]
        public RunItemsChoiceType[] ItemsElementName
        {
            get
            {
                return this.itemsElementNameField.ToArray();
            }
            set
            {
                if (value != null && value.Length != 0)
                    this.itemsElementNameField = new List<RunItemsChoiceType>(value);
                else
                    this.itemsElementNameField = new List<RunItemsChoiceType>();
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

        public CT_Text AddNewT()
        {
            return AddNewObject<CT_Text>(RunItemsChoiceType.t);
        }

        public CT_RPr AddNewRPr()
        {
            if (this.rPrField == null)
                this.rPrField = new CT_RPr();
            return this.rPrField;
        }

        public CT_Empty AddNewTab()
        {
            return AddNewObject<CT_Empty>(RunItemsChoiceType.tab);
        }

        public CT_FldChar AddNewFldChar()
        {
            return AddNewObject<CT_FldChar>(RunItemsChoiceType.fldChar);
        }

        public CT_Text AddNewInstrText()
        {
            return AddNewObject<CT_Text>(RunItemsChoiceType.instrText);
        }

        public CT_Empty AddNewCr()
        {
            return AddNewObject<CT_Empty>(RunItemsChoiceType.cr);
        }

        public CT_Br AddNewBr()
        {
            return AddNewObject<CT_Br>(RunItemsChoiceType.br);
        }

        public bool IsSetRPr()
        {
            return this.rPrField != null;
        }

        public int SizeOfTArray()
        {
            return SizeOfArray(RunItemsChoiceType.t);
        }

        public CT_Text GetTArray(int pos)
        {
            return GetObjectArray<CT_Text>(pos, RunItemsChoiceType.t);
        }

        public List<CT_Text> GetTList()
        {
            return GetObjectList<CT_Text>(RunItemsChoiceType.t);
        }

        public CT_Drawing AddNewDrawing()
        {
            return AddNewObject<CT_Drawing>(RunItemsChoiceType.drawing);
        }

        public IList<CT_Drawing> GetDrawingList()
        {
            return GetObjectList<CT_Drawing>(RunItemsChoiceType.drawing);
        }

        public IList<CT_Picture> GetPictList()
        {
            return GetObjectList<CT_Picture>(RunItemsChoiceType.pict);
        }

        public CT_Picture AddNewPict()
        {
            return AddNewObject<CT_Picture>(RunItemsChoiceType.pict);
        }
        #region Generic methods for object operation

        private List<T> GetObjectList<T>(RunItemsChoiceType type) where T : class
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
        private int SizeOfArray(RunItemsChoiceType type)
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
        private T GetObjectArray<T>(int p, RunItemsChoiceType type) where T : class
        {
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                if (pos < 0 || pos >= this.itemsField.Count)
                    return null;
                return itemsField[pos] as T;
            }
        }
        private T InsertNewObject<T>(RunItemsChoiceType type, int p) where T : class, new()
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
        private T AddNewObject<T>(RunItemsChoiceType type) where T : class, new()
        {
            T t = new T();
            lock (this)
            {
                this.itemsElementNameField.Add(type);
                this.itemsField.Add(t);
            }
            return t;
        }
        private void SetObject<T>(RunItemsChoiceType type, int p, T obj) where T : class
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
        private int GetObjectIndex(RunItemsChoiceType type, int p)
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
        private void RemoveObject(RunItemsChoiceType type, int p)
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

        public CT_Drawing GetDrawingArray(int p)
        {
            return GetObjectArray<CT_Drawing>(p, RunItemsChoiceType.drawing);
        }

        public int SizeOfCrArray()
        {
            return SizeOfArray(RunItemsChoiceType.cr);
        }

        public IList<CT_Empty> GetCrList()
        {
            return GetObjectList<CT_Empty>(RunItemsChoiceType.cr);
        }

        public int SizeOfBrArray()
        {
            return SizeOfArray(RunItemsChoiceType.br);
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum RunItemsChoiceType
    {

    
        annotationRef,

    
        br,

    
        commentReference,

    
        continuationSeparator,

    
        cr,

    
        dayLong,

    
        dayShort,

    
        delInstrText,

    
        delText,

    
        drawing,

    
        endnoteRef,

    
        endnoteReference,

    
        fldChar,

    
        footnoteRef,

    
        footnoteReference,

    
        instrText,

    
        lastRenderedPageBreak,

    
        monthLong,

    
        monthShort,

    
        noBreakHyphen,

    
        @object,

    
        pgNum,

    
        pict,

    
        ptab,

    
        ruby,

    
        separator,

    
        softHyphen,

    
        sym,

    
        t,

    
        tab,

    
        yearLong,

    
        yearShort,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_RubyAlign
    {

        private ST_RubyAlign valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_RubyAlign val
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
    public enum ST_RubyAlign
    {

    
        center,

    
        distributeLetter,

    
        distributeSpace,

    
        left,

    
        right,

    
        rightVertical,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_RubyPr
    {

        private CT_RubyAlign rubyAlignField;

        private CT_HpsMeasure hpsField;

        private CT_HpsMeasure hpsRaiseField;

        private CT_HpsMeasure hpsBaseTextField;

        private CT_Lang lidField;

        private CT_OnOff dirtyField;

        public CT_RubyPr()
        {
            this.dirtyField = new CT_OnOff();
            this.lidField = new CT_Lang();
            this.hpsBaseTextField = new CT_HpsMeasure();
            this.hpsRaiseField = new CT_HpsMeasure();
            this.hpsField = new CT_HpsMeasure();
            this.rubyAlignField = new CT_RubyAlign();
        }

        [XmlElement(Order = 0)]
        public CT_RubyAlign rubyAlign
        {
            get
            {
                return this.rubyAlignField;
            }
            set
            {
                this.rubyAlignField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_HpsMeasure hps
        {
            get
            {
                return this.hpsField;
            }
            set
            {
                this.hpsField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_HpsMeasure hpsRaise
        {
            get
            {
                return this.hpsRaiseField;
            }
            set
            {
                this.hpsRaiseField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_HpsMeasure hpsBaseText
        {
            get
            {
                return this.hpsBaseTextField;
            }
            set
            {
                this.hpsBaseTextField = value;
            }
        }

        [XmlElement(Order = 4)]
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

        [XmlElement(Order = 5)]
        public CT_OnOff dirty
        {
            get
            {
                return this.dirtyField;
            }
            set
            {
                this.dirtyField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_RubyContent
    {

        private object[] itemsField;

        private ItemsChoiceType16[] itemsElementNameField;

        public CT_RubyContent()
        {
            this.itemsElementNameField = new ItemsChoiceType16[0];
            this.itemsField = new object[0];
        }

        [XmlElement("oMath", typeof(CT_OMath), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("oMathPara", typeof(CT_OMathPara), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
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
        [XmlElement("r", typeof(CT_R), Order = 0)]
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

        [XmlElement("ItemsElementName", Order = 1)]
        [XmlIgnore]
        public ItemsChoiceType16[] ItemsElementName
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
    public enum ItemsChoiceType16
    {

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:oMath")]
        oMath,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:oMathPara")]
        oMathPara,

    
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

    
        r,
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Ruby
    {

        private CT_RubyPr rubyPrField;

        private CT_RubyContent rtField;

        private CT_RubyContent rubyBaseField;

        public CT_Ruby()
        {
            this.rubyBaseField = new CT_RubyContent();
            this.rtField = new CT_RubyContent();
            this.rubyPrField = new CT_RubyPr();
        }

        [XmlElement(Order = 0)]
        public CT_RubyPr rubyPr
        {
            get
            {
                return this.rubyPrField;
            }
            set
            {
                this.rubyPrField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_RubyContent rt
        {
            get
            {
                return this.rtField;
            }
            set
            {
                this.rtField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_RubyContent rubyBase
        {
            get
            {
                return this.rubyBaseField;
            }
            set
            {
                this.rubyBaseField = value;
            }
        }
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Jc
    {
        public static CT_Jc Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Jc ctObj = new CT_Jc();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_Jc)Enum.Parse(typeof(ST_Jc), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }


        private ST_Jc valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_Jc val
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
    public enum ST_Jc
    {
    
        left,

    
        center,

    
        right,

    
        both,

    
        mediumKashida,

    
        distribute,

    
        numTab,

    
        highKashida,

    
        lowKashida,

    
        thaiDistribute,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TextDirection
    {
        public static CT_TextDirection Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TextDirection ctObj = new CT_TextDirection();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_TextDirection)Enum.Parse(typeof(ST_TextDirection), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }


        private ST_TextDirection valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_TextDirection val
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
    public enum ST_TextDirection
    {

    
        lrTb,

    
        tbRl,

    
        btLr,

    
        lrTbV,

    
        tbRlV,

    
        tbLrV,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TextAlignment
    {
        public static CT_TextAlignment Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TextAlignment ctObj = new CT_TextAlignment();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_TextAlignment)Enum.Parse(typeof(ST_TextAlignment), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }


        private ST_TextAlignment valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_TextAlignment val
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
    public enum ST_TextAlignment
    {

    
        top,

    
        center,

    
        baseline,

    
        bottom,

    
        auto,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TextboxTightWrap
    {
        public static CT_TextboxTightWrap Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TextboxTightWrap ctObj = new CT_TextboxTightWrap();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_TextboxTightWrap)Enum.Parse(typeof(ST_TextboxTightWrap), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        private ST_TextboxTightWrap valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_TextboxTightWrap val
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
    public enum ST_TextboxTightWrap
    {

    
        none,

    
        allLines,

    
        firstAndLastLine,

    
        firstLineOnly,

    
        lastLineOnly,
    }



    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_VerticalAlignRun
    {
        public static CT_VerticalAlignRun Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_VerticalAlignRun ctObj = new CT_VerticalAlignRun();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_VerticalAlignRun)Enum.Parse(typeof(ST_VerticalAlignRun), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }


        private ST_VerticalAlignRun valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_VerticalAlignRun val
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
    public enum ST_VerticalAlignRun
    {

    
        baseline,

    
        superscript,

    
        subscript,
    }




    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType5
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
    public class CT_RunTrackChange : CT_TrackChange
    {

        public CT_RunTrackChange()
        {
        }
        List<CT_RunTrackChange> moveToField;
        public List<CT_RunTrackChange> moveTo
        {
            get { return this.moveToField; }
            set { this.moveToField = value; }
        }

        List<CT_TrackChange> customXmlDelRangeStartField;
        public List<CT_TrackChange> customXmlDelRangeStart
        {
            get { return this.customXmlDelRangeStartField; }
            set { this.customXmlDelRangeStartField = value; }
        }

        List<CT_Acc> accField;
        public List<CT_Acc> acc
        {
            get { return this.accField; }
            set { this.accField = value; }
        }

        List<CT_Bar> barField;
        public List<CT_Bar> bar
        {
            get { return this.barField; }
            set { this.barField = value; }
        }

        List<CT_BorderBox> borderBoxField;
        public List<CT_BorderBox> borderBox
        {
            get { return this.borderBoxField; }
            set { this.borderBoxField = value; }
        }

        List<CT_Box> boxField;
        public List<CT_Box> box
        {
            get { return this.boxField; }
            set { this.boxField = value; }
        }

        List<CT_D> dField;
        public List<CT_D> d
        {
            get { return this.dField; }
            set { this.dField = value; }
        }

        List<CT_EqArr> eqArrField;
        public List<CT_EqArr> eqArr
        {
            get { return this.eqArrField; }
            set { this.eqArrField = value; }
        }

        List<CT_F> fField;
        public List<CT_F> f
        {
            get { return this.fField; }
            set { this.fField = value; }
        }

        List<CT_Func> funcField;
        public List<CT_Func> func
        {
            get { return this.funcField; }
            set { this.funcField = value; }
        }

        List<CT_GroupChr> groupChrField;
        public List<CT_GroupChr> groupChr
        {
            get { return this.groupChrField; }
            set { this.groupChrField = value; }
        }

        List<CT_LimLow> limLowField;
        public List<CT_LimLow> limLow
        {
            get { return this.limLowField; }
            set { this.limLowField = value; }
        }

        List<CT_LimUpp> limUppField;
        public List<CT_LimUpp> limUpp
        {
            get { return this.limUppField; }
            set { this.limUppField = value; }
        }

        List<CT_M> mField;
        public List<CT_M> m
        {
            get { return this.mField; }
            set { this.mField = value; }
        }

        List<CT_Nary> naryField;
        public List<CT_Nary> nary
        {
            get { return this.naryField; }
            set { this.naryField = value; }
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

        List<CT_Phant> phantField;
        public List<CT_Phant> phant
        {
            get { return this.phantField; }
            set { this.phantField = value; }
        }

        List<CT_R> rField;
        public List<CT_R> r
        {
            get { return this.rField; }
            set { this.rField = value; }
        }

        List<CT_Rad> radField;
        public List<CT_Rad> rad
        {
            get { return this.radField; }
            set { this.radField = value; }
        }

        List<CT_SPre> sPreField;
        public List<CT_SPre> sPre
        {
            get { return this.sPreField; }
            set { this.sPreField = value; }
        }

        List<CT_SSub> sSubField;
        public List<CT_SSub> sSub
        {
            get { return this.sSubField; }
            set { this.sSubField = value; }
        }

        List<CT_SSubSup> sSubSupField;
        public List<CT_SSubSup> sSubSup
        {
            get { return this.sSubSupField; }
            set { this.sSubSupField = value; }
        }

        List<CT_SSup> sSupField;
        public List<CT_SSup> sSup
        {
            get { return this.sSupField; }
            set { this.sSupField = value; }
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

        List<CT_PermStart> permStartField;
        public List<CT_PermStart> permStart
        {
            get { return this.permStartField; }
            set { this.permStartField = value; }
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

        List<CT_MoveBookmark> moveFromRangeStartField;
        public List<CT_MoveBookmark> moveFromRangeStart
        {
            get { return this.moveFromRangeStartField; }
            set { this.moveFromRangeStartField = value; }
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

        List<CT_MarkupRange> moveFromRangeEndField;
        public List<CT_MarkupRange> moveFromRangeEnd
        {
            get { return this.moveFromRangeEndField; }
            set { this.moveFromRangeEndField = value; }
        }

        List<CT_ProofErr> proofErrField;
        public List<CT_ProofErr> proofErr
        {
            get { return this.proofErrField; }
            set { this.proofErrField = value; }
        }

        List<CT_R> r1Field;
        public List<CT_R> r1
        {
            get { return this.r1Field; }
            set { this.r1Field = value; }
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



        public IEnumerable<CT_R> GetRList()
        {
            return r1Field;
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType6
    {

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:acc")]
        acc,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:bar")]
        bar,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:borderBox")]
        borderBox,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:box")]
        box,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:d")]
        d,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:eqArr")]
        eqArr,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:f")]
        f,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:func")]
        func,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:groupChr")]
        groupChr,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:limLow")]
        limLow,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:limUpp")]
        limUpp,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:m")]
        m,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:nary")]
        nary,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:oMath")]
        oMath,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:oMathPara")]
        oMathPara,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:phant")]
        phant,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:r")]
        r,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:rad")]
        rad,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:sPre")]
        sPre,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:sSub")]
        sSub,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:sSubSup")]
        sSubSup,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/math:sSup")]
        sSup,

    
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

    
        [XmlEnum("r")]
        r1,

    
        sdt,

    
        smartTag,
    }
    
}