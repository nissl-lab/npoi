using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using NPOI.OpenXmlFormats.Shared;

namespace NPOI.OpenXmlFormats.Wordprocessing
{


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
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

        [System.Xml.Serialization.XmlElement(Order = 0)]
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

        [System.Xml.Serialization.XmlElement("annotationRef", typeof(CT_Empty), Order = 1)]
        [System.Xml.Serialization.XmlElement("br", typeof(CT_Br), Order = 1)]
        [System.Xml.Serialization.XmlElement("commentReference", typeof(CT_Markup), Order = 1)]
        [System.Xml.Serialization.XmlElement("continuationSeparator", typeof(CT_Empty), Order = 1)]
        [System.Xml.Serialization.XmlElement("cr", typeof(CT_Empty), Order = 1)]
        [System.Xml.Serialization.XmlElement("dayLong", typeof(CT_Empty), Order = 1)]
        [System.Xml.Serialization.XmlElement("dayShort", typeof(CT_Empty), Order = 1)]
        [System.Xml.Serialization.XmlElement("delInstrText", typeof(CT_Text), Order = 1)]
        [System.Xml.Serialization.XmlElement("delText", typeof(CT_Text), Order = 1)]
        [System.Xml.Serialization.XmlElement("drawing", typeof(CT_Drawing), Order = 1)]
        [System.Xml.Serialization.XmlElement("endnoteRef", typeof(CT_Empty), Order = 1)]
        [System.Xml.Serialization.XmlElement("endnoteReference", typeof(CT_FtnEdnRef), Order = 1)]
        [System.Xml.Serialization.XmlElement("fldChar", typeof(CT_FldChar), Order = 1)]
        [System.Xml.Serialization.XmlElement("footnoteRef", typeof(CT_Empty), Order = 1)]
        [System.Xml.Serialization.XmlElement("footnoteReference", typeof(CT_FtnEdnRef), Order = 1)]
        [System.Xml.Serialization.XmlElement("instrText", typeof(CT_Text), Order = 1)]
        [System.Xml.Serialization.XmlElement("lastRenderedPageBreak", typeof(CT_Empty), Order = 1)]
        [System.Xml.Serialization.XmlElement("monthLong", typeof(CT_Empty), Order = 1)]
        [System.Xml.Serialization.XmlElement("monthShort", typeof(CT_Empty), Order = 1)]
        [System.Xml.Serialization.XmlElement("noBreakHyphen", typeof(CT_Empty), Order = 1)]
        [System.Xml.Serialization.XmlElement("object", typeof(CT_Object), Order = 1)]
        [System.Xml.Serialization.XmlElement("pgNum", typeof(CT_Empty), Order = 1)]
        [System.Xml.Serialization.XmlElement("pict", typeof(CT_Picture), Order = 1)]
        [System.Xml.Serialization.XmlElement("ptab", typeof(CT_PTab), Order = 1)]
        [System.Xml.Serialization.XmlElement("ruby", typeof(CT_Ruby), Order = 1)]
        [System.Xml.Serialization.XmlElement("separator", typeof(CT_Empty), Order = 1)]
        [System.Xml.Serialization.XmlElement("softHyphen", typeof(CT_Empty), Order = 1)]
        [System.Xml.Serialization.XmlElement("sym", typeof(CT_Sym), Order = 1)]
        [System.Xml.Serialization.XmlElement("t", typeof(CT_Text), Order = 1)]
        [System.Xml.Serialization.XmlElement("tab", typeof(CT_Empty), Order = 1)]
        [System.Xml.Serialization.XmlElement("yearLong", typeof(CT_Empty), Order = 1)]
        [System.Xml.Serialization.XmlElement("yearShort", typeof(CT_Empty), Order = 1)]
        [System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemsElementName")]
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

        [System.Xml.Serialization.XmlElement("ItemsElementName", Order = 2)]
        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
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


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
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


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_RubyAlign
    {

        private ST_RubyAlign valField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_RubyAlign
    {

    
        center,

    
        distributeLetter,

    
        distributeSpace,

    
        left,

    
        right,

    
        rightVertical,
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
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

        [System.Xml.Serialization.XmlElement(Order = 0)]
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

        [System.Xml.Serialization.XmlElement(Order = 1)]
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

        [System.Xml.Serialization.XmlElement(Order = 2)]
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

        [System.Xml.Serialization.XmlElement(Order = 3)]
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

        [System.Xml.Serialization.XmlElement(Order = 4)]
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

        [System.Xml.Serialization.XmlElement(Order = 5)]
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


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_RubyContent
    {

        private object[] itemsField;

        private ItemsChoiceType16[] itemsElementNameField;

        public CT_RubyContent()
        {
            this.itemsElementNameField = new ItemsChoiceType16[0];
            this.itemsField = new object[0];
        }

        [System.Xml.Serialization.XmlElement("oMath", typeof(CT_OMath), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("oMathPara", typeof(CT_OMathPara), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("bookmarkEnd", typeof(CT_MarkupRange), Order = 0)]
        [System.Xml.Serialization.XmlElement("bookmarkStart", typeof(CT_Bookmark), Order = 0)]
        [System.Xml.Serialization.XmlElement("commentRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [System.Xml.Serialization.XmlElement("commentRangeStart", typeof(CT_MarkupRange), Order = 0)]
        [System.Xml.Serialization.XmlElement("customXmlDelRangeEnd", typeof(CT_Markup), Order = 0)]
        [System.Xml.Serialization.XmlElement("customXmlDelRangeStart", typeof(CT_TrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElement("customXmlInsRangeEnd", typeof(CT_Markup), Order = 0)]
        [System.Xml.Serialization.XmlElement("customXmlInsRangeStart", typeof(CT_TrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElement("customXmlMoveFromRangeEnd", typeof(CT_Markup), Order = 0)]
        [System.Xml.Serialization.XmlElement("customXmlMoveFromRangeStart", typeof(CT_TrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElement("customXmlMoveToRangeEnd", typeof(CT_Markup), Order = 0)]
        [System.Xml.Serialization.XmlElement("customXmlMoveToRangeStart", typeof(CT_TrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElement("del", typeof(CT_RunTrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElement("ins", typeof(CT_RunTrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElement("moveFrom", typeof(CT_RunTrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElement("moveFromRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [System.Xml.Serialization.XmlElement("moveFromRangeStart", typeof(CT_MoveBookmark), Order = 0)]
        [System.Xml.Serialization.XmlElement("moveTo", typeof(CT_RunTrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElement("moveToRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [System.Xml.Serialization.XmlElement("moveToRangeStart", typeof(CT_MoveBookmark), Order = 0)]
        [System.Xml.Serialization.XmlElement("permEnd", typeof(CT_Perm), Order = 0)]
        [System.Xml.Serialization.XmlElement("permStart", typeof(CT_PermStart), Order = 0)]
        [System.Xml.Serialization.XmlElement("proofErr", typeof(CT_ProofErr), Order = 0)]
        [System.Xml.Serialization.XmlElement("r", typeof(CT_R), Order = 0)]
        [System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemsElementName")]
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

        [System.Xml.Serialization.XmlElement("ItemsElementName", Order = 1)]
        [System.Xml.Serialization.XmlIgnoreAttribute()]
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


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType16
    {

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:oMath")]
        oMath,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:oMathPara")]
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

    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
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

        [System.Xml.Serialization.XmlElement(Order = 0)]
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

        [System.Xml.Serialization.XmlElement(Order = 1)]
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

        [System.Xml.Serialization.XmlElement(Order = 2)]
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

    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Jc
    {

        private ST_Jc valField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
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


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TextDirection
    {

        private ST_TextDirection valField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_TextDirection
    {

    
        lrTb,

    
        tbRl,

    
        btLr,

    
        lrTbV,

    
        tbRlV,

    
        tbLrV,
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TextAlignment
    {

        private ST_TextAlignment valField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_TextAlignment
    {

    
        top,

    
        center,

    
        baseline,

    
        bottom,

    
        auto,
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TextboxTightWrap
    {

        private ST_TextboxTightWrap valField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_TextboxTightWrap
    {

    
        none,

    
        allLines,

    
        firstAndLastLine,

    
        firstLineOnly,

    
        lastLineOnly,
    }



    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_VerticalAlignRun
    {

        private ST_VerticalAlignRun valField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_VerticalAlignRun
    {

    
        baseline,

    
        superscript,

    
        subscript,
    }




    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
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

    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_RunTrackChange : CT_TrackChange
    {

        private List<object> itemsField;

        private List<ItemsChoiceType6> itemsElementNameField;

        public CT_RunTrackChange()
        {
            this.itemsElementNameField = new List<ItemsChoiceType6>();
            this.itemsField = new List<object>();
        }

        [System.Xml.Serialization.XmlElement("acc", typeof(CT_Acc), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("bar", typeof(CT_Bar), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("borderBox", typeof(CT_BorderBox), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("box", typeof(CT_Box), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("d", typeof(CT_D), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("eqArr", typeof(CT_EqArr), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("f", typeof(CT_F), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("func", typeof(CT_Func), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("groupChr", typeof(CT_GroupChr), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("limLow", typeof(CT_LimLow), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("limUpp", typeof(CT_LimUpp), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("m", typeof(CT_M), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("nary", typeof(CT_Nary), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("oMath", typeof(CT_OMath), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("oMathPara", typeof(CT_OMathPara), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("phant", typeof(CT_Phant), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("r", typeof(NPOI.OpenXmlFormats.Shared.CT_R), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("rad", typeof(CT_Rad), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("sPre", typeof(CT_SPre), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("sSub", typeof(CT_SSub), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("sSubSup", typeof(CT_SSubSup), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("sSup", typeof(CT_SSup), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("bookmarkEnd", typeof(CT_MarkupRange), Order = 0)]
        [System.Xml.Serialization.XmlElement("bookmarkStart", typeof(CT_Bookmark), Order = 0)]
        [System.Xml.Serialization.XmlElement("commentRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [System.Xml.Serialization.XmlElement("commentRangeStart", typeof(CT_MarkupRange), Order = 0)]
        [System.Xml.Serialization.XmlElement("customXml", typeof(CT_CustomXmlRun), Order = 0)]
        [System.Xml.Serialization.XmlElement("customXmlDelRangeEnd", typeof(CT_Markup), Order = 0)]
        [System.Xml.Serialization.XmlElement("customXmlDelRangeStart", typeof(CT_TrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElement("customXmlInsRangeEnd", typeof(CT_Markup), Order = 0)]
        [System.Xml.Serialization.XmlElement("customXmlInsRangeStart", typeof(CT_TrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElement("customXmlMoveFromRangeEnd", typeof(CT_Markup), Order = 0)]
        [System.Xml.Serialization.XmlElement("customXmlMoveFromRangeStart", typeof(CT_TrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElement("customXmlMoveToRangeEnd", typeof(CT_Markup), Order = 0)]
        [System.Xml.Serialization.XmlElement("customXmlMoveToRangeStart", typeof(CT_TrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElement("del", typeof(CT_RunTrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElement("ins", typeof(CT_RunTrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElement("moveFrom", typeof(CT_RunTrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElement("moveFromRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [System.Xml.Serialization.XmlElement("moveFromRangeStart", typeof(CT_MoveBookmark), Order = 0)]
        [System.Xml.Serialization.XmlElement("moveTo", typeof(CT_RunTrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElement("moveToRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [System.Xml.Serialization.XmlElement("moveToRangeStart", typeof(CT_MoveBookmark), Order = 0)]
        [System.Xml.Serialization.XmlElement("permEnd", typeof(CT_Perm), Order = 0)]
        [System.Xml.Serialization.XmlElement("permStart", typeof(CT_PermStart), Order = 0)]
        [System.Xml.Serialization.XmlElement("proofErr", typeof(CT_ProofErr), Order = 0)]
        [System.Xml.Serialization.XmlElement("r", typeof(CT_R), Order = 0)]
        [System.Xml.Serialization.XmlElement("sdt", typeof(CT_SdtRun), Order = 0)]
        [System.Xml.Serialization.XmlElement("smartTag", typeof(CT_SmartTagRun), Order = 0)]
        [System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemsElementName")]
        public object[] Items
        {
            get
            {
                return this.itemsField.ToArray();
            }
            set
            {
                if (value == null || value.Length == 0)
                    this.itemsField = new List<object>();
                else
                    this.itemsField = new List<object>(value);
            }
        }

        [System.Xml.Serialization.XmlElement("ItemsElementName", Order = 1)]
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public ItemsChoiceType6[] ItemsElementName
        {
            get
            {
                return this.itemsElementNameField.ToArray();
            }
            set
            {
                if (value == null || value.Length == 0)
                    this.itemsElementNameField = new List<ItemsChoiceType6>();
                else
                    this.itemsElementNameField = new List<ItemsChoiceType6>(value);
            }
        }

        public IEnumerable<CT_R> GetRList()
        {
            return GetObjectList<CT_R>(ItemsChoiceType6.r1);
        }
        #region Generic methods for object operation

        private List<T> GetObjectList<T>(ItemsChoiceType6 type) where T : class
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
        private int SizeOfArray(ItemsChoiceType6 type)
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
        private T GetObjectArray<T>(int p, ItemsChoiceType6 type) where T : class
        {
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                if (pos < 0 || pos >= this.itemsField.Count)
                    return null;
                return itemsField[pos] as T;
            }
        }
        private T InsertNewObject<T>(ItemsChoiceType6 type, int p) where T : class, new()
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
        private T AddNewObject<T>(ItemsChoiceType6 type) where T : class, new()
        {
            T t = new T();
            lock (this)
            {
                this.itemsElementNameField.Add(type);
                this.itemsField.Add(t);
            }
            return t;
        }
        private void SetObject<T>(ItemsChoiceType6 type, int p, T obj) where T : class
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
        private int GetObjectIndex(ItemsChoiceType6 type, int p)
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
        private void RemoveObject(ItemsChoiceType6 type, int p)
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
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType6
    {

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:acc")]
        acc,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:bar")]
        bar,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:borderBox")]
        borderBox,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:box")]
        box,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:d")]
        d,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:eqArr")]
        eqArr,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:f")]
        f,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:func")]
        func,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:groupChr")]
        groupChr,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:limLow")]
        limLow,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:limUpp")]
        limUpp,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:m")]
        m,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:nary")]
        nary,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:oMath")]
        oMath,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:oMathPara")]
        oMathPara,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:phant")]
        phant,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:r")]
        r,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:rad")]
        rad,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:sPre")]
        sPre,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:sSub")]
        sSub,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:sSubSup")]
        sSubSup,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:sSup")]
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

    
        [System.Xml.Serialization.XmlEnumAttribute("r")]
        r1,

    
        sdt,

    
        smartTag,
    }
    
}