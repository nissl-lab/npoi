using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Serialization;
using NPOI.OpenXmlFormats.Shared;

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

        private List<object> itemsField;

        private List<ItemsChoiceType6> itemsElementNameField;

        public CT_RunTrackChange()
        {
            this.itemsElementNameField = new List<ItemsChoiceType6>();
            this.itemsField = new List<object>();
        }

        [XmlElement("acc", typeof(CT_Acc), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("bar", typeof(CT_Bar), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("borderBox", typeof(CT_BorderBox), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("box", typeof(CT_Box), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("d", typeof(CT_D), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("eqArr", typeof(CT_EqArr), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("f", typeof(CT_F), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("func", typeof(CT_Func), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("groupChr", typeof(CT_GroupChr), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("limLow", typeof(CT_LimLow), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("limUpp", typeof(CT_LimUpp), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("m", typeof(CT_M), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("nary", typeof(CT_Nary), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("oMath", typeof(CT_OMath), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("oMathPara", typeof(CT_OMathPara), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("phant", typeof(CT_Phant), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("r", typeof(NPOI.OpenXmlFormats.Shared.CT_R), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("rad", typeof(CT_Rad), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("sPre", typeof(CT_SPre), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("sSub", typeof(CT_SSub), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("sSubSup", typeof(CT_SSubSup), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("sSup", typeof(CT_SSup), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
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
        [XmlChoiceIdentifier("ItemsElementName")]
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

        [XmlElement("ItemsElementName", Order = 1)]
        [XmlIgnore]
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