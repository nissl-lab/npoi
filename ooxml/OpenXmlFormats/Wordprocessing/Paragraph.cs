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
    public class CT_P
    {

        private CT_PPr pPrField;

        private List<object> itemsField;

        private List<ParagraphItemsChoiceType> itemsElementNameField;

        private byte[] rsidRPrField;

        private byte[] rsidRField;

        private byte[] rsidDelField;

        private byte[] rsidPField;

        private byte[] rsidRDefaultField;

        public CT_P()
        {
            this.itemsElementNameField = new List<ParagraphItemsChoiceType>();
            this.itemsField = new List<object>();
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

        [XmlElement("ItemsElementName", Order = 2)]
        [XmlIgnore]
        public ParagraphItemsChoiceType[] ItemsElementName
        {
            get
            {
                return this.itemsElementNameField.ToArray();
            }
            set
            {
                if (value != null && value.Length != 0)
                    this.itemsElementNameField = new List<ParagraphItemsChoiceType>(value);
                else
                    this.itemsElementNameField = new List<ParagraphItemsChoiceType>();
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
            //this.pPrChangeField = new CT_PPrChange();
            //this.sectPrField = new CT_SectPr();
            //this.rPrField = new CT_ParaRPr();
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
            this.rPrChangeField = new CT_ParaRPrChange();
            this.oMathField = new CT_OnOff();
            this.specVanishField = new CT_OnOff();
            this.eastAsianLayoutField = new CT_EastAsianLayout();
            this.langField = new CT_Language();
            this.emField = new CT_Em();
            this.csField = new CT_OnOff();
            this.rtlField = new CT_OnOff();
            this.vertAlignField = new CT_VerticalAlignRun();
            this.fitTextField = new CT_FitText();
            this.shdField = new CT_Shd();
            this.bdrField = new CT_Border();
            this.effectField = new CT_TextEffect();
            this.uField = new CT_Underline();
            this.highlightField = new CT_Highlight();
            this.szCsField = new CT_HpsMeasure();
            this.szField = new CT_HpsMeasure();
            this.positionField = new CT_SignedHpsMeasure();
            this.kernField = new CT_HpsMeasure();
            this.wField = new CT_TextScale();
            this.spacingField = new CT_SignedTwipsMeasure();
            this.colorField = new CT_Color();
            this.webHiddenField = new CT_OnOff();
            this.vanishField = new CT_OnOff();
            this.snapToGridField = new CT_OnOff();
            this.noProofField = new CT_OnOff();
            this.imprintField = new CT_OnOff();
            this.embossField = new CT_OnOff();
            this.shadowField = new CT_OnOff();
            this.outlineField = new CT_OnOff();
            this.dstrikeField = new CT_OnOff();
            this.strikeField = new CT_OnOff();
            this.smallCapsField = new CT_OnOff();
            this.capsField = new CT_OnOff();
            this.iCsField = new CT_OnOff();
            this.iField = new CT_OnOff();
            this.bCsField = new CT_OnOff();
            this.bField = new CT_OnOff();
            this.rFontsField = new CT_Fonts();
            this.rStyleField = new CT_String();
            this.moveToField = new CT_TrackChange();
            this.moveFromField = new CT_TrackChange();
            this.delField = new CT_TrackChange();
            this.insField = new CT_TrackChange();
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
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SectPr
    {

        private List<CT_HdrFtrRef> itemsField;

        private List<ItemsChoiceHdrFtrRefType> itemsElementNameField;

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
            //this.sectPrChangeField = new CT_SectPrChange();
            //this.printerSettingsField = new CT_Rel();
            this.docGridField = new CT_DocGrid();
            this.docGrid.type = ST_DocGrid.lines;
            this.docGrid.typeSpecified = true;
            this.docGrid.linePitch = "312";
            
            //this.rtlGutterField = new CT_OnOff();
            //this.bidiField = new CT_OnOff();
            //this.textDirectionField = new CT_TextDirection();
            //this.titlePgField = new CT_OnOff();
            //this.noEndnoteField = new CT_OnOff();
            //this.vAlignField = new CT_VerticalJc();
            //this.formProtField = new CT_OnOff();
            this.colsField = new CT_Columns();
            this.cols.space = 425;
            this.cols.spaceSpecified = true;
            //this.pgNumTypeField = new CT_PageNumber();
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
            this.itemsElementNameField = new List<ItemsChoiceHdrFtrRefType>();
            this.itemsField = new List<CT_HdrFtrRef>();
        }

        [XmlElement("footerReference", typeof(CT_HdrFtrRef), Order = 0)]
        [XmlElement("headerReference", typeof(CT_HdrFtrRef), Order = 0)]
        [XmlChoiceIdentifier("ItemsElementName")]
        public CT_HdrFtrRef[] Items
        {
            get
            {
                return this.itemsField.ToArray();
            }
            set
            {
                if (value != null && value.Length > 0)
                {
                    this.itemsField = new List<CT_HdrFtrRef>(value);
                }
                else
                {
                    this.itemsField = new List<CT_HdrFtrRef>();
                }
            }
        }

        [XmlElement("ItemsElementName", Order = 1)]
        [XmlIgnore]
        public ItemsChoiceHdrFtrRefType[] ItemsElementName
        {
            get
            {
                return this.itemsElementNameField.ToArray();
            }
            set
            {
                if (value != null && value.Length > 0)
                {
                    this.itemsElementNameField = new List<ItemsChoiceHdrFtrRefType>(value);
                }
                else
                {
                    this.itemsElementNameField = new List<ItemsChoiceHdrFtrRefType>();
                }
            }
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
        private CT_HdrFtrRef AddNewReference(ItemsChoiceHdrFtrRefType type)
        {
            CT_HdrFtrRef ref1 = new CT_HdrFtrRef();
            lock (this)
            {
                itemsField.Add(ref1);
                itemsElementNameField.Add(type);
            }
            return ref1;
        }
        public CT_HdrFtrRef AddNewHeaderReference()
        {
            return AddNewReference(ItemsChoiceHdrFtrRefType.headerReference);
        }

        public CT_HdrFtrRef AddNewFooterReference()
        {
            return AddNewReference(ItemsChoiceHdrFtrRefType.footerReference);
        }
        private int SizeOfObjectArray(ItemsChoiceHdrFtrRefType type)
        {
            int size = 0;
            for (int i = 0; i < itemsElementNameField.Count; i++)
            {
                if (itemsElementNameField[i] == type)
                    size++;
            }
            return size;
        }
        public int SizeOfHeaderReferenceArray()
        {
            return SizeOfObjectArray(ItemsChoiceHdrFtrRefType.headerReference);
        }
        private CT_HdrFtrRef GetObjectArray(int i, ItemsChoiceHdrFtrRefType type)
        {
            int pos = 0;
            for (int p = 0; p < itemsElementNameField.Count; p++)
            {
                if (itemsElementNameField[p] == type)
                {
                    if (pos == i)
                        return itemsField[p] as CT_HdrFtrRef;
                    else
                        pos++;
                }
            }
            return null;
        }
        public CT_HdrFtrRef GetHeaderReferenceArray(int i)
        {
            return GetObjectArray(i, ItemsChoiceHdrFtrRefType.headerReference);
        }

        public int SizeOfFooterReferenceArray()
        {
            return SizeOfObjectArray(ItemsChoiceHdrFtrRefType.footerReference);
        }

        public CT_HdrFtrRef GetFooterReferenceArray(int i)
        {
            return GetObjectArray(i, ItemsChoiceHdrFtrRefType.footerReference);
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
            this.barField = new CT_Border();
            this.betweenField = new CT_Border();
            this.rightField = new CT_Border();
            this.bottomField = new CT_Border();
            this.leftField = new CT_Border();
            this.topField = new CT_Border();
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
            this.bottomField = new CT_Border();
        }

        public bool IsSetRight()
        {
            return this.rightField != null && this.rightField.val != ST_Border.none && this.rightField.val != ST_Border.nil;
        }

        public void UnsetRight()
        {
            this.rightField = new CT_Border();
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
            this.betweenField = new CT_Border();
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
            this.leftField = new CT_Border();
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

        private ulong firstLineField;

        private bool firstLineFieldSpecified;

        private string firstLineCharsField;

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
        public ulong firstLine
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
            //this.rPrChangeField = new CT_RPrChange();
            //this.oMathField = new CT_OnOff();
            //this.specVanishField = new CT_OnOff();
            //this.eastAsianLayoutField = new CT_EastAsianLayout();
            //this.langField = new CT_Language();
            //this.emField = new CT_Em();
            //this.csField = new CT_OnOff();
            //this.rtlField = new CT_OnOff();
            //this.vertAlignField = new CT_VerticalAlignRun();
            //this.fitTextField = new CT_FitText();
            //this.shdField = new CT_Shd();
            //this.bdrField = new CT_Border();
            //this.effectField = new CT_TextEffect();
            //this.uField = new CT_Underline();
            //this.highlightField = new CT_Highlight();
            //this.szCsField = new CT_HpsMeasure();
            //this.szField = new CT_HpsMeasure();
            //this.positionField = new CT_SignedHpsMeasure();
            //this.kernField = new CT_HpsMeasure();
            //this.wField = new CT_TextScale();
            //this.spacingField = new CT_SignedTwipsMeasure();
            //this.colorField = new CT_Color();
            //this.webHiddenField = new CT_OnOff();
            //this.vanishField = new CT_OnOff();
            //this.snapToGridField = new CT_OnOff();
            //this.noProofField = new CT_OnOff();
            //this.imprintField = new CT_OnOff();
            //this.embossField = new CT_OnOff();
            //this.shadowField = new CT_OnOff();
            //this.outlineField = new CT_OnOff();
            //this.dstrikeField = new CT_OnOff();
            //this.strikeField = new CT_OnOff();
            //this.smallCapsField = new CT_OnOff();
            //this.capsField = new CT_OnOff();
            //this.iCsField = new CT_OnOff();
            //this.iField = new CT_OnOff();
            //this.bCsField = new CT_OnOff();
            //this.bField = new CT_OnOff();
            //this.rFontsField = new CT_Fonts();
            //this.rStyleField = new CT_String();
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
                this.bField.valSpecified = this.bField.val == ST_OnOff.on || 
                    this.bField.val == ST_OnOff.True || this.bField.val == ST_OnOff.Value1;
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
                this.iField.valSpecified = this.iField.val == ST_OnOff.on ||
                    this.iField.val == ST_OnOff.True || this.iField.val == ST_OnOff.Value1;
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
            return !(this.uField.val == ST_Underline.none);
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
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_ParaRPrChange : CT_TrackChange
    {

        private CT_ParaRPrOriginal rPrField;

        public CT_ParaRPrChange()
        {
            this.rPrField = new CT_ParaRPrOriginal();
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

        private object[] itemsField;

        private ItemsChoiceType3[] itemsElementNameField;

        public CT_ParaRPrOriginal()
        {
            this.itemsElementNameField = new ItemsChoiceType3[0];
            this.itemsField = new object[0];
            this.moveToField = new CT_TrackChange();
            this.moveFromField = new CT_TrackChange();
            this.delField = new CT_TrackChange();
            this.insField = new CT_TrackChange();
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

        [XmlElement("b", typeof(CT_OnOff), Order = 4)]
        [XmlElement("bCs", typeof(CT_OnOff), Order = 4)]
        [XmlElement("bdr", typeof(CT_Border), Order = 4)]
        [XmlElement("caps", typeof(CT_OnOff), Order = 4)]
        [XmlElement("color", typeof(CT_Color), Order = 4)]
        [XmlElement("cs", typeof(CT_OnOff), Order = 4)]
        [XmlElement("dstrike", typeof(CT_OnOff), Order = 4)]
        [XmlElement("eastAsianLayout", typeof(CT_EastAsianLayout), Order = 4)]
        [XmlElement("effect", typeof(CT_TextEffect), Order = 4)]
        [XmlElement("em", typeof(CT_Em), Order = 4)]
        [XmlElement("emboss", typeof(CT_OnOff), Order = 4)]
        [XmlElement("fitText", typeof(CT_FitText), Order = 4)]
        [XmlElement("highlight", typeof(CT_Highlight), Order = 4)]
        [XmlElement("i", typeof(CT_OnOff), Order = 4)]
        [XmlElement("iCs", typeof(CT_OnOff), Order = 4)]
        [XmlElement("imprint", typeof(CT_OnOff), Order = 4)]
        [XmlElement("kern", typeof(CT_HpsMeasure), Order = 4)]
        [XmlElement("lang", typeof(CT_Language), Order = 4)]
        [XmlElement("noProof", typeof(CT_OnOff), Order = 4)]
        [XmlElement("oMath", typeof(CT_OnOff), Order = 4)]
        [XmlElement("outline", typeof(CT_OnOff), Order = 4)]
        [XmlElement("position", typeof(CT_SignedHpsMeasure), Order = 4)]
        [XmlElement("rFonts", typeof(CT_Fonts), Order = 4)]
        [XmlElement("rStyle", typeof(CT_String), Order = 4)]
        [XmlElement("rtl", typeof(CT_OnOff), Order = 4)]
        [XmlElement("shadow", typeof(CT_OnOff), Order = 4)]
        [XmlElement("shd", typeof(CT_Shd), Order = 4)]
        [XmlElement("smallCaps", typeof(CT_OnOff), Order = 4)]
        [XmlElement("snapToGrid", typeof(CT_OnOff), Order = 4)]
        [XmlElement("spacing", typeof(CT_SignedTwipsMeasure), Order = 4)]
        [XmlElement("specVanish", typeof(CT_OnOff), Order = 4)]
        [XmlElement("strike", typeof(CT_OnOff), Order = 4)]
        [XmlElement("sz", typeof(CT_HpsMeasure), Order = 4)]
        [XmlElement("szCs", typeof(CT_HpsMeasure), Order = 4)]
        [XmlElement("u", typeof(CT_Underline), Order = 4)]
        [XmlElement("vanish", typeof(CT_OnOff), Order = 4)]
        [XmlElement("vertAlign", typeof(CT_VerticalAlignRun), Order = 4)]
        [XmlElement("w", typeof(CT_TextScale), Order = 4)]
        [XmlElement("webHidden", typeof(CT_OnOff), Order = 4)]
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

        [XmlElement("ItemsElementName", Order = 5)]
        [XmlIgnore]
        public ItemsChoiceType3[] ItemsElementName
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
            this.pPrField = new CT_PPrBase();
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

        public CT_PPrBase()
        {
            //this.cnfStyleField = new CT_Cnf();
            //this.divIdField = new CT_DecimalNumber();
            //this.outlineLvlField = new CT_DecimalNumber();
            //this.textboxTightWrapField = new CT_TextboxTightWrap();
            //this.textAlignmentField = new CT_TextAlignment();
            //this.textDirectionField = new CT_TextDirection();
            //this.jcField = new CT_Jc();
            //this.suppressOverlapField = new CT_OnOff();
            //this.mirrorIndentsField = new CT_OnOff();
            //this.contextualSpacingField = new CT_OnOff();
            //this.indField = new CT_Ind();
            //this.spacingField = new CT_Spacing();
            //this.snapToGridField = new CT_OnOff();
            //this.adjustRightIndField = new CT_OnOff();
            //this.bidiField = new CT_OnOff();
            //this.autoSpaceDNField = new CT_OnOff();
            //this.autoSpaceDEField = new CT_OnOff();
            //this.topLinePunctField = new CT_OnOff();
            //this.overflowPunctField = new CT_OnOff();
            //this.wordWrapField = new CT_OnOff();
            //this.kinsokuField = new CT_OnOff();
            //this.suppressAutoHyphensField = new CT_OnOff();
            //this.tabsField = new List<CT_TabStop>();
            //this.shdField = new CT_Shd();
            //this.pBdrField = new CT_PBdr();
            //this.suppressLineNumbersField = new CT_OnOff();
            //this.numPrField = new CT_NumPr();
            //this.widowControlField = new CT_OnOff();
            //this.framePrField = new CT_FramePr();
            //this.pageBreakBeforeField = new CT_OnOff();
            //this.keepLinesField = new CT_OnOff();
            //this.keepNextField = new CT_OnOff();
            //this.pStyleField = new CT_String();
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
            return this.pageBreakBeforeField.val == ST_OnOff.on || this.pageBreakBeforeField.val == ST_OnOff.True ||
                this.pageBreakBeforeField.val == ST_OnOff.Value1;
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
            return this.wordWrapField.val == ST_OnOff.Value1 || this.wordWrapField.val == ST_OnOff.True ||
                this.wordWrapField.val == ST_OnOff.on;
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
            this.rPrField = new CT_RPrOriginal();
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

        private object[] itemsField;

        private ItemsChoiceType5[] itemsElementNameField;

        public CT_RPrOriginal()
        {
            this.itemsElementNameField = new ItemsChoiceType5[0];
            this.itemsField = new object[0];
        }

        [XmlElement("b", typeof(CT_OnOff), Order = 0)]
        [XmlElement("bCs", typeof(CT_OnOff), Order = 0)]
        [XmlElement("bdr", typeof(CT_Border), Order = 0)]
        [XmlElement("caps", typeof(CT_OnOff), Order = 0)]
        [XmlElement("color", typeof(CT_Color), Order = 0)]
        [XmlElement("cs", typeof(CT_OnOff), Order = 0)]
        [XmlElement("dstrike", typeof(CT_OnOff), Order = 0)]
        [XmlElement("eastAsianLayout", typeof(CT_EastAsianLayout), Order = 0)]
        [XmlElement("effect", typeof(CT_TextEffect), Order = 0)]
        [XmlElement("em", typeof(CT_Em), Order = 0)]
        [XmlElement("emboss", typeof(CT_OnOff), Order = 0)]
        [XmlElement("fitText", typeof(CT_FitText), Order = 0)]
        [XmlElement("highlight", typeof(CT_Highlight), Order = 0)]
        [XmlElement("i", typeof(CT_OnOff), Order = 0)]
        [XmlElement("iCs", typeof(CT_OnOff), Order = 0)]
        [XmlElement("imprint", typeof(CT_OnOff), Order = 0)]
        [XmlElement("kern", typeof(CT_HpsMeasure), Order = 0)]
        [XmlElement("lang", typeof(CT_Language), Order = 0)]
        [XmlElement("noProof", typeof(CT_OnOff), Order = 0)]
        [XmlElement("oMath", typeof(CT_OnOff), Order = 0)]
        [XmlElement("outline", typeof(CT_OnOff), Order = 0)]
        [XmlElement("position", typeof(CT_SignedHpsMeasure), Order = 0)]
        [XmlElement("rFonts", typeof(CT_Fonts), Order = 0)]
        [XmlElement("rStyle", typeof(CT_String), Order = 0)]
        [XmlElement("rtl", typeof(CT_OnOff), Order = 0)]
        [XmlElement("shadow", typeof(CT_OnOff), Order = 0)]
        [XmlElement("shd", typeof(CT_Shd), Order = 0)]
        [XmlElement("smallCaps", typeof(CT_OnOff), Order = 0)]
        [XmlElement("snapToGrid", typeof(CT_OnOff), Order = 0)]
        [XmlElement("spacing", typeof(CT_SignedTwipsMeasure), Order = 0)]
        [XmlElement("specVanish", typeof(CT_OnOff), Order = 0)]
        [XmlElement("strike", typeof(CT_OnOff), Order = 0)]
        [XmlElement("sz", typeof(CT_HpsMeasure), Order = 0)]
        [XmlElement("szCs", typeof(CT_HpsMeasure), Order = 0)]
        [XmlElement("u", typeof(CT_Underline), Order = 0)]
        [XmlElement("vanish", typeof(CT_OnOff), Order = 0)]
        [XmlElement("vertAlign", typeof(CT_VerticalAlignRun), Order = 0)]
        [XmlElement("w", typeof(CT_TextScale), Order = 0)]
        [XmlElement("webHidden", typeof(CT_OnOff), Order = 0)]
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
        public ItemsChoiceType5[] ItemsElementName
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
    public class CT_Br
    {

        private ST_BrType typeField;

        private bool typeFieldSpecified;

        private ST_BrClear clearField;

        private bool clearFieldSpecified;

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

        public CT_TblStylePr()
        {
            this.tcPrField = new CT_TcPr();
            this.trPrField = new CT_TrPr();
            this.tblPrField = new CT_TblPrBase();
            this.rPrField = new CT_RPr();
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