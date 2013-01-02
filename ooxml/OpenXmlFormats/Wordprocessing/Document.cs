using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Serialization;
using NPOI.OpenXmlFormats.Shared;

namespace NPOI.OpenXmlFormats.Wordprocessing
{
    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot("glossaryDocument", Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = false)]
    public class CT_GlossaryDocument : CT_DocumentBase
    {

        private CT_DocParts docPartsField;

        public CT_GlossaryDocument()
        {
            this.docPartsField = new CT_DocParts();
        }

        [XmlElement(Order = 0)]
        public CT_DocParts docParts
        {
            get
            {
                return this.docPartsField;
            }
            set
            {
                this.docPartsField = value;
            }
        }
    }
    [XmlInclude(typeof(CT_GlossaryDocument))]
    [XmlInclude(typeof(CT_Document))]
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocumentBase
    {

        private CT_Background backgroundField;

        public CT_DocumentBase()
        {
            //this.backgroundField = new CT_Background();
        }

        [XmlElement(Order = 0)]
        public CT_Background background
        {
            get
            {
                return this.backgroundField;
            }
            set
            {
                this.backgroundField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot("document", Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = false)]
    public class CT_Document : CT_DocumentBase
    {

        private CT_Body bodyField;

        public CT_Document()
        {
            this.bodyField = new CT_Body();
        }

        [XmlElement(Order = 0)]
        public CT_Body body
        {
            get
            {
                return this.bodyField;
            }
            set
            {
                this.bodyField = value;
            }
        }

        public void AddNewBody()
        {
            this.bodyField = new CT_Body();
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Body
    {

        private List<object> itemsField;

        private List<DocumentBodyItemChoiceType> itemsElementNameField;

        private CT_SectPr sectPrField;

        public CT_Body()
        {
            this.sectPrField = new CT_SectPr();
            this.itemsElementNameField = new List<DocumentBodyItemChoiceType>();
            this.itemsField = new List<object>();
        }

        [XmlElement("oMath", typeof(CT_OMath), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("oMathPara", typeof(CT_OMathPara), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [XmlElement("altChunk", typeof(CT_AltChunk), Order = 0)]
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
        [XmlElement("del", Type = typeof(CT_RunTrackChange), Order = 0)]
        [XmlElement("ins", Type = typeof(CT_RunTrackChange), Order = 0)]
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
        public object[] Items
        {
            get
            {
                return this.itemsField.ToArray();
            }
            set
            {
                if (value != null && value.Length > 0)
                    this.itemsField = new List<object>(value);
                else
                    this.itemsField = new List<object>();
            }
        }

        public CT_P AddNewP()
        {
            CT_P p = new CT_P();
            lock (this)
            {
                this.itemsField.Add(p);
                this.itemsElementNameField.Add(DocumentBodyItemChoiceType.p);
            }
            return p;
        }
        
        [XmlElement("ItemsElementName", Order = 1)]
        [XmlIgnore]
        public DocumentBodyItemChoiceType[] ItemsElementName
        {
            get
            {
                return this.itemsElementNameField.ToArray();
            }
            set
            {
                if (value != null && value.Length > 0)
                    this.itemsElementNameField = new List<DocumentBodyItemChoiceType>(value);
                else
                    this.itemsElementNameField = new List<DocumentBodyItemChoiceType>();
            }
        }

        [XmlElement(Order = 2)]
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
        public CT_Tbl AddNewTbl()
        {
            CT_Tbl tbl = new CT_Tbl();
            lock (this)
            {
                this.itemsField.Add(tbl);
                this.itemsElementNameField.Add(DocumentBodyItemChoiceType.tbl);
            }
            return tbl;
        }
        public int sizeOfTblArray()
        {
            throw new NotImplementedException();
        }
        public CT_Tbl[] getTblArray()
        {
            throw new NotImplementedException();
        }
        public CT_Tbl insertNewTbl(int paramInt)
        {
            throw new NotImplementedException();
        }
        public void removeTbl(int paramInt)
        {
            throw new NotImplementedException();
        }
        public CT_Tbl GetTblArray(int i)
        {
            return GetObjectArray<CT_Tbl>(i, DocumentBodyItemChoiceType.tbl);
        }

        public void SetTblArray(int pos, CT_Tbl cT_Tbl)
        {
            SetObject<CT_Tbl>(DocumentBodyItemChoiceType.tbl, pos, cT_Tbl);
        }

        public CT_Tbl[] GetTblArray()
        {
            return GetObjectList<CT_Tbl>(DocumentBodyItemChoiceType.tbl).ToArray();
        }

        public CT_P GetPArray(int p)
        {
            return GetObjectArray<CT_P>(p, DocumentBodyItemChoiceType.p);
        }
        public void RemoveP(int paraPos)
        {
            RemoveObject(DocumentBodyItemChoiceType.p, paraPos);
        }
        public void RemoveTbl(int tablePos)
        {
            RemoveObject(DocumentBodyItemChoiceType.tbl, tablePos);
        }

        public CT_SdtBlock AddNewSdt()
        {
            return AddNewObject<CT_SdtBlock>(DocumentBodyItemChoiceType.sdt);
        }
        #region Generic methods for object operation
  
        private List<T> GetObjectList<T>(DocumentBodyItemChoiceType type) where T : class
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
        private int SizeOfArray(DocumentBodyItemChoiceType type)
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
        private T GetObjectArray<T>(int p, DocumentBodyItemChoiceType type) where T : class
        {
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                if (pos < 0 || pos >= this.itemsField.Count)
                    return null;
                return itemsField[pos] as T;
            }
        }
        private T AddNewObject<T>(DocumentBodyItemChoiceType type) where T : class, new()
        {
            T t = new T();
            lock (this)
            {
                this.itemsElementNameField.Add(type);
                this.itemsField.Add(t);
            }
            return t;
        }
        private void SetObject<T> (DocumentBodyItemChoiceType type, int p, T obj) where T : class
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
        private int GetObjectIndex(DocumentBodyItemChoiceType type, int p)
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
        private void RemoveObject(DocumentBodyItemChoiceType type, int p)
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




        public void SetPArray(int pos, CT_P cT_P)
        {
            SetObject<CT_P>(DocumentBodyItemChoiceType.p, pos, cT_P);
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum DocumentBodyItemChoiceType
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

        /// <summary>
        /// Paragraph
        /// </summary>
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
    public class CT_DocParts
    {

        private List<CT_DocPart> itemsField;

        public CT_DocParts()
        {
            this.itemsField = new List<CT_DocPart>();
        }

        [XmlElement("docPart", Order = 0)]
        public List<CT_DocPart> Items
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
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocPart
    {

        private CT_DocPartPr docPartPrField;

        private CT_Body docPartBodyField;

        public CT_DocPart()
        {
            this.docPartBodyField = new CT_Body();
            this.docPartPrField = new CT_DocPartPr();
        }

        [XmlElement(Order = 0)]
        public CT_DocPartPr docPartPr
        {
            get
            {
                return this.docPartPrField;
            }
            set
            {
                this.docPartPrField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_Body docPartBody
        {
            get
            {
                return this.docPartBodyField;
            }
            set
            {
                this.docPartBodyField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocPartPr
    {

        private object[] itemsField;

        private ItemsChoiceType11[] itemsElementNameField;

        public CT_DocPartPr()
        {
            this.itemsElementNameField = new ItemsChoiceType11[0];
            this.itemsField = new object[0];
        }

        [XmlElement("behaviors", typeof(CT_DocPartBehaviors), Order = 0)]
        [XmlElement("category", typeof(CT_DocPartCategory), Order = 0)]
        [XmlElement("description", typeof(CT_String), Order = 0)]
        [XmlElement("guid", typeof(CT_Guid), Order = 0)]
        [XmlElement("name", typeof(CT_DocPartName), Order = 0)]
        [XmlElement("style", typeof(CT_String), Order = 0)]
        [XmlElement("types", typeof(CT_DocPartTypes), Order = 0)]
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
        public ItemsChoiceType11[] ItemsElementName
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
    public class CT_DocPartBehaviors
    {

        private List<CT_DocPartBehavior> itemsField;

        public CT_DocPartBehaviors()
        {
            this.itemsField = new List<CT_DocPartBehavior>();
        }

        [XmlElement("behavior", Order = 0)]
        public List<CT_DocPartBehavior> Items
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
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocPartBehavior
    {

        private ST_DocPartBehavior valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_DocPartBehavior val
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
    public enum ST_DocPartBehavior
    {

    
        content,

    
        p,

    
        pg,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocPartCategory
    {

        private CT_String nameField;

        private CT_DocPartGallery galleryField;

        public CT_DocPartCategory()
        {
            this.galleryField = new CT_DocPartGallery();
            this.nameField = new CT_String();
        }

        [XmlElement(Order = 0)]
        public CT_String name
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

        [XmlElement(Order = 1)]
        public CT_DocPartGallery gallery
        {
            get
            {
                return this.galleryField;
            }
            set
            {
                this.galleryField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocPartGallery
    {

        private ST_DocPartGallery valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_DocPartGallery val
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
    public enum ST_DocPartGallery
    {

    
        placeholder,

    
        any,

    
        @default,

    
        docParts,

    
        coverPg,

    
        eq,

    
        ftrs,

    
        hdrs,

    
        pgNum,

    
        tbls,

    
        watermarks,

    
        autoTxt,

    
        txtBox,

    
        pgNumT,

    
        pgNumB,

    
        pgNumMargins,

    
        tblOfContents,

    
        bib,

    
        custQuickParts,

    
        custCoverPg,

    
        custEq,

    
        custFtrs,

    
        custHdrs,

    
        custPgNum,

    
        custTbls,

    
        custWatermarks,

    
        custAutoTxt,

    
        custTxtBox,

    
        custPgNumT,

    
        custPgNumB,

    
        custPgNumMargins,

    
        custTblOfContents,

    
        custBib,

    
        custom1,

    
        custom2,

    
        custom3,

    
        custom4,

    
        custom5,
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocPartName
    {

        private string valField;

        private ST_OnOff decoratedField;

        private bool decoratedFieldSpecified;

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

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff decorated
        {
            get
            {
                return this.decoratedField;
            }
            set
            {
                this.decoratedField = value;
            }
        }

        [XmlIgnore]
        public bool decoratedSpecified
        {
            get
            {
                return this.decoratedFieldSpecified;
            }
            set
            {
                this.decoratedFieldSpecified = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocPartTypes
    {

        private List<CT_DocPartType> itemsField;

        private ST_OnOff allField;

        private bool allFieldSpecified;

        public CT_DocPartTypes()
        {
            this.itemsField = new List<CT_DocPartType>();
        }

        [XmlElement("type", Order = 0)]
        public List<CT_DocPartType> Items
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

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff all
        {
            get
            {
                return this.allField;
            }
            set
            {
                this.allField = value;
            }
        }

        [XmlIgnore]
        public bool allSpecified
        {
            get
            {
                return this.allFieldSpecified;
            }
            set
            {
                this.allFieldSpecified = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocPartType
    {

        private ST_DocPartType valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_DocPartType val
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
    public enum ST_DocPartType
    {

    
        none,

    
        normal,

    
        autoExp,

    
        toolbar,

    
        speller,

    
        formFld,

    
        bbPlcHdr,
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType11
    {

    
        behaviors,

    
        category,

    
        description,

    
        guid,

    
        name,

    
        style,

    
        types,
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocGrid
    {

        private ST_DocGrid typeField;

        private bool typeFieldSpecified;

        private string linePitchField;

        private string charSpaceField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_DocGrid type
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

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string linePitch
        {
            get
            {
                return this.linePitchField;
            }
            set
            {
                this.linePitchField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string charSpace
        {
            get
            {
                return this.charSpaceField;
            }
            set
            {
                this.charSpaceField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_DocGrid
    {

    
        @default,

    
        lines,

    
        linesAndChars,

    
        snapToChars,
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocVars
    {

        private List<CT_DocVar> docVarField;

        public CT_DocVars()
        {
            this.docVarField = new List<CT_DocVar>();
        }

        [XmlElement("docVar", Order = 0)]
        public List<CT_DocVar> docVar
        {
            get
            {
                return this.docVarField;
            }
            set
            {
                this.docVarField = value;
            }
        }
    }

}
