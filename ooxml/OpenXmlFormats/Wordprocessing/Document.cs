using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using NPOI.OpenXmlFormats.Shared;

namespace NPOI.OpenXmlFormats.Wordprocessing
{
    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute("glossaryDocument", Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = false)]
    public class CT_GlossaryDocument : CT_DocumentBase
    {

        private CT_DocParts docPartsField;

        public CT_GlossaryDocument()
        {
            this.docPartsField = new CT_DocParts();
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
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
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(CT_GlossaryDocument))]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(CT_Document))]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocumentBase
    {

        private CT_Background backgroundField;

        public CT_DocumentBase()
        {
            this.backgroundField = new CT_Background();
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
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


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute("document", Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = false)]
    public class CT_Document : CT_DocumentBase
    {

        private CT_Body bodyField;

        public CT_Document()
        {
            this.bodyField = new CT_Body();
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
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
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
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
        [System.Xml.Serialization.XmlElementAttribute("del", Type = typeof(CT_RunTrackChange), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("ins", Type = typeof(CT_RunTrackChange), Order = 0)]
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
            lock(this)
            {
            this.itemsField.Add(p);
            this.itemsElementNameField.Add(DocumentBodyItemChoiceType.p);
            }
            return p;
        }
        
        [System.Xml.Serialization.XmlElementAttribute("ItemsElementName", Order = 1)]
        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

        [System.Xml.Serialization.XmlElementAttribute(Order = 2)]
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


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum DocumentBodyItemChoiceType
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

        /// <summary>
        /// Paragraph
        /// </summary>
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


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocParts
    {

        private List<CT_DocPart> itemsField;

        public CT_DocParts()
        {
            this.itemsField = new List<CT_DocPart>();
        }

        [System.Xml.Serialization.XmlElementAttribute("docPart", Order = 0)]
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


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocPart
    {

        private CT_DocPartPr docPartPrField;

        private CT_Body docPartBodyField;

        public CT_DocPart()
        {
            this.docPartBodyField = new CT_Body();
            this.docPartPrField = new CT_DocPartPr();
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
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

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
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


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocPartPr
    {

        private object[] itemsField;

        private ItemsChoiceType11[] itemsElementNameField;

        public CT_DocPartPr()
        {
            this.itemsElementNameField = new ItemsChoiceType11[0];
            this.itemsField = new object[0];
        }

        [System.Xml.Serialization.XmlElementAttribute("behaviors", typeof(CT_DocPartBehaviors), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("category", typeof(CT_DocPartCategory), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("description", typeof(CT_String), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("guid", typeof(CT_Guid), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("name", typeof(CT_DocPartName), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("style", typeof(CT_String), Order = 0)]
        [System.Xml.Serialization.XmlElementAttribute("types", typeof(CT_DocPartTypes), Order = 0)]
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

        [System.Xml.Serialization.XmlElementAttribute("ItemsElementName", Order = 1)]
        [System.Xml.Serialization.XmlIgnoreAttribute()]
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


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocPartBehaviors
    {

        private List<CT_DocPartBehavior> itemsField;

        public CT_DocPartBehaviors()
        {
            this.itemsField = new List<CT_DocPartBehavior>();
        }

        [System.Xml.Serialization.XmlElementAttribute("behavior", Order = 0)]
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


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocPartBehavior
    {

        private ST_DocPartBehavior valField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_DocPartBehavior
    {

        /// <remarks/>
        content,

        /// <remarks/>
        p,

        /// <remarks/>
        pg,
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocPartCategory
    {

        private CT_String nameField;

        private CT_DocPartGallery galleryField;

        public CT_DocPartCategory()
        {
            this.galleryField = new CT_DocPartGallery();
            this.nameField = new CT_String();
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
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

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
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


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocPartGallery
    {

        private ST_DocPartGallery valField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_DocPartGallery
    {

        /// <remarks/>
        placeholder,

        /// <remarks/>
        any,

        /// <remarks/>
        @default,

        /// <remarks/>
        docParts,

        /// <remarks/>
        coverPg,

        /// <remarks/>
        eq,

        /// <remarks/>
        ftrs,

        /// <remarks/>
        hdrs,

        /// <remarks/>
        pgNum,

        /// <remarks/>
        tbls,

        /// <remarks/>
        watermarks,

        /// <remarks/>
        autoTxt,

        /// <remarks/>
        txtBox,

        /// <remarks/>
        pgNumT,

        /// <remarks/>
        pgNumB,

        /// <remarks/>
        pgNumMargins,

        /// <remarks/>
        tblOfContents,

        /// <remarks/>
        bib,

        /// <remarks/>
        custQuickParts,

        /// <remarks/>
        custCoverPg,

        /// <remarks/>
        custEq,

        /// <remarks/>
        custFtrs,

        /// <remarks/>
        custHdrs,

        /// <remarks/>
        custPgNum,

        /// <remarks/>
        custTbls,

        /// <remarks/>
        custWatermarks,

        /// <remarks/>
        custAutoTxt,

        /// <remarks/>
        custTxtBox,

        /// <remarks/>
        custPgNumT,

        /// <remarks/>
        custPgNumB,

        /// <remarks/>
        custPgNumMargins,

        /// <remarks/>
        custTblOfContents,

        /// <remarks/>
        custBib,

        /// <remarks/>
        custom1,

        /// <remarks/>
        custom2,

        /// <remarks/>
        custom3,

        /// <remarks/>
        custom4,

        /// <remarks/>
        custom5,
    }

    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocPartName
    {

        private string valField;

        private ST_OnOff decoratedField;

        private bool decoratedFieldSpecified;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocPartTypes
    {

        private List<CT_DocPartType> itemsField;

        private ST_OnOff allField;

        private bool allFieldSpecified;

        public CT_DocPartTypes()
        {
            this.itemsField = new List<CT_DocPartType>();
        }

        [System.Xml.Serialization.XmlElementAttribute("type", Order = 0)]
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocPartType
    {

        private ST_DocPartType valField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_DocPartType
    {

        /// <remarks/>
        none,

        /// <remarks/>
        normal,

        /// <remarks/>
        autoExp,

        /// <remarks/>
        toolbar,

        /// <remarks/>
        speller,

        /// <remarks/>
        formFld,

        /// <remarks/>
        bbPlcHdr,
    }

    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType11
    {

        /// <remarks/>
        behaviors,

        /// <remarks/>
        category,

        /// <remarks/>
        description,

        /// <remarks/>
        guid,

        /// <remarks/>
        name,

        /// <remarks/>
        style,

        /// <remarks/>
        types,
    }

    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocGrid
    {

        private ST_DocGrid typeField;

        private bool typeFieldSpecified;

        private string linePitchField;

        private string charSpaceField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
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


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_DocGrid
    {

        /// <remarks/>
        @default,

        /// <remarks/>
        lines,

        /// <remarks/>
        linesAndChars,

        /// <remarks/>
        snapToChars,
    }

    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocVars
    {

        private List<CT_DocVar> docVarField;

        public CT_DocVars()
        {
            this.docVarField = new List<CT_DocVar>();
        }

        [System.Xml.Serialization.XmlElementAttribute("docVar", Order = 0)]
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
