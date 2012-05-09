
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using NPOI.OpenXmlFormats.Shared;
using System.Xml.Serialization;
using System.Xml.Schema;

namespace NPOI.OpenXmlFormats.Wordprocessing
{

    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    //[System.Xml.Serialization.XmlRoot("hdr", Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = false)]
    public class CT_HdrFtr
    {

        private List<object> itemsField;

        private List<ItemsChoiceType8> itemsElementNameField;

        public CT_HdrFtr()
        {
            this.itemsElementNameField = new List<ItemsChoiceType8>();
            this.itemsField = new List<object>();
        }

        [System.Xml.Serialization.XmlElement("oMath", typeof(CT_OMath), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("oMathPara", typeof(CT_OMathPara), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("altChunk", typeof(CT_AltChunk), Order = 0)]
        [System.Xml.Serialization.XmlElement("bookmarkEnd", typeof(CT_MarkupRange), Order = 0)]
        [System.Xml.Serialization.XmlElement("bookmarkStart", typeof(CT_Bookmark), Order = 0)]
        [System.Xml.Serialization.XmlElement("commentRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [System.Xml.Serialization.XmlElement("commentRangeStart", typeof(CT_MarkupRange), Order = 0)]
        [System.Xml.Serialization.XmlElement("customXml", typeof(CT_CustomXmlBlock), Order = 0)]
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
        [System.Xml.Serialization.XmlElement("p", typeof(CT_P), Order = 0)]
        [System.Xml.Serialization.XmlElement("permEnd", typeof(CT_Perm), Order = 0)]
        [System.Xml.Serialization.XmlElement("permStart", typeof(CT_PermStart), Order = 0)]
        [System.Xml.Serialization.XmlElement("proofErr", typeof(CT_ProofErr), Order = 0)]
        [System.Xml.Serialization.XmlElement("sdt", typeof(CT_SdtBlock), Order = 0)]
        [System.Xml.Serialization.XmlElement("tbl", typeof(CT_Tbl), Order = 0)]
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
        public ItemsChoiceType8[] ItemsElementName
        {
            get
            {
                return this.itemsElementNameField.ToArray();
            }
            set
            {
                if (value == null || value.Length == 0)
                    this.itemsElementNameField = new List<ItemsChoiceType8>();
                else
                    this.itemsElementNameField = new List<ItemsChoiceType8>(value);
            }
        }

        public CT_Tbl GetTblArray(int i)
        {
            return GetObjectArray<CT_Tbl>(i, ItemsChoiceType8.tbl);
        }

        public IList<CT_Tbl> GetTblList()
        {
            return GetObjectList<CT_Tbl>(ItemsChoiceType8.tbl);
        }

        public CT_P AddNewP()
        {
            return AddNewObject<CT_P>(ItemsChoiceType8.p);
        }

        public void SetPArray(int i, CT_P cT_P)
        {
            SetObject<CT_P>(ItemsChoiceType8.p, i, cT_P);
        }
        #region Generic methods for object operation

        private List<T> GetObjectList<T>(ItemsChoiceType8 type) where T : class
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
        private int SizeOfArray(ItemsChoiceType8 type)
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
        private T GetObjectArray<T>(int p, ItemsChoiceType8 type) where T : class
        {
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                if (pos < 0 || pos >= this.itemsField.Count)
                    return null;
                return itemsField[pos] as T;
            }
        }
        private T AddNewObject<T>(ItemsChoiceType8 type) where T : class, new()
        {
            T t = new T();
            lock (this)
            {
                this.itemsElementNameField.Add(type);
                this.itemsField.Add(t);
            }
            return t;
        }
        private void SetObject<T>(ItemsChoiceType8 type, int p, T obj) where T : class
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
        private int GetObjectIndex(ItemsChoiceType8 type, int p)
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
        private void RemoveObject(ItemsChoiceType8 type, int p)
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
    [System.Xml.Serialization.XmlRoot("ftr", Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = false)]
    public class CT_Ftr : CT_HdrFtr
    {

    }
    [System.Xml.Serialization.XmlRoot("hdr", Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = false)]
    public class CT_Hdr : CT_HdrFtr
    {

    }
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType8
    {

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:oMath")]
        oMath,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:oMathPara")]
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

    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_HdrFtrRef : CT_Rel
    {

        private ST_HdrFtr typeField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_HdrFtr type
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


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_HdrFtr
    {

    
        [XmlEnum("even")]
        EVEN,

        [XmlEnum("default")]
        DEFAULT,

       [XmlEnum("first")]
        FIRST,
    }


    [System.Xml.Serialization.XmlIncludeAttribute(typeof(CT_FtnDocProps))]

    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FtnProps
    {

        private CT_FtnPos posField;

        private CT_NumFmt numFmtField;

        private CT_DecimalNumber numStartField;

        private CT_NumRestart numRestartField;

        public CT_FtnProps()
        {
            this.numRestartField = new CT_NumRestart();
            this.numStartField = new CT_DecimalNumber();
            this.numFmtField = new CT_NumFmt();
            this.posField = new CT_FtnPos();
        }

        [System.Xml.Serialization.XmlElement(Order = 0)]
        public CT_FtnPos pos
        {
            get
            {
                return this.posField;
            }
            set
            {
                this.posField = value;
            }
        }

        [System.Xml.Serialization.XmlElement(Order = 1)]
        public CT_NumFmt numFmt
        {
            get
            {
                return this.numFmtField;
            }
            set
            {
                this.numFmtField = value;
            }
        }

        [System.Xml.Serialization.XmlElement(Order = 2)]
        public CT_DecimalNumber numStart
        {
            get
            {
                return this.numStartField;
            }
            set
            {
                this.numStartField = value;
            }
        }

        [System.Xml.Serialization.XmlElement(Order = 3)]
        public CT_NumRestart numRestart
        {
            get
            {
                return this.numRestartField;
            }
            set
            {
                this.numRestartField = value;
            }
        }
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FtnPos
    {

        private ST_FtnPos valField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_FtnPos val
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
    public enum ST_FtnPos
    {

    
        pageBottom,

    
        beneathText,

    
        sectEnd,

    
        docEnd,
    }

    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot("footnotes", Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = false)]
    public class CT_Footnotes
    {

        private List<CT_FtnEdn> footnoteField;

        public CT_Footnotes()
        {
            this.footnoteField = new List<CT_FtnEdn>();
        }

        [System.Xml.Serialization.XmlElement("footnote", Order = 0)]
        public List<CT_FtnEdn> footnote
        {
            get
            {
                return this.footnoteField;
            }
            set
            {
                this.footnoteField = value;
            }
        }
        public CT_FtnEdn AddNewFootnote()
        {
            CT_FtnEdn f = new CT_FtnEdn();
            footnoteField.Add(f);
            return f;
        }
        [XmlIgnore]
        public IEnumerable<CT_FtnEdn> FootnoteList { get { return footnoteField.AsReadOnly(); } private set{} }
    }


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot("endnote", Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FtnEdn
    {

        private List<object> itemsField = new List<object>();

        private List<ItemsChoiceType7> itemsElementNameField = new List<ItemsChoiceType7>();

        private ST_FtnEdn typeField;

        private bool typeFieldSpecified;

        private string idField = string.Empty;

        [System.Xml.Serialization.XmlElement("oMath", typeof(CT_OMath), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("oMathPara", typeof(CT_OMathPara), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 0)]
        [System.Xml.Serialization.XmlElement("altChunk", typeof(CT_AltChunk), Order = 0)]
        [System.Xml.Serialization.XmlElement("bookmarkEnd", typeof(CT_MarkupRange), Order = 0)]
        [System.Xml.Serialization.XmlElement("bookmarkStart", typeof(CT_Bookmark), Order = 0)]
        [System.Xml.Serialization.XmlElement("commentRangeEnd", typeof(CT_MarkupRange), Order = 0)]
        [System.Xml.Serialization.XmlElement("commentRangeStart", typeof(CT_MarkupRange), Order = 0)]
        [System.Xml.Serialization.XmlElement("customXml", typeof(CT_CustomXmlBlock), Order = 0)]
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
        [System.Xml.Serialization.XmlElement("p", typeof(CT_P), Order = 0)]
        [System.Xml.Serialization.XmlElement("permEnd", typeof(CT_Perm), Order = 0)]
        [System.Xml.Serialization.XmlElement("permStart", typeof(CT_PermStart), Order = 0)]
        [System.Xml.Serialization.XmlElement("proofErr", typeof(CT_ProofErr), Order = 0)]
        [System.Xml.Serialization.XmlElement("sdt", typeof(CT_SdtBlock), Order = 0)]
        [System.Xml.Serialization.XmlElement("tbl", typeof(CT_Tbl), Order = 0)]
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
        public ItemsChoiceType7[] ItemsElementName
        {
            get
            {
                return this.itemsElementNameField.ToArray();
            }
            set
            {
                if (value == null || value.Length == 0)
                    this.itemsElementNameField = new List<ItemsChoiceType7>();
                else
                    this.itemsElementNameField = new List<ItemsChoiceType7>(value);
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_FtnEdn type
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

       // correct namespace, but no need for ns prefix: [XmlAttribute(Form = XmlSchemaForm.Qualified, DataType = "integer", Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
        [XmlAttribute(DataType = "integer")]
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
        public void Set(CT_FtnEdn note)
        {
            this.idField = note.idField;
            this.itemsElementNameField = note.itemsElementNameField;
            this.itemsField = note.itemsField;
            this.typeField = note.typeField;
        }
        private List<T> GetObjectList<T>(ItemsChoiceType7 type) where T : class
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
        private int SizeOfArray(ItemsChoiceType7 type)
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
        private T GetObjectArray<T>(int p, ItemsChoiceType7 type) where T : class
        {
            lock (this)
            {
                int pos = 0;
                for (int i = 0; i < itemsElementNameField.Count; i++)
                {
                    if (itemsElementNameField[i] == type)
                    {
                        if (pos == p)
                            return itemsField[i] as T;
                        else
                            pos++;
                    }
                }
                return null;
            }
        }
        private T AddNewObject<T>(ItemsChoiceType7 type) where T : class, new()
        {
            T t = new T();
            lock (this)
            {
                this.itemsElementNameField.Add(type);
                this.itemsField.Add(t);
            }
            return t;
        }
        private T AddNewObject<T>(ItemsChoiceType7 type, T t) where T : class, new()
        {
            lock (this)
            {
                this.itemsElementNameField.Add(type);
                this.itemsField.Add(t);
            }
            return t;
        }
        public IList<CT_P> GetPList()
        {
            return GetObjectList<CT_P>(ItemsChoiceType7.p);
        }

        public IList<CT_Tbl> GetTblList()
        {
            return GetObjectList<CT_Tbl>(ItemsChoiceType7.tbl);
        }

        public CT_Tbl GetTblArray(int i)
        {
            return GetObjectArray<CT_Tbl>(i, ItemsChoiceType7.tbl);
        }
        
        public CT_Tbl AddNewTbl()
        {
            return AddNewObject<CT_Tbl>(ItemsChoiceType7.tbl);
        }

        public CT_P AddNewP()
        {
            return AddNewObject<CT_P>(ItemsChoiceType7.p);
        }

        public CT_P AddNewP(CT_P paragraph)
        {
            return AddNewObject<CT_P>(ItemsChoiceType7.p, paragraph);
        }
    }


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType7
    {

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:oMath")]
        oMath,

    
        [System.Xml.Serialization.XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/math:oMathPara")]
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


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_FtnEdn
    {

    
        normal,

    
        separator,

    
        continuationSeparator,

    
        continuationNotice,
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot("endnotes", Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = false)]
    public class CT_Endnotes
    {

        private List<CT_FtnEdn> endnoteField = new List<CT_FtnEdn>();

        [System.Xml.Serialization.XmlElement("endnote", Order = 0)]
        public List<CT_FtnEdn> endnote
        {
            get
            {
                return this.endnoteField;
            }
            set
            {
                this.endnoteField = value;
            }
        }
    }

    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FtnDocProps : CT_FtnProps
    {

        private List<CT_FtnEdnSepRef> footnoteField;

        public CT_FtnDocProps()
        {
            this.footnoteField = new List<CT_FtnEdnSepRef>();
        }

        [System.Xml.Serialization.XmlElement("footnote", Order = 0)]
        public List<CT_FtnEdnSepRef> footnote
        {
            get
            {
                return this.footnoteField;
            }
            set
            {
                this.footnoteField = value;
            }
        }
    }


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public class CT_FtnEdnSepRef
    {
        private string idField;

        //[System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        [XmlAttribute(DataType = "integer")]
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

    [System.Xml.Serialization.XmlIncludeAttribute(typeof(CT_EdnDocProps))]

    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public class CT_EdnProps
    {

        private CT_EdnPos posField;

        private CT_NumFmt numFmtField;

        private CT_DecimalNumber numStartField;

        private CT_NumRestart numRestartField;

        public CT_EdnProps()
        {
            this.numRestartField = new CT_NumRestart();
            this.numStartField = new CT_DecimalNumber();
            this.numFmtField = new CT_NumFmt();
            this.posField = new CT_EdnPos();
        }

        [System.Xml.Serialization.XmlElement(Order = 0)]
        public CT_EdnPos pos
        {
            get
            {
                return this.posField;
            }
            set
            {
                this.posField = value;
            }
        }

        [System.Xml.Serialization.XmlElement(Order = 1)]
        public CT_NumFmt numFmt
        {
            get
            {
                return this.numFmtField;
            }
            set
            {
                this.numFmtField = value;
            }
        }

        [System.Xml.Serialization.XmlElement(Order = 2)]
        public CT_DecimalNumber numStart
        {
            get
            {
                return this.numStartField;
            }
            set
            {
                this.numStartField = value;
            }
        }

        [System.Xml.Serialization.XmlElement(Order = 3)]
        public CT_NumRestart numRestart
        {
            get
            {
                return this.numRestartField;
            }
            set
            {
                this.numRestartField = value;
            }
        }
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_EdnPos
    {

        private ST_EdnPos valField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_EdnPos val
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
    public enum ST_EdnPos
    {

    
        sectEnd,

    
        docEnd,
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_EdnDocProps : CT_EdnProps
    {

        private List<CT_FtnEdnSepRef> endnoteField;

        public CT_EdnDocProps()
        {
            this.endnoteField = new List<CT_FtnEdnSepRef>();
        }

        [System.Xml.Serialization.XmlElement("endnote", Order = 0)]
        public List<CT_FtnEdnSepRef> endnote
        {
            get
            {
                return this.endnoteField;
            }
            set
            {
                this.endnoteField = value;
            }
        }
    }

    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FtnEdnRef
    {

        private ST_OnOff customMarkFollowsField;

        private bool customMarkFollowsFieldSpecified;

        private string idField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff customMarkFollows
        {
            get
            {
                return this.customMarkFollowsField;
            }
            set
            {
                this.customMarkFollowsField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool customMarkFollowsSpecified
        {
            get
            {
                return this.customMarkFollowsFieldSpecified;
            }
            set
            {
                this.customMarkFollowsFieldSpecified = value;
            }
        }

        // TODO is the following correct?
        //[XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
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
}