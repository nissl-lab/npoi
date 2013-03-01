using System;
using System.Collections.Generic;
using System.Xml.Serialization;
using NPOI.OpenXmlFormats.Shared;


namespace NPOI.OpenXmlFormats.Wordprocessing
{


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Tbl
    {
        //EG_RangeMarkupElements
        private List<object> itemsField;

        private List<ItemsChoiceType30> itemsElementNameField;

        private CT_TblPr tblPrField;

        private CT_TblGrid tblGridField;

        private List<object> items1Field;

        private List<Items1ChoiceType> items1ElementNameField;

        public CT_Tbl()
        {
            this.items1ElementNameField = new List<Items1ChoiceType>();
            this.items1Field = new List<object>();
            //this.tblGridField = new CT_TblGrid();
            //this.tblPrField = new CT_TblPr();
            this.itemsElementNameField = new List<ItemsChoiceType30>();
            this.itemsField = new List<object>();
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

        [XmlElement("ItemsElementName", Order = 1)]
        [XmlIgnore]
        public ItemsChoiceType30[] ItemsElementName
        {
            get
            {
                return this.itemsElementNameField.ToArray();
            }
            set
            {
                if (value != null && value.Length != 0)
                    this.itemsElementNameField = new List<ItemsChoiceType30>(value);
                else
                    this.itemsElementNameField = new List<ItemsChoiceType30>();
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
        public object[] Items1
        {
            get
            {
                return this.items1Field.ToArray();
            }
            set
            {
                if (value != null && value.Length != 0)
                    this.items1Field = new List<object>(value);
                else
                    this.items1Field = new List<object>();
            }
        }

        [XmlElement("Items1ElementName", Order = 5)]
        [XmlIgnore]
        public Items1ChoiceType[] Items1ElementName
        {
            get
            {
                return this.items1ElementNameField.ToArray();
            }
            set
            {
                if (value != null && value.Length != 0)
                    this.items1ElementNameField = new List<Items1ChoiceType>(value);
                else
                    this.items1ElementNameField = new List<Items1ChoiceType>();
            }
        }

        public void Set(CT_Tbl table)
        {
            this.items1ElementNameField = new List<Items1ChoiceType>(table.Items1ElementName);
            this.items1Field = new List<object>(table.items1Field);
            this.itemsElementNameField = new List<ItemsChoiceType30>(table.itemsElementNameField);
            this.itemsField = new List<object>(table.itemsField);
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
        
        public IList<CT_Row> GetTrList()
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
            this.tblGridField = new List<CT_TblGridCol>();
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
            this.tblGridChangeField = new CT_TblGridChange();
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
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TblOverlap
    {

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

        private CT_TblPrBase tblPrField;

        public CT_TblPrChange()
        {
            this.tblPrField = new CT_TblPrBase();
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
            //this.tblLookField = new CT_ShortHexNumber();
            //this.tblCellMarField = new CT_TblCellMar();
            //this.tblLayoutField = new CT_TblLayoutType();
            //this.shdField = new CT_Shd();
            //this.tblBordersField = new CT_TblBorders();
            //this.tblIndField = new CT_TblWidth();
            //this.tblCellSpacingField = new CT_TblWidth();
            //this.jcField = new CT_Jc();
            //this.tblWField = new CT_TblWidth();
            //this.tblStyleColBandSizeField = new CT_DecimalNumber();
            //this.tblStyleRowBandSizeField = new CT_DecimalNumber();
            //this.bidiVisualField = new CT_OnOff();
            //this.tblOverlapField = new CT_TblOverlap();
            //this.tblpPrField = new CT_TblPPr();
            //this.tblStyleField = new CT_String();
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

        private bool valFieldSpecified;

        private ST_HeightRule hRuleField;

        private bool hRuleFieldSpecified;

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

        [XmlIgnore]
        public bool hRuleSpecified
        {
            get
            {
                return this.hRuleFieldSpecified;
            }
            set
            {
                this.hRuleFieldSpecified = value;
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
            this.tblPrExField = new CT_TblPrExBase();
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

        public CT_TblPrExBase()
        {
            this.tblLookField = new CT_ShortHexNumber();
            this.tblCellMarField = new CT_TblCellMar();
            this.tblLayoutField = new CT_TblLayoutType();
            this.shdField = new CT_Shd();
            this.tblBordersField = new CT_TblBorders();
            this.tblIndField = new CT_TblWidth();
            this.tblCellSpacingField = new CT_TblWidth();
            this.jcField = new CT_Jc();
            this.tblWField = new CT_TblWidth();
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

        public CT_TblPrEx()
        {
            this.tblPrExChangeField = new CT_TblPrExChange();
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
            this.insideVField = new CT_Border();
            this.insideHField = new CT_Border();
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

        public CT_TblCellMar()
        {
            this.rightField = new CT_TblWidth();
            this.bottomField = new CT_TblWidth();
            this.leftField = new CT_TblWidth();
            this.topField = new CT_TblWidth();
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
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TrPrChange : CT_TrackChange
    {

        private CT_TrPrBase trPrField;

        public CT_TrPrChange()
        {
            this.trPrField = new CT_TrPrBase();
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

        private List<object> itemsField;

        private List<ItemsChoiceType2> itemsElementNameField;

        public CT_TrPrBase()
        {
            this.itemsElementNameField = new List<ItemsChoiceType2>();
            this.itemsField = new List<object>();
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
        public ItemsChoiceType2[] ItemsElementName
        {
            get
            {
                return this.itemsElementNameField.ToArray();
            }
            set
            {
                if (value == null || value.Length == 0)
                    this.itemsElementNameField = new List<ItemsChoiceType2>();
                else
                    this.itemsElementNameField = new List<ItemsChoiceType2>(value);
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

        private List<object> itemsField;

        private List<ItemsChoiceTableCellType> itemsElementNameField;

        public CT_Tc()
        {
            this.itemsElementNameField = new List<ItemsChoiceTableCellType>();
            this.itemsField = new List<object>();
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

        [XmlElement("ItemsElementName", Order = 2)]
        [XmlIgnore]
        public ItemsChoiceTableCellType[] ItemsElementName
        {
            get
            {
                return this.itemsElementNameField.ToArray();
            }
            set
            {
                if (value == null || value.Length == 0)
                    this.itemsElementNameField = new List<ItemsChoiceTableCellType>();
                else
                    this.itemsElementNameField = new List<ItemsChoiceTableCellType>(value);
            }
        }
        [XmlIgnore]
        public CT_Tbl Table
        {
            get
            {
                if (_tableRow == null)
                    return null;
                return _tableRow.Table; 
            }
        }
        private CT_Row _tableRow;
        [XmlIgnore]
        public CT_Row TableRow
        {
            get { return _tableRow; }
            set { _tableRow = value; }
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
            this.trPrChangeField = new CT_TrPrChange();
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

        private CT_TcPrInner tcPrField;

        public CT_TcPrChange()
        {
            this.tcPrField = new CT_TcPrInner();
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
            this.cellMergeField = new CT_CellMergeTrackChange();
            this.cellDelField = new CT_TrackChange();
            this.cellInsField = new CT_TrackChange();
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

        public CT_TcBorders()
        {
            this.tr2blField = new CT_Border();
            this.tl2brField = new CT_Border();
            this.insideVField = new CT_Border();
            this.insideHField = new CT_Border();
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
            this.rightField = new CT_TblWidth();
            this.bottomField = new CT_TblWidth();
            this.leftField = new CT_TblWidth();
            this.topField = new CT_TblWidth();
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

        public CT_TcPr()
        {
            this.tcPrChangeField = new CT_TcPrChange();
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
            this.hideMarkField = new CT_OnOff();
            this.vAlignField = new CT_VerticalJc();
            this.tcFitTextField = new CT_OnOff();
            this.textDirectionField = new CT_TextDirection();
            this.tcMarField = new CT_TcMar();
            this.noWrapField = new CT_OnOff();
            this.shdField = new CT_Shd();
            this.tcBordersField = new CT_TcBorders();
            this.vMergeField = new CT_VMerge();
            this.hMergeField = new CT_HMerge();
            this.gridSpanField = new CT_DecimalNumber();
            this.tcWField = new CT_TblWidth();
            this.cnfStyleField = new CT_Cnf();
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
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_HMerge
    {

        private ST_Merge valField;

        private bool valFieldSpecified;

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

        private ST_Merge valField;

        private bool valFieldSpecified;

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

#region Table Row
    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Row
    {

        private CT_TblPrEx tblPrExField;

        private CT_TrPr trPrField;

        private List<object> itemsField;

        private List<ItemsChoiceTableRowType> itemsElementNameField;

        private byte[] rsidRPrField;

        private byte[] rsidRField;

        private byte[] rsidDelField;

        private byte[] rsidTrField;

        public CT_Row()
        {
            this.itemsElementNameField = new List<ItemsChoiceTableRowType>();
            this.itemsField = new List<object>();
            //this.trPrField = new CT_TrPr();
            //this.tblPrExField = new CT_TblPrEx();
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

        [XmlElement("ItemsElementName", Order = 3)]
        [XmlIgnore]
        public ItemsChoiceTableRowType[] ItemsElementName
        {
            get
            {
                return this.itemsElementNameField.ToArray();
            }
            set
            {
                if (value != null && value.Length != 0)
                    this.itemsElementNameField = new List<ItemsChoiceTableRowType>(value);
                else
                    this.itemsElementNameField = new List<ItemsChoiceTableRowType>();
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
            return this.trPrField.Items.Length > 0;
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