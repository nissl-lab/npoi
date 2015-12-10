using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Serialization;
using NPOI.OpenXmlFormats.Shared;
using System.IO;
using NPOI.OpenXml4Net.Util;
using System.Collections;
using NPOI.OpenXmlFormats.Dml.WordProcessing;

namespace NPOI.OpenXmlFormats.Wordprocessing
{


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_R
    {

        private CT_RPr rPrField;

        private ArrayList itemsField;

        private List<RunItemsChoiceType> itemsElementNameField;

        private byte[] rsidRPrField;

        private byte[] rsidDelField;

        private byte[] rsidRField;

        public CT_R()
        {
            this.itemsElementNameField = new List<RunItemsChoiceType>();
            this.itemsField = new ArrayList();
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
        public ArrayList Items
        {
            get
            {
                return this.itemsField;
            }
            set
            {
               this.itemsField = new ArrayList();
            }
        }

        [XmlElement("ItemsElementName", Order = 2)]
        [XmlIgnore]
        public List<RunItemsChoiceType> ItemsElementName
        {
            get
            {
                return this.itemsElementNameField;
            }
            set
            {
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
        public void RemoveT(int pos)
        {
            RemoveObject(RunItemsChoiceType.t, pos);
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

        public static CT_R Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_R ctObj = new CT_R();
            ctObj.rsidRPr = XmlHelper.ReadBytes(node.Attributes["w:rsidRPr"]);
            ctObj.rsidDel = XmlHelper.ReadBytes(node.Attributes["w:rsidDel"]);
            ctObj.rsidR = XmlHelper.ReadBytes(node.Attributes["w:rsidR"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "rPr")
                    ctObj.rPr = CT_RPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "instrText")
                {
                    ctObj.Items.Add(CT_Text.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.instrText);
                }
                else if (childNode.LocalName == "object")
                {
                    ctObj.Items.Add(CT_Object.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.@object);
                }
                else if (childNode.LocalName == "pgNum")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.pgNum);
                }
                else if (childNode.LocalName == "annotationRef")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.annotationRef);
                }
                else if (childNode.LocalName == "br")
                {
                    ctObj.Items.Add(CT_Br.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.br);
                }
                else if (childNode.LocalName == "commentReference")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.commentReference);
                }
                else if (childNode.LocalName == "continuationSeparator")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.continuationSeparator);
                }
                else if (childNode.LocalName == "cr")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.cr);
                }
                else if (childNode.LocalName == "dayLong")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.dayLong);
                }
                else if (childNode.LocalName == "dayShort")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.dayShort);
                }
                else if (childNode.LocalName == "delInstrText")
                {
                    ctObj.Items.Add(CT_Text.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.delInstrText);
                }
                else if (childNode.LocalName == "delText")
                {
                    ctObj.Items.Add(CT_Text.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.delText);
                }
                else if (childNode.LocalName == "drawing")
                {
                    ctObj.Items.Add(CT_Drawing.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.drawing);
                }
                else if (childNode.LocalName == "endnoteRef")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.endnoteRef);
                }
                else if (childNode.LocalName == "endnoteReference")
                {
                    ctObj.Items.Add(CT_FtnEdnRef.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.endnoteReference);
                }
                else if (childNode.LocalName == "fldChar")
                {
                    ctObj.Items.Add(CT_FldChar.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.fldChar);
                }
                else if (childNode.LocalName == "footnoteRef")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.footnoteRef);
                }
                else if (childNode.LocalName == "footnoteReference")
                {
                    ctObj.Items.Add(CT_FtnEdnRef.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.footnoteReference);
                }
                else if (childNode.LocalName == "lastRenderedPageBreak")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.lastRenderedPageBreak);
                }
                else if (childNode.LocalName == "monthLong")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.monthLong);
                }
                else if (childNode.LocalName == "monthShort")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.monthShort);
                }
                else if (childNode.LocalName == "noBreakHyphen")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.noBreakHyphen);
                }
                else if (childNode.LocalName == "yearLong")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.yearLong);
                }
                else if (childNode.LocalName == "pict")
                {
                    ctObj.Items.Add(CT_Picture.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.pict);
                }
                else if (childNode.LocalName == "ptab")
                {
                    ctObj.Items.Add(CT_PTab.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.ptab);
                }
                else if (childNode.LocalName == "ruby")
                {
                    ctObj.Items.Add(CT_Ruby.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.ruby);
                }
                else if (childNode.LocalName == "separator")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.separator);
                }
                else if (childNode.LocalName == "softHyphen")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.softHyphen);
                }
                else if (childNode.LocalName == "sym")
                {
                    ctObj.Items.Add(CT_Sym.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.sym);
                }
                else if (childNode.LocalName == "t")
                {
                    ctObj.Items.Add(CT_Text.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.t);
                }
                else if (childNode.LocalName == "tab")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.tab);
                }
                else if (childNode.LocalName == "yearShort")
                {
                    ctObj.Items.Add(new CT_Empty());
                    ctObj.ItemsElementName.Add(RunItemsChoiceType.yearShort);
                }
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:rsidR", this.rsidR);
            XmlHelper.WriteAttribute(sw, "w:rsidRPr", this.rsidRPr);
            XmlHelper.WriteAttribute(sw, "w:rsidDel", this.rsidDel);
            sw.Write(">");
            if (this.rPr != null)
                this.rPr.Write(sw, "rPr");
            int i = 0;
            foreach (object o in this.Items)
            {
                if ((o is CT_Text) && this.ItemsElementName[i] == RunItemsChoiceType.instrText)
                    ((CT_Text)o).Write(sw, "instrText");
                else if (o is CT_Object)
                    ((CT_Object)o).Write(sw, "object");
                else if (o is CT_Empty && this.ItemsElementName[i] == RunItemsChoiceType.pgNum)
                    sw.Write("<w:pgNum/>");
                else if (o is CT_Empty && this.ItemsElementName[i] == RunItemsChoiceType.annotationRef)
                    sw.Write("<w:annotationRef/>");
                else if (o is CT_Br)
                    ((CT_Br)o).Write(sw, "br");
                else if (o is CT_Markup)
                    ((CT_Markup)o).Write(sw, "w:commentReference");
                else if (o is CT_Empty && this.ItemsElementName[i] == RunItemsChoiceType.continuationSeparator)
                    sw.Write("<w:continuationSeparator/>");
                else if (o is CT_Empty && this.ItemsElementName[i] == RunItemsChoiceType.cr)
                    sw.Write("<w:cr/>");
                else if (o is CT_Empty && this.ItemsElementName[i] == RunItemsChoiceType.dayLong)
                    sw.Write("<w:dayLong/>");
                else if (o is CT_Empty && this.ItemsElementName[i] == RunItemsChoiceType.dayShort)
                    sw.Write("<w:dayShort/>");
                else if ((o is CT_Text) && this.ItemsElementName[i] == RunItemsChoiceType.delInstrText)
                    ((CT_Text)o).Write(sw, "delInstrText");
                else if (o is CT_Text && this.ItemsElementName[i] == RunItemsChoiceType.delText)
                    ((CT_Text)o).Write(sw, "delText");
                else if (o is CT_Drawing)
                    ((CT_Drawing)o).Write(sw, "drawing");
                else if (o is CT_Empty && this.ItemsElementName[i] == RunItemsChoiceType.endnoteRef)
                    sw.Write("<w:endnoteRef/>");
                else if (o is CT_FtnEdnRef)
                    ((CT_FtnEdnRef)o).Write(sw, "endnoteReference");
                else if (o is CT_FldChar)
                    ((CT_FldChar)o).Write(sw, "fldChar");
                else if (o is CT_Empty && this.ItemsElementName[i] == RunItemsChoiceType.footnoteRef)
                    sw.Write("<w:footnoteRef/>");
                else if (o is CT_FtnEdnRef)
                    ((CT_FtnEdnRef)o).Write(sw, "footnoteReference");
                else if (o is CT_Empty && this.ItemsElementName[i] == RunItemsChoiceType.lastRenderedPageBreak)
                    sw.Write("<w:lastRenderedPageBreak/>");
                else if (o is CT_Empty && this.ItemsElementName[i] == RunItemsChoiceType.monthLong)
                    sw.Write("<w:monthLong/>");
                else if (o is CT_Empty && this.ItemsElementName[i] == RunItemsChoiceType.monthShort)
                    sw.Write("<w:monthShort/>");
                else if (o is CT_Empty && this.ItemsElementName[i] == RunItemsChoiceType.noBreakHyphen)
                    sw.Write("<w:noBreakHyphen/>");
                else if (o is CT_Empty && this.ItemsElementName[i] == RunItemsChoiceType.yearLong)
                    sw.Write("<w:yearLong/>");
                else if (o is CT_Picture)
                    ((CT_Picture)o).Write(sw, "pict");
                else if (o is CT_PTab)
                    ((CT_PTab)o).Write(sw, "ptab");
                else if (o is CT_Ruby)
                    ((CT_Ruby)o).Write(sw, "ruby");
                else if (o is CT_Empty && this.ItemsElementName[i] == RunItemsChoiceType.separator)
                    sw.Write("<w:separator/>");
                else if (o is CT_Empty && this.ItemsElementName[i] == RunItemsChoiceType.softHyphen)
                    sw.Write("<w:softHyphen/>");
                else if (o is CT_Sym)
                    ((CT_Sym)o).Write(sw, "sym");
                else if ((o is CT_Text) && this.ItemsElementName[i] == RunItemsChoiceType.t)
                    ((CT_Text)o).Write(sw, "t");
                else if (o is CT_Empty&& this.ItemsElementName[i]== RunItemsChoiceType.tab)
                    sw.Write("<w:tab/>");
                else if (o is CT_Empty&& this.ItemsElementName[i]== RunItemsChoiceType.yearShort)
                    sw.Write("<w:yearShort/>");
                i++;
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }


        public int SizeOfTabArray()
        {
            return SizeOfArray(RunItemsChoiceType.tab);
        }

        public IList<CT_Empty> GetTabList()
        {
            return GetObjectList<CT_Empty>(RunItemsChoiceType.tab);
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
        public static CT_RubyAlign Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_RubyAlign ctObj = new CT_RubyAlign();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_RubyAlign)Enum.Parse(typeof(ST_RubyAlign), node.Attributes["w:val"].Value);
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
            //this.dirtyField = new CT_OnOff();
            //this.lidField = new CT_Lang();
            //this.hpsBaseTextField = new CT_HpsMeasure();
            //this.hpsRaiseField = new CT_HpsMeasure();
            //this.hpsField = new CT_HpsMeasure();
            //this.rubyAlignField = new CT_RubyAlign();
        }
        public static CT_RubyPr Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_RubyPr ctObj = new CT_RubyPr();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "rubyAlign")
                    ctObj.rubyAlign = CT_RubyAlign.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "hps")
                    ctObj.hps = CT_HpsMeasure.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "hpsRaise")
                    ctObj.hpsRaise = CT_HpsMeasure.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "hpsBaseText")
                    ctObj.hpsBaseText = CT_HpsMeasure.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lid")
                    ctObj.lid = CT_Lang.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dirty")
                    ctObj.dirty = CT_OnOff.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.rubyAlign != null)
                this.rubyAlign.Write(sw, "rubyAlign");
            if (this.hps != null)
                this.hps.Write(sw, "hps");
            if (this.hpsRaise != null)
                this.hpsRaise.Write(sw, "hpsRaise");
            if (this.hpsBaseText != null)
                this.hpsBaseText.Write(sw, "hpsBaseText");
            if (this.lid != null)
                this.lid.Write(sw, "lid");
            if (this.dirty != null)
                this.dirty.Write(sw, "dirty");
            sw.Write(string.Format("</w:{0}>", nodeName));
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

        private ArrayList itemsField;

        private List<ItemsChoiceType16> itemsElementNameField;

        public CT_RubyContent()
        {
            this.itemsElementNameField = new List<ItemsChoiceType16>();
            this.itemsField = new ArrayList();
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
        public List<ItemsChoiceType16> ItemsElementName
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

        public static CT_RubyContent Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_RubyContent ctObj = new CT_RubyContent();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "oMath")
                {
                    ctObj.Items.Add(CT_OMath.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.oMath);
                }
                else if (childNode.LocalName == "permStart")
                {
                    ctObj.Items.Add(CT_PermStart.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.permStart);
                }
                else if (childNode.LocalName == "proofErr")
                {
                    ctObj.Items.Add(CT_ProofErr.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.proofErr);
                }
                else if (childNode.LocalName == "permEnd")
                {
                    ctObj.Items.Add(CT_Perm.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.permEnd);
                }
                else if (childNode.LocalName == "moveToRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.moveToRangeStart);
                }
                else if (childNode.LocalName == "moveToRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.moveToRangeEnd);
                }
                else if (childNode.LocalName == "moveTo")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.moveTo);
                }
                else if (childNode.LocalName == "moveFromRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.moveFromRangeStart);
                }
                else if (childNode.LocalName == "moveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.moveFromRangeEnd);
                }
                else if (childNode.LocalName == "moveFrom")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.moveFrom);
                }
                else if (childNode.LocalName == "customXmlInsRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.customXmlInsRangeEnd);
                }
                else if (childNode.LocalName == "customXmlInsRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.customXmlInsRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.customXmlMoveFromRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.customXmlMoveFromRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.customXmlMoveToRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.customXmlMoveToRangeStart);
                }
                else if (childNode.LocalName == "del")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.del);
                }
                else if (childNode.LocalName == "ins")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.ins);
                }
                else if (childNode.LocalName == "customXmlDelRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.customXmlDelRangeStart);
                }
                else if (childNode.LocalName == "commentRangeStart")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.commentRangeStart);
                }
                else if (childNode.LocalName == "commentRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.commentRangeEnd);
                }
                else if (childNode.LocalName == "bookmarkStart")
                {
                    ctObj.Items.Add(CT_Bookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.bookmarkStart);
                }
                else if (childNode.LocalName == "bookmarkEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.bookmarkEnd);
                }
                else if (childNode.LocalName == "oMathPara")
                {
                    ctObj.Items.Add(CT_OMathPara.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.oMathPara);
                }
                else if (childNode.LocalName == "customXmlDelRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.customXmlDelRangeEnd);
                }
                else if (childNode.LocalName == "r")
                {
                    ctObj.Items.Add(CT_R.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType16.r);
                }
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            int i = 0;
            foreach (object o in this.Items)
            {
                if (o is CT_OMath)
                    ((CT_OMath)o).Write(sw, "oMath");
                else if (o is CT_PermStart)
                    ((CT_PermStart)o).Write(sw, "permStart");
                else if (o is CT_ProofErr)
                    ((CT_ProofErr)o).Write(sw, "proofErr");
                else if (o is CT_Perm)
                    ((CT_Perm)o).Write(sw, "permEnd");
                else if (o is CT_MoveBookmark && this.itemsElementNameField[i] == ItemsChoiceType16.moveToRangeStart)
                    ((CT_MoveBookmark)o).Write(sw, "moveToRangeStart");
                else if (o is CT_MarkupRange && this.itemsElementNameField[i] == ItemsChoiceType16.moveToRangeEnd)
                    ((CT_MarkupRange)o).Write(sw, "moveToRangeEnd");
                else if (o is CT_RunTrackChange&& this.itemsElementNameField[i]== ItemsChoiceType16.moveTo)
                    ((CT_RunTrackChange)o).Write(sw, "moveTo");
                else if (o is CT_RunTrackChange && this.itemsElementNameField[i] == ItemsChoiceType16.moveFrom)
                    ((CT_RunTrackChange)o).Write(sw, "moveFrom");
                else if (o is CT_MoveBookmark && this.itemsElementNameField[i] == ItemsChoiceType16.moveFromRangeStart)
                    ((CT_MoveBookmark)o).Write(sw, "moveFromRangeStart");
                else if (o is CT_MarkupRange && this.itemsElementNameField[i] == ItemsChoiceType16.moveFromRangeEnd)
                    ((CT_MarkupRange)o).Write(sw, "moveFromRangeEnd");
                else if (o is CT_Markup && this.itemsElementNameField[i] == ItemsChoiceType16.customXmlInsRangeEnd)
                    ((CT_Markup)o).Write(sw, "customXmlInsRangeEnd");
                else if (o is CT_TrackChange && this.itemsElementNameField[i] == ItemsChoiceType16.customXmlInsRangeStart)
                    ((CT_TrackChange)o).Write(sw, "customXmlInsRangeStart");
                else if (o is CT_Markup && this.itemsElementNameField[i] == ItemsChoiceType16.customXmlMoveFromRangeEnd)
                    ((CT_Markup)o).Write(sw, "customXmlMoveFromRangeEnd");
                else if (o is CT_TrackChange && this.itemsElementNameField[i] == ItemsChoiceType16.customXmlMoveFromRangeStart)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveFromRangeStart");
                else if (o is CT_Markup && this.itemsElementNameField[i] == ItemsChoiceType16.customXmlMoveToRangeEnd)
                    ((CT_Markup)o).Write(sw, "customXmlMoveToRangeEnd");
                else if (o is CT_TrackChange && this.itemsElementNameField[i] == ItemsChoiceType16.customXmlMoveToRangeStart)
                    ((CT_TrackChange)o).Write(sw, "customXmlMoveToRangeStart");
                else if (o is CT_RunTrackChange && this.itemsElementNameField[i] == ItemsChoiceType16.del)
                    ((CT_RunTrackChange)o).Write(sw, "del");
                else if (o is CT_RunTrackChange && this.itemsElementNameField[i] == ItemsChoiceType16.ins)
                    ((CT_RunTrackChange)o).Write(sw, "ins");
                else if (o is CT_TrackChange && this.itemsElementNameField[i] == ItemsChoiceType16.customXmlDelRangeStart)
                    ((CT_TrackChange)o).Write(sw, "customXmlDelRangeStart");
                else if (o is CT_Markup && this.itemsElementNameField[i] == ItemsChoiceType16.customXmlDelRangeEnd)
                    ((CT_Markup)o).Write(sw, "customXmlDelRangeEnd");
                else if (o is CT_MarkupRange && this.itemsElementNameField[i] == ItemsChoiceType16.commentRangeStart)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeStart");
                else if (o is CT_MarkupRange && this.itemsElementNameField[i] == ItemsChoiceType16.commentRangeEnd)
                    ((CT_MarkupRange)o).Write(sw, "commentRangeEnd");
                else if (o is CT_Bookmark)
                    ((CT_Bookmark)o).Write(sw, "bookmarkStart");
                else if (o is CT_MarkupRange && this.itemsElementNameField[i] == ItemsChoiceType16.bookmarkEnd)
                    ((CT_MarkupRange)o).Write(sw, "bookmarkEnd");
                else if (o is CT_OMathPara)
                    ((CT_OMathPara)o).Write(sw, "oMathPara");
                else if (o is CT_R)
                    ((CT_R)o).Write(sw, "r");
                i++;
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
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
            //this.rubyBaseField = new CT_RubyContent();
            //this.rtField = new CT_RubyContent();
            //this.rubyPrField = new CT_RubyPr();
        }

        public static CT_Ruby Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Ruby ctObj = new CT_Ruby();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "rubyPr")
                    ctObj.rubyPr = CT_RubyPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "rt")
                    ctObj.rt = CT_RubyContent.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "rubyBase")
                    ctObj.rubyBase = CT_RubyContent.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.rubyPr != null)
                this.rubyPr.Write(sw, "rubyPr");
            if (this.rt != null)
                this.rt.Write(sw, "rt");
            if (this.rubyBase != null)
                this.rubyBase.Write(sw, "rubyBase");
            sw.Write(string.Format("</w:{0}>", nodeName));
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
            sw.Write("/>");
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
            sw.Write("/>");
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
        private ArrayList itemsField;

        private List<ItemsChoiceType6> itemsElementNameField;

        public CT_RunTrackChange()
        {
            this.itemsElementNameField = new List<ItemsChoiceType6>();
            this.itemsField = new ArrayList();
        }
        public new static CT_RunTrackChange Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_RunTrackChange ctObj = new CT_RunTrackChange();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "func")
                {
                    ctObj.Items.Add(CT_Func.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.func);
                }
                else if (childNode.LocalName == "sSubSup")
                {
                    ctObj.Items.Add(CT_SSubSup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.sSubSup);
                }
                else if (childNode.LocalName == "customXmlInsRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.customXmlInsRangeStart);
                }
                else if (childNode.LocalName == "acc")
                {
                    ctObj.Items.Add(CT_Acc.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.acc);
                }
                else if (childNode.LocalName == "bar")
                {
                    ctObj.Items.Add(CT_Bar.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.bar);
                }
                else if (childNode.LocalName == "borderBox")
                {
                    ctObj.Items.Add(CT_BorderBox.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.borderBox);
                }
                else if (childNode.LocalName == "box")
                {
                    ctObj.Items.Add(CT_Box.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.box);
                }
                else if (childNode.LocalName == "d")
                {
                    ctObj.Items.Add(CT_D.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.d);
                }
                else if (childNode.LocalName == "eqArr")
                {
                    ctObj.Items.Add(CT_EqArr.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.eqArr);
                }
                else if (childNode.LocalName == "f")
                {
                    ctObj.Items.Add(CT_F.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.f);
                }
                else if (childNode.LocalName == "groupChr")
                {
                    ctObj.Items.Add(CT_GroupChr.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.groupChr);
                }
                else if (childNode.LocalName == "limLow")
                {
                    ctObj.Items.Add(CT_LimLow.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.limLow);
                }
                else if (childNode.LocalName == "m")
                {
                    ctObj.Items.Add(CT_M.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.m);
                }
                else if (childNode.LocalName == "nary")
                {
                    ctObj.Items.Add(CT_Nary.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.nary);
                }
                else if (childNode.LocalName == "oMath")
                {
                    ctObj.Items.Add(CT_OMath.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.oMath);
                }
                else if (childNode.LocalName == "oMathPara")
                {
                    ctObj.Items.Add(CT_OMathPara.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.oMathPara);
                }
                else if (childNode.LocalName == "phant")
                {
                    ctObj.Items.Add(CT_Phant.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.phant);
                }
                else if (childNode.LocalName == "r")
                {
                    ctObj.Items.Add(CT_R.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.r);
                }
                else if (childNode.LocalName == "rad")
                {
                    ctObj.Items.Add(CT_Rad.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.rad);
                }
                else if (childNode.LocalName == "sPre")
                {
                    ctObj.Items.Add(CT_SPre.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.sPre);
                }
                else if (childNode.LocalName == "sSub")
                {
                    ctObj.Items.Add(CT_SSub.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.sSub);
                }
                else if (childNode.LocalName == "sSup")
                {
                    ctObj.Items.Add(CT_SSup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.sSup);
                }
                else if (childNode.LocalName == "bookmarkEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.bookmarkEnd);
                }
                else if (childNode.LocalName == "bookmarkStart")
                {
                    ctObj.Items.Add(CT_Bookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.bookmarkStart);
                }
                else if (childNode.LocalName == "commentRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.commentRangeEnd);
                }
                else if (childNode.LocalName == "commentRangeStart")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.commentRangeStart);
                }
                else if (childNode.LocalName == "customXml")
                {
                    ctObj.Items.Add(CT_CustomXmlRun.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.customXml);
                }
                else if (childNode.LocalName == "customXmlDelRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.customXmlDelRangeEnd);
                }
                else if (childNode.LocalName == "customXmlDelRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.customXmlDelRangeStart);
                }
                else if (childNode.LocalName == "customXmlInsRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.customXmlInsRangeEnd);
                }
                else if (childNode.LocalName == "limUpp")
                {
                    ctObj.Items.Add(CT_LimUpp.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.limUpp);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.customXmlMoveFromRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveFromRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.customXmlMoveFromRangeStart);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeEnd")
                {
                    ctObj.Items.Add(CT_Markup.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.customXmlMoveToRangeEnd);
                }
                else if (childNode.LocalName == "customXmlMoveToRangeStart")
                {
                    ctObj.Items.Add(CT_TrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.customXmlMoveToRangeStart);
                }
                else if (childNode.LocalName == "del")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.del);
                }
                else if (childNode.LocalName == "ins")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.ins);
                }
                else if (childNode.LocalName == "moveFrom")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.moveFrom);
                }
                else if (childNode.LocalName == "moveFromRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.moveFromRangeEnd);
                }
                else if (childNode.LocalName == "moveFromRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.moveFromRangeStart);
                }
                else if (childNode.LocalName == "moveTo")
                {
                    ctObj.Items.Add(CT_RunTrackChange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.moveTo);
                }
                else if (childNode.LocalName == "moveToRangeEnd")
                {
                    ctObj.Items.Add(CT_MarkupRange.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.moveToRangeEnd);
                }
                else if (childNode.LocalName == "moveToRangeStart")
                {
                    ctObj.Items.Add(CT_MoveBookmark.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.moveToRangeStart);
                }
                else if (childNode.LocalName == "permEnd")
                {
                    ctObj.Items.Add(CT_Perm.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.permEnd);
                }
                else if (childNode.LocalName == "permStart")
                {
                    ctObj.Items.Add(CT_PermStart.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.permStart);
                }
                else if (childNode.LocalName == "proofErr")
                {
                    ctObj.Items.Add(CT_ProofErr.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.proofErr);
                }
                else if (childNode.LocalName == "r")
                {
                    ctObj.Items.Add(CT_R.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.r);
                }
                else if (childNode.LocalName == "sdt")
                {
                    ctObj.Items.Add(CT_SdtRun.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.sdt);
                }
                else if (childNode.LocalName == "smartTag")
                {
                    ctObj.Items.Add(CT_SmartTagRun.Parse(childNode, namespaceManager));
                    ctObj.ItemsElementName.Add(ItemsChoiceType6.smartTag);
                }
            }
            return ctObj;
        }

        internal new void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            foreach (object o in this.Items)
            {
                if (o is CT_Func)
                    ((CT_Func)o).Write(sw, "func");
                else if (o is CT_SSubSup)
                    ((CT_SSubSup)o).Write(sw, "sSubSup");
                else if (o is CT_TrackChange)
                    ((CT_TrackChange)o).Write(sw, "customXmlInsRangeStart");
                else if (o is CT_Acc)
                    ((CT_Acc)o).Write(sw, "acc");
                else if (o is CT_Bar)
                    ((CT_Bar)o).Write(sw, "bar");
                else if (o is CT_BorderBox)
                    ((CT_BorderBox)o).Write(sw, "borderBox");
                else if (o is CT_Box)
                    ((CT_Box)o).Write(sw, "box");
                else if (o is CT_D)
                    ((CT_D)o).Write(sw, "d");
                else if (o is CT_EqArr)
                    ((CT_EqArr)o).Write(sw, "eqArr");
                else if (o is CT_F)
                    ((CT_F)o).Write(sw, "f");
                else if (o is CT_GroupChr)
                    ((CT_GroupChr)o).Write(sw, "groupChr");
                else if (o is CT_LimLow)
                    ((CT_LimLow)o).Write(sw, "limLow");
                else if (o is CT_M)
                    ((CT_M)o).Write(sw, "m");
                else if (o is CT_Nary)
                    ((CT_Nary)o).Write(sw, "nary");
                else if (o is CT_OMath)
                    ((CT_OMath)o).Write(sw, "oMath");
                else if (o is CT_OMathPara)
                    ((CT_OMathPara)o).Write(sw, "oMathPara");
                else if (o is CT_Phant)
                    ((CT_Phant)o).Write(sw, "phant");
                else if (o is CT_R)
                    ((CT_R)o).Write(sw, "r");
                else if (o is CT_Rad)
                    ((CT_Rad)o).Write(sw, "rad");
                else if (o is CT_SPre)
                    ((CT_SPre)o).Write(sw, "sPre");
                else if (o is CT_SSub)
                    ((CT_SSub)o).Write(sw, "sSub");
                else if (o is CT_SSup)
                    ((CT_SSup)o).Write(sw, "sSup");
                else if (o is CT_MarkupRange)
                    ((CT_MarkupRange)o).Write(sw, "bookmarkEnd");
                else if (o is CT_Bookmark)
                    ((CT_Bookmark)o).Write(sw, "bookmarkStart");
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
                else if (o is CT_LimUpp)
                    ((CT_LimUpp)o).Write(sw, "limUpp");
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
                else if (o is CT_R)
                    ((CT_R)o).Write(sw, "r");
                else if (o is CT_SdtRun)
                    ((CT_SdtRun)o).Write(sw, "sdt");
                else if (o is CT_SmartTagRun)
                    ((CT_SmartTagRun)o).Write(sw, "smartTag");
            }
            sw.Write(string.Format("</w:{0}", nodeName));
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
        public List<ItemsChoiceType6> ItemsElementName
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
            return GetObjectList<CT_R>(ItemsChoiceType6.r);
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