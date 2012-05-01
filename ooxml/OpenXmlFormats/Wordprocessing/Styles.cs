using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using NPOI.OpenXmlFormats.Shared;

namespace NPOI.OpenXmlFormats.Wordprocessing
{

    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute("styles", Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = false)]
    public class CT_Styles
    {

        private CT_DocDefaults docDefaultsField;

        private CT_LatentStyles latentStylesField;

        private List<CT_Style> styleField;

        public CT_Styles()
        {
            this.styleField = new List<CT_Style>();
            this.latentStylesField = new CT_LatentStyles();
            this.docDefaultsField = new CT_DocDefaults();
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_DocDefaults docDefaults
        {
            get
            {
                return this.docDefaultsField;
            }
            set
            {
                this.docDefaultsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
        public CT_LatentStyles latentStyles
        {
            get
            {
                return this.latentStylesField;
            }
            set
            {
                this.latentStylesField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute("style", Order = 2)]
        public List<CT_Style> style
        {
            get
            {
                return this.styleField;
            }
            set
            {
                this.styleField = value;
            }
        }

        public IList<CT_Style> GetStyleList()
        {
            return style;
        }

        public void AddNewStyle()
        {
            CT_Style s = new CT_Style();
            styleField.Add(s);
        }

        public void SetStyleArray(int pos, CT_Style cT_Style)
        {
            lock (this)
            {
                this.styleField[pos] = cT_Style;
            }
        }

        public bool IsSetDocDefaults()
        {
            throw new NotImplementedException();
        }

        public CT_DocDefaults AddNewDocDefaults()
        {
            throw new NotImplementedException();
        }
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocDefaults
    {

        private CT_RPrDefault rPrDefaultField;

        private CT_PPrDefault pPrDefaultField;

        public CT_DocDefaults()
        {
            this.pPrDefaultField = new CT_PPrDefault();
            this.rPrDefaultField = new CT_RPrDefault();
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_RPrDefault rPrDefault
        {
            get
            {
                return this.rPrDefaultField;
            }
            set
            {
                this.rPrDefaultField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
        public CT_PPrDefault pPrDefault
        {
            get
            {
                return this.pPrDefaultField;
            }
            set
            {
                this.pPrDefaultField = value;
            }
        }

        public bool IsSetRPrDefault()
        {
            throw new NotImplementedException();
        }

        public CT_RPrDefault AddNewRPrDefault()
        {
            throw new NotImplementedException();
        }
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_RPrDefault
    {

        private CT_RPr rPrField;

        public CT_RPrDefault()
        {
            this.rPrField = new CT_RPr();
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
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

        public bool IsSetRPr()
        {
            throw new NotImplementedException();
        }

        public CT_RPr AddNewRPr()
        {
            throw new NotImplementedException();
        }
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_PPrDefault
    {

        private CT_PPr pPrField;

        public CT_PPrDefault()
        {
            this.pPrField = new CT_PPr();
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
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
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_LatentStyles
    {

        private List<CT_LsdException> lsdExceptionField;

        private ST_OnOff defLockedStateField;

        private bool defLockedStateFieldSpecified;

        private string defUIPriorityField;

        private ST_OnOff defSemiHiddenField;

        private bool defSemiHiddenFieldSpecified;

        private ST_OnOff defUnhideWhenUsedField;

        private bool defUnhideWhenUsedFieldSpecified;

        private ST_OnOff defQFormatField;

        private bool defQFormatFieldSpecified;

        private string countField;

        public CT_LatentStyles()
        {
            this.lsdExceptionField = new List<CT_LsdException>();
        }

        [System.Xml.Serialization.XmlElementAttribute("lsdException", Order = 0)]
        public List<CT_LsdException> lsdException
        {
            get
            {
                return this.lsdExceptionField;
            }
            set
            {
                this.lsdExceptionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff defLockedState
        {
            get
            {
                return this.defLockedStateField;
            }
            set
            {
                this.defLockedStateField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool defLockedStateSpecified
        {
            get
            {
                return this.defLockedStateFieldSpecified;
            }
            set
            {
                this.defLockedStateFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string defUIPriority
        {
            get
            {
                return this.defUIPriorityField;
            }
            set
            {
                this.defUIPriorityField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff defSemiHidden
        {
            get
            {
                return this.defSemiHiddenField;
            }
            set
            {
                this.defSemiHiddenField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool defSemiHiddenSpecified
        {
            get
            {
                return this.defSemiHiddenFieldSpecified;
            }
            set
            {
                this.defSemiHiddenFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff defUnhideWhenUsed
        {
            get
            {
                return this.defUnhideWhenUsedField;
            }
            set
            {
                this.defUnhideWhenUsedField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool defUnhideWhenUsedSpecified
        {
            get
            {
                return this.defUnhideWhenUsedFieldSpecified;
            }
            set
            {
                this.defUnhideWhenUsedFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff defQFormat
        {
            get
            {
                return this.defQFormatField;
            }
            set
            {
                this.defQFormatField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool defQFormatSpecified
        {
            get
            {
                return this.defQFormatFieldSpecified;
            }
            set
            {
                this.defQFormatFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }
    }

    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Style
    {

        private CT_String nameField;

        private CT_String aliasesField;

        private CT_String basedOnField;

        private CT_String nextField;

        private CT_String linkField;

        private CT_OnOff autoRedefineField;

        private CT_OnOff hiddenField;

        private CT_DecimalNumber uiPriorityField;

        private CT_OnOff semiHiddenField;

        private CT_OnOff unhideWhenUsedField;

        private CT_OnOff qFormatField;

        private CT_OnOff lockedField;

        private CT_OnOff personalField;

        private CT_OnOff personalComposeField;

        private CT_OnOff personalReplyField;

        private CT_LongHexNumber rsidField;

        private CT_PPr pPrField;

        private CT_RPr rPrField;

        private CT_TblPrBase tblPrField;

        private CT_TrPr trPrField;

        private CT_TcPr tcPrField;

        private List<CT_TblStylePr> tblStylePrField;

        private ST_StyleType typeField;

        private bool typeFieldSpecified;

        private string styleIdField;

        private ST_OnOff defaultField;

        private bool defaultFieldSpecified;

        private ST_OnOff customStyleField;

        private bool customStyleFieldSpecified;

        public CT_Style()
        {
            this.tblStylePrField = new List<CT_TblStylePr>();
            this.tcPrField = new CT_TcPr();
            this.trPrField = new CT_TrPr();
            this.tblPrField = new CT_TblPrBase();
            this.rPrField = new CT_RPr();
            this.pPrField = new CT_PPr();
            this.rsidField = new CT_LongHexNumber();
            this.personalReplyField = new CT_OnOff();
            this.personalComposeField = new CT_OnOff();
            this.personalField = new CT_OnOff();
            this.lockedField = new CT_OnOff();
            this.qFormatField = new CT_OnOff();
            this.unhideWhenUsedField = new CT_OnOff();
            this.semiHiddenField = new CT_OnOff();
            this.uiPriorityField = new CT_DecimalNumber();
            this.hiddenField = new CT_OnOff();
            this.autoRedefineField = new CT_OnOff();
            this.linkField = new CT_String();
            this.nextField = new CT_String();
            this.basedOnField = new CT_String();
            this.aliasesField = new CT_String();
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
        public CT_String aliases
        {
            get
            {
                return this.aliasesField;
            }
            set
            {
                this.aliasesField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 2)]
        public CT_String basedOn
        {
            get
            {
                return this.basedOnField;
            }
            set
            {
                this.basedOnField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 3)]
        public CT_String next
        {
            get
            {
                return this.nextField;
            }
            set
            {
                this.nextField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 4)]
        public CT_String link
        {
            get
            {
                return this.linkField;
            }
            set
            {
                this.linkField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 5)]
        public CT_OnOff autoRedefine
        {
            get
            {
                return this.autoRedefineField;
            }
            set
            {
                this.autoRedefineField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 6)]
        public CT_OnOff hidden
        {
            get
            {
                return this.hiddenField;
            }
            set
            {
                this.hiddenField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 7)]
        public CT_DecimalNumber uiPriority
        {
            get
            {
                return this.uiPriorityField;
            }
            set
            {
                this.uiPriorityField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 8)]
        public CT_OnOff semiHidden
        {
            get
            {
                return this.semiHiddenField;
            }
            set
            {
                this.semiHiddenField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 9)]
        public CT_OnOff unhideWhenUsed
        {
            get
            {
                return this.unhideWhenUsedField;
            }
            set
            {
                this.unhideWhenUsedField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 10)]
        public CT_OnOff qFormat
        {
            get
            {
                return this.qFormatField;
            }
            set
            {
                this.qFormatField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 11)]
        public CT_OnOff locked
        {
            get
            {
                return this.lockedField;
            }
            set
            {
                this.lockedField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 12)]
        public CT_OnOff personal
        {
            get
            {
                return this.personalField;
            }
            set
            {
                this.personalField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 13)]
        public CT_OnOff personalCompose
        {
            get
            {
                return this.personalComposeField;
            }
            set
            {
                this.personalComposeField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 14)]
        public CT_OnOff personalReply
        {
            get
            {
                return this.personalReplyField;
            }
            set
            {
                this.personalReplyField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 15)]
        public CT_LongHexNumber rsid
        {
            get
            {
                return this.rsidField;
            }
            set
            {
                this.rsidField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 16)]
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

        [System.Xml.Serialization.XmlElementAttribute(Order = 17)]
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

        [System.Xml.Serialization.XmlElementAttribute(Order = 18)]
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

        [System.Xml.Serialization.XmlElementAttribute(Order = 19)]
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

        [System.Xml.Serialization.XmlElementAttribute(Order = 20)]
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

        [System.Xml.Serialization.XmlElementAttribute("tblStylePr", Order = 21)]
        public List<CT_TblStylePr> tblStylePr
        {
            get
            {
                return this.tblStylePrField;
            }
            set
            {
                this.tblStylePrField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_StyleType type
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string styleId
        {
            get
            {
                return this.styleIdField;
            }
            set
            {
                this.styleIdField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff @default
        {
            get
            {
                return this.defaultField;
            }
            set
            {
                this.defaultField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool defaultSpecified
        {
            get
            {
                return this.defaultFieldSpecified;
            }
            set
            {
                this.defaultFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff customStyle
        {
            get
            {
                return this.customStyleField;
            }
            set
            {
                this.customStyleField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool customStyleSpecified
        {
            get
            {
                return this.customStyleFieldSpecified;
            }
            set
            {
                this.customStyleFieldSpecified = value;
            }
        }

        public bool IsSetName()
        {
            throw new NotImplementedException();
        }
    }

    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Shd
    {

        private ST_Shd valField;

        private string colorField;

        private ST_ThemeColor themeColorField;

        private bool themeColorFieldSpecified;

        private byte[] themeTintField;

        private byte[] themeShadeField;

        private string fillField;

        private ST_ThemeColor themeFillField;

        private bool themeFillFieldSpecified;

        private byte[] themeFillTintField;

        private byte[] themeFillShadeField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_Shd val
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
        public string color
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ThemeColor themeColor
        {
            get
            {
                return this.themeColorField;
            }
            set
            {
                this.themeColorField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool themeColorSpecified
        {
            get
            {
                return this.themeColorFieldSpecified;
            }
            set
            {
                this.themeColorFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] themeTint
        {
            get
            {
                return this.themeTintField;
            }
            set
            {
                this.themeTintField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] themeShade
        {
            get
            {
                return this.themeShadeField;
            }
            set
            {
                this.themeShadeField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string fill
        {
            get
            {
                return this.fillField;
            }
            set
            {
                this.fillField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ThemeColor themeFill
        {
            get
            {
                return this.themeFillField;
            }
            set
            {
                this.themeFillField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool themeFillSpecified
        {
            get
            {
                return this.themeFillFieldSpecified;
            }
            set
            {
                this.themeFillFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] themeFillTint
        {
            get
            {
                return this.themeFillTintField;
            }
            set
            {
                this.themeFillTintField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] themeFillShade
        {
            get
            {
                return this.themeFillShadeField;
            }
            set
            {
                this.themeFillShadeField = value;
            }
        }
    }


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_Shd
    {

        /// <remarks/>
        nil,

        /// <remarks/>
        clear,

        /// <remarks/>
        solid,

        /// <remarks/>
        horzStripe,

        /// <remarks/>
        vertStripe,

        /// <remarks/>
        reverseDiagStripe,

        /// <remarks/>
        diagStripe,

        /// <remarks/>
        horzCross,

        /// <remarks/>
        diagCross,

        /// <remarks/>
        thinHorzStripe,

        /// <remarks/>
        thinVertStripe,

        /// <remarks/>
        thinReverseDiagStripe,

        /// <remarks/>
        thinDiagStripe,

        /// <remarks/>
        thinHorzCross,

        /// <remarks/>
        thinDiagCross,

        /// <remarks/>
        pct5,

        /// <remarks/>
        pct10,

        /// <remarks/>
        pct12,

        /// <remarks/>
        pct15,

        /// <remarks/>
        pct20,

        /// <remarks/>
        pct25,

        /// <remarks/>
        pct30,

        /// <remarks/>
        pct35,

        /// <remarks/>
        pct37,

        /// <remarks/>
        pct40,

        /// <remarks/>
        pct45,

        /// <remarks/>
        pct50,

        /// <remarks/>
        pct55,

        /// <remarks/>
        pct60,

        /// <remarks/>
        pct62,

        /// <remarks/>
        pct65,

        /// <remarks/>
        pct70,

        /// <remarks/>
        pct75,

        /// <remarks/>
        pct80,

        /// <remarks/>
        pct85,

        /// <remarks/>
        pct87,

        /// <remarks/>
        pct90,

        /// <remarks/>
        pct95,
    }










    //==============


    /// <summary>
    /// Text Expansion/Compression Percentage
    /// </summary>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TextScale
    {

        private string valField;
        /// <summary>
        /// Text Expansion/Compression Value
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
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
    }

    /// <summary>
    /// Text Highlight Colors
    /// </summary>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Highlight
    {

        private ST_HighlightColor valField;
        /// <summary>
        /// Highlighting Color
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_HighlightColor val
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
    /// <summary>
    /// Color Value
    /// </summary>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Color
    {

        private string valField;

        private ST_ThemeColor themeColorField;

        private bool themeColorFieldSpecified;

        private byte[] themeTintField;

        private byte[] themeShadeField;
        /// <summary>
        /// Run Content Color
        /// </summary>
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
        /// <summary>
        /// Run Content Theme Color
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ThemeColor themeColor
        {
            get
            {
                return this.themeColorField;
            }
            set
            {
                this.themeColorField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool themeColorSpecified
        {
            get
            {
                return this.themeColorFieldSpecified;
            }
            set
            {
                this.themeColorFieldSpecified = value;
            }
        }
        /// <summary>
        /// Run Content Theme Color Tint
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] themeTint
        {
            get
            {
                return this.themeTintField;
            }
            set
            {
                this.themeTintField = value;
            }
        }
        /// <summary>
        /// Run Content Theme Color Shade
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] themeShade
        {
            get
            {
                return this.themeShadeField;
            }
            set
            {
                this.themeShadeField = value;
            }
        }
    }

    /// <summary>
    /// Underline Style
    /// </summary>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Underline
    {

        private ST_Underline valField;

        private bool valFieldSpecified;

        private string colorField;

        private ST_ThemeColor themeColorField;

        private bool themeColorFieldSpecified;

        private byte[] themeTintField;

        private byte[] themeShadeField;
        /// <summary>
        /// Underline Style value
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_Underline val
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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
        /// <summary>
        /// Underline Color
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string color
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
        /// <summary>
        /// Underline Theme Color
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ThemeColor themeColor
        {
            get
            {
                return this.themeColorField;
            }
            set
            {
                this.themeColorField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool themeColorSpecified
        {
            get
            {
                return this.themeColorFieldSpecified;
            }
            set
            {
                this.themeColorFieldSpecified = value;
            }
        }
        /// <summary>
        /// Underline Theme Color Tint
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] themeTint
        {
            get
            {
                return this.themeTintField;
            }
            set
            {
                this.themeTintField = value;
            }
        }
        /// <summary>
        /// Underline Theme Color Shade
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] themeShade
        {
            get
            {
                return this.themeShadeField;
            }
            set
            {
                this.themeShadeField = value;
            }
        }
    }

    /// <summary>
    /// Underline Patterns
    /// </summary>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_Underline
    {

        /// <summary>
        /// Single Underline
        /// </summary>
        single,

        /// <summary>
        /// Underline Non-Space Characters Only
        /// </summary>
        words,

        /// <summary>
        /// Double Underline
        /// </summary>
        @double,

        /// <summary>
        /// Thick Underline
        /// </summary>
        thick,

        /// <summary>
        /// Dotted Underline
        /// </summary>
        dotted,

        /// <summary>
        /// Thick Dotted Underline
        /// </summary>
        dottedHeavy,

        /// <summary>
        /// Dashed Underline
        /// </summary>
        dash,

        /// <summary>
        /// Thick Dashed Underline
        /// </summary>
        dashedHeavy,

        /// <summary>
        /// Long Dashed Underline
        /// </summary>
        dashLong,

        /// <summary>
        /// Thick Long Dashed Underline
        /// </summary>
        dashLongHeavy,

        /// <summary>
        /// Dash-Dot Underline
        /// </summary>
        dotDash,

        /// <summary>
        /// Thick Dash-Dot Underline
        /// </summary>
        dashDotHeavy,

        /// <summary>
        /// Dash-Dot-Dot Underline
        /// </summary>
        dotDotDash,

        /// <summary>
        /// Thick Dash-Dot-Dot Underline
        /// </summary>
        dashDotDotHeavy,

        /// <summary>
        /// Wave Underline
        /// </summary>
        wave,

        /// <summary>
        /// Heavy Wave Underline
        /// </summary>
        wavyHeavy,

        /// <summary>
        /// Double Wave Underline
        /// </summary>
        wavyDouble,

        /// <summary>
        /// No Underline
        /// </summary>
        none,
    }

    /// <summary>
    /// Animated Text Effects
    /// </summary>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TextEffect
    {

        private ST_TextEffect valField;
        /// <summary>
        /// Animated Text Effect Type
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_TextEffect val
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
    public enum ST_TextEffect
    {

        /// <summary>
        /// Blinking Background Animation
        /// </summary>
        blinkBackground,

        /// <summary>
        /// Colored Lights Animation
        /// </summary>
        lights,

        /// <summary>
        /// Black Dashed Line Animation
        /// </summary>
        antsBlack,

        /// <summary>
        /// Marching Red Ants
        /// </summary>
        antsRed,

        /// <summary>
        /// Shimmer Animation
        /// </summary>
        shimmer,

        /// <summary>
        /// Sparkling Lights Animation
        /// </summary>
        sparkle,

        /// <summary>
        /// No Animation
        /// </summary>
        none,
    }

    /// <summary>
    /// Border Style
    /// </summary>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Border
    {

        private ST_Border valField;

        private string colorField;

        private ST_ThemeColor themeColorField;

        private bool themeColorFieldSpecified;

        private byte[] themeTintField;

        private byte[] themeShadeField;

        private ulong szField;

        private bool szFieldSpecified;

        private ulong spaceField;

        private bool spaceFieldSpecified;

        private ST_OnOff shadowField;

        private bool shadowFieldSpecified;

        private ST_OnOff frameField;

        private bool frameFieldSpecified;
        /// <summary>
        /// Border Style
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_Border val
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
        /// <summary>
        /// Border Color
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string color
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
        /// <summary>
        /// Border Theme Color
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ThemeColor themeColor
        {
            get
            {
                return this.themeColorField;
            }
            set
            {
                this.themeColorField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool themeColorSpecified
        {
            get
            {
                return this.themeColorFieldSpecified;
            }
            set
            {
                this.themeColorFieldSpecified = value;
            }
        }
        /// <summary>
        /// Border Theme Color Tint
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] themeTint
        {
            get
            {
                return this.themeTintField;
            }
            set
            {
                this.themeTintField = value;
            }
        }
        /// <summary>
        /// Border Theme Color Shade
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] themeShade
        {
            get
            {
                return this.themeShadeField;
            }
            set
            {
                this.themeShadeField = value;
            }
        }
        /// <summary>
        /// Border Width
        /// </summary>
        /// ST_EighthPointMeasure
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong sz
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool szSpecified
        {
            get
            {
                return this.szFieldSpecified;
            }
            set
            {
                this.szFieldSpecified = value;
            }
        }
        /// <summary>
        /// Border Spacing Measurement
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong space
        {
            get
            {
                return this.spaceField;
            }
            set
            {
                this.spaceField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool spaceSpecified
        {
            get
            {
                return this.spaceFieldSpecified;
            }
            set
            {
                this.spaceFieldSpecified = value;
            }
        }
        /// <summary>
        /// Border Shadow
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff shadow
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool shadowSpecified
        {
            get
            {
                return this.shadowFieldSpecified;
            }
            set
            {
                this.shadowFieldSpecified = value;
            }
        }
        /// <summary>
        /// Create Frame Effect
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff frame
        {
            get
            {
                return this.frameField;
            }
            set
            {
                this.frameField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool frameSpecified
        {
            get
            {
                return this.frameFieldSpecified;
            }
            set
            {
                this.frameFieldSpecified = value;
            }
        }
    }

    /// <summary>
    /// Border Styles
    /// </summary>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_Border
    {

        /// <summary>
        /// No Border
        /// </summary>
        nil,

        /// <summary>
        /// No Border
        /// </summary>
        none,

        /// <summary>
        /// Single Line Border
        /// </summary>
        single,

        /// <summary>
        /// Single Line Border
        /// </summary>
        thick,

        /// <summary>
        /// Double Line Border
        /// </summary>
        @double,

        /// <summary>
        /// Dotted Line Border
        /// </summary>
        dotted,

        /// <summary>
        /// Dashed Line Border
        /// </summary>
        dashed,

        /// <summary>
        /// Dot Dash Line Border
        /// </summary>
        dotDash,

        /// <summary>
        /// Dot Dot Dash Line Border
        /// </summary>
        dotDotDash,

        /// <summary>
        /// Triple Line Border
        /// </summary>
        triple,

        /// <summary>
        /// Thin, Thick Line Border<
        /// </summary>
        thinThickSmallGap,

        /// <summary>
        /// Thick, Thin Line Border
        /// </summary>
        thickThinSmallGap,

        /// <summary>
        /// Thin, Thick, Thin Line Border
        /// </summary>
        thinThickThinSmallGap,

        /// <summary>
        /// Thin, Thick Line Border
        /// </summary>
        thinThickMediumGap,

        /// <summary>
        /// Thick, Thin Line Border
        /// </summary>
        thickThinMediumGap,

        /// <summary>
        /// Thin, Thick, Thin Line Border
        /// </summary>
        thinThickThinMediumGap,

        /// <summary>
        /// Thin, Thick Line Border
        /// </summary>
        thinThickLargeGap,

        /// <summary>
        /// Thick, Thin Line Border
        /// </summary>
        thickThinLargeGap,

        /// <summary>
        /// Thin, Thick, Thin Line Border
        /// </summary>
        thinThickThinLargeGap,

        /// <summary>
        /// Wavy Line Border
        /// </summary>
        wave,

        /// <summary>
        /// Double Wave Line Border
        /// </summary>
        doubleWave,

        /// <summary>
        /// Dashed Line Border
        /// </summary>
        dashSmallGap,

        /// <summary>
        /// Dash Dot Strokes Line Border
        /// </summary>
        dashDotStroked,

        /// <summary>
        /// 3D Embossed Line Border
        /// </summary>
        threeDEmboss,

        /// <summary>
        /// 3D Engraved Line Border
        /// </summary>
        threeDEngrave,

        /// <summary>
        /// Outset Line Border
        /// </summary>
        outset,

        /// <summary>
        /// Inset Line Border
        /// </summary>
        inset,

        /// <summary>
        /// Apples Art Border
        /// </summary>
        apples,

        /// <summary>
        /// Arched Scallops Art Border
        /// </summary>
        archedScallops,

        /// <summary>
        /// Baby Pacifier Art Border
        /// </summary>
        babyPacifier,

        /// <summary>
        /// Baby Rattle Art Border
        /// </summary>
        babyRattle,

        /// <summary>
        /// Three Color Balloons Art Border
        /// </summary>
        balloons3Colors,

        /// <summary>
        /// Hot Air Balloons Art Border
        /// </summary>
        balloonsHotAir,

        /// <summary>
        /// Black Dash Art Border
        /// </summary>
        basicBlackDashes,

        /// <summary>
        /// Black Dot Art Border
        /// </summary>
        basicBlackDots,

        /// <summary>
        /// Black Square Art Border
        /// </summary>
        basicBlackSquares,

        /// <summary>
        /// Thin Line Art Border
        /// </summary>
        basicThinLines,

        /// <summary>
        /// White Dash Art Border
        /// </summary>
        basicWhiteDashes,

        /// <summary>
        /// White Dot Art Border
        /// </summary>
        basicWhiteDots,

        /// <summary>
        /// White Square Art Border
        /// </summary>
        basicWhiteSquares,

        /// <summary>
        /// Wide Inline Art Border
        /// </summary>
        basicWideInline,

        /// <summary>
        /// Wide Midline Art Border
        /// </summary>
        basicWideMidline,

        /// <summary>
        /// 
        /// </summary>
        basicWideOutline,

        /// <summary>
        /// Wide Outline Art Border
        /// </summary>
        bats,

        /// <summary>
        /// Bats Art Border
        /// </summary>
        birds,

        /// <summary>
        /// Birds Art Border
        /// </summary>
        birdsFlight,

        /// <summary>
        /// Cabin Art Border
        /// </summary>
        cabins,

        /// <summary>
        /// Cake Art Border
        /// </summary>
        cakeSlice,

        /// <summary>
        /// Candy Corn Art Border
        /// </summary>
        candyCorn,

        /// <summary>
        /// Knot Work Art Border
        /// </summary>
        celticKnotwork,

        /// <summary>
        /// Certificate Banner Art Border
        /// </summary>
        certificateBanner,

        /// <summary>
        /// Chain Link Art Border
        /// </summary>
        chainLink,

        /// <summary>
        /// Champagne Bottle Art Border
        /// </summary>
        champagneBottle,

        /// <summary>
        /// Black and White Bar Art Border
        /// </summary>
        checkedBarBlack,

        /// <summary>
        /// Color Checked Bar Art Border
        /// </summary>
        checkedBarColor,

        /// <summary>
        /// Checkerboard Art Border
        /// </summary>
        checkered,

        /// <summary>
        /// Christmas Tree Art Border
        /// </summary>
        christmasTree,

        /// <summary>
        /// Circles And Lines Art Border
        /// </summary>
        circlesLines,

        /// <summary>
        /// Circles and Rectangles Art Border
        /// </summary>
        circlesRectangles,

        /// <summary>
        /// Wave Art Border
        /// </summary>
        classicalWave,

        /// <summary>
        /// Clocks Art Border
        /// </summary>
        clocks,

        /// <summary>
        /// Compass Art Border
        /// </summary>
        compass,

        /// <summary>
        /// Confetti Art Border
        /// </summary>
        confetti,

        /// <summary>
        /// Confetti Art Border
        /// </summary>
        confettiGrays,

        /// <summary>
        /// Confetti Art Border
        /// </summary>
        confettiOutline,

        /// <summary>
        /// Confetti Streamers Art Border
        /// </summary>
        confettiStreamers,

        /// <summary>
        /// Confetti Art Border
        /// </summary>
        confettiWhite,

        /// <summary>
        /// Corner Triangle Art Border
        /// </summary>
        cornerTriangles,

        /// <summary>
        /// Dashed Line Art Border
        /// </summary>
        couponCutoutDashes,

        /// <summary>
        /// Dotted Line Art Border
        /// </summary>
        couponCutoutDots,

        /// <summary>
        /// Maze Art Border
        /// </summary>
        crazyMaze,

        /// <summary>
        /// Butterfly Art Border
        /// </summary>
        creaturesButterfly,

        /// <summary>
        /// Fish Art Border
        /// </summary>
        creaturesFish,

        /// <summary>
        /// Insects Art Border
        /// </summary>
        creaturesInsects,

        /// <summary>
        /// Ladybug Art Border
        /// </summary>
        creaturesLadyBug,

        /// <summary>
        /// Cross-stitch Art Border
        /// </summary>
        crossStitch,

        /// <summary>
        /// Cupid Art Border
        /// </summary>
        cup,

        /// <summary>
        /// Archway Art Border
        /// </summary>
        decoArch,

        /// <summary>
        /// Color Archway Art Border
        /// </summary>
        decoArchColor,

        /// <summary>
        /// Blocks Art Border
        /// </summary>
        decoBlocks,

        /// <summary>
        /// Gray Diamond Art Border
        /// </summary>
        diamondsGray,

        /// <summary>
        /// Double D Art Border
        /// </summary>
        doubleD,

        /// <summary>
        /// Diamond Art Border
        /// </summary>
        doubleDiamonds,

        /// <summary>
        /// Earth Art Border
        /// </summary>
        earth1,

        /// <summary>
        /// Earth Art Border
        /// </summary>
        earth2,

        /// <summary>
        /// Shadowed Square Art Border
        /// </summary>
        eclipsingSquares1,

        /// <summary>
        /// Shadowed Square Art Border
        /// </summary>
        eclipsingSquares2,

        /// <summary>
        /// Painted Egg Art Border
        /// </summary>
        eggsBlack,

        /// <summary>
        /// Fans Art Border
        /// </summary>
        fans,

        /// <summary>
        /// Film Reel Art Border
        /// </summary>
        film,

        /// <summary>
        /// Firecracker Art Border
        /// </summary>
        firecrackers,

        /// <summary>
        /// Flowers Art Border
        /// </summary>
        flowersBlockPrint,

        /// <summary>
        /// Daisy Art Border
        /// </summary>
        flowersDaisies,

        /// <summary>
        /// Flowers Art Border
        /// </summary>
        flowersModern1,

        /// <summary>
        /// Flowers Art Border
        /// </summary>
        flowersModern2,

        /// <summary>
        /// Pansy Art Border
        /// </summary>
        flowersPansy,

        /// <summary>
        /// Red Rose Art Border
        /// </summary>
        flowersRedRose,

        /// <summary>
        /// Roses Art Border
        /// </summary>
        flowersRoses,

        /// <summary>
        /// Flowers in a Teacup Art Border
        /// </summary>
        flowersTeacup,

        /// <summary>
        /// Small Flower Art Border
        /// </summary>
        flowersTiny,

        /// <summary>
        /// Gems Art Border
        /// </summary>
        gems,

        /// <summary>
        /// Gingerbread Man Art Border
        /// </summary>
        gingerbreadMan,

        /// <summary>
        /// Triangle Gradient Art Border
        /// </summary>
        gradient,

        /// <summary>
        /// Handmade Art Border
        /// </summary>
        handmade1,

        /// <summary>
        /// Handmade Art Border
        /// </summary>
        handmade2,

        /// <summary>
        /// Heart-Shaped Balloon Art Border
        /// </summary>
        heartBalloon,

        /// <summary>
        /// Gray Heart Art Border
        /// </summary>
        heartGray,

        /// <summary>
        /// Hearts Art Border
        /// </summary>
        hearts,

        /// <summary>
        /// Pattern Art Border
        /// </summary>
        heebieJeebies,

        /// <summary>
        /// Holly Art Border
        /// </summary>
        holly,

        /// <summary>
        /// House Art Border
        /// </summary>
        houseFunky,

        /// <summary>
        /// Circular Art Border
        /// </summary>
        hypnotic,

        /// <summary>
        /// Ice Cream Cone Art Border
        /// </summary>
        iceCreamCones,

        /// <summary>
        /// Light Bulb Art Border
        /// </summary>
        lightBulb,

        /// <summary>
        /// Light Bulb Art Border
        /// </summary>
        lightning1,

        /// <summary>
        /// Light Bulb Art Border
        /// </summary>
        lightning2,

        /// <summary>
        /// Map Pins Art Border
        /// </summary>
        mapPins,

        /// <summary>
        /// Maple Leaf Art Border
        /// </summary>
        mapleLeaf,
        
        /// <summary>
        /// Muffin Art Border
        /// </summary>
        mapleMuffins,

        /// <summary>
        /// Marquee Art Border
        /// </summary>
        marquee,

        /// <summary>
        /// Marquee Art Border
        /// </summary>
        marqueeToothed,

        /// <summary>
        /// Moon Art Border
        /// </summary>
        moons,

        /// <summary>
        /// Mosaic Art Border
        /// </summary>
        mosaic,

        /// <summary>
        /// Musical Note Art Border
        /// </summary>
        musicNotes,

        /// <summary>
        /// Patterned Art Border
        /// </summary>
        northwest,

        /// <summary>
        /// Oval Art Border
        /// </summary>
        ovals,

        /// <summary>
        /// Package Art Border
        /// </summary>
        packages,

        /// <summary>
        /// Black Palm Tree Art Border
        /// </summary>
        palmsBlack,

        /// <summary>
        /// Color Palm Tree Art Border
        /// </summary>
        palmsColor,

        /// <summary>
        /// Paper Clip Art Border
        /// </summary>
        paperClips,

        /// <summary>
        /// Papyrus Art Border
        /// </summary>
        papyrus,

        /// <summary>
        /// Party Favor Art Border
        /// </summary>
        partyFavor,

        /// <summary>
        /// Party Glass Art Border
        /// </summary>
        partyGlass,

        /// <summary>
        /// Pencils Art Border
        /// </summary>
        pencils,

        /// <summary>
        /// Character Art Border
        /// </summary>
        people,

        /// <summary>
        /// Waving Character Border
        /// </summary>
        peopleWaving,

        /// <summary>
        /// Character With Hat Art Border
        /// </summary>
        peopleHats,

        /// <summary>
        /// Poinsettia Art Border
        /// </summary>
        poinsettias,

        /// <summary>
        /// Postage Stamp Art Border
        /// </summary>
        postageStamp,

        /// <summary>
        /// Pumpkin Art Border
        /// </summary>
        pumpkin1,

        /// <summary>
        /// Push Pin Art Border
        /// </summary>
        pushPinNote2,

        /// <summary>
        /// Push Pin Art Border
        /// </summary>
        pushPinNote1,

        /// <summary>
        /// Pyramid Art Border
        /// </summary>
        pyramids,

        /// <summary>
        /// Pyramid Art Border
        /// </summary>
        pyramidsAbove,

        /// <summary>
        /// Quadrants Art Border
        /// </summary>
        quadrants,

        /// <summary>
        /// Rings Art Border
        /// </summary>
        rings,

        /// <summary>
        /// Safari Art Border
        /// </summary>
        safari,

        /// <summary>
        /// Saw tooth Art Border
        /// </summary>
        sawtooth,

        /// <summary>
        /// Gray Saw tooth Art Border
        /// </summary>
        sawtoothGray,

        /// <summary>
        /// Scared Cat Art Border
        /// </summary>
        scaredCat,

        /// <summary>
        /// Umbrella Art Border
        /// </summary>
        seattle,

        /// <summary>
        /// Shadowed Squares Art Border
        /// </summary>
        shadowedSquares,

        /// <summary>
        /// Shark Tooth Art Border
        /// </summary>
        sharksTeeth,

        /// <summary>
        /// Bird Tracks Art Border
        /// </summary>
        shorebirdTracks,

        /// <summary>
        /// Rocket Art Border
        /// </summary>
        skyrocket,

        /// <summary>
        /// Snowflake Art Border
        /// </summary>
        snowflakeFancy,

        /// <summary>
        /// Snowflake Art Border
        /// </summary>
        snowflakes,

        /// <summary>
        /// Sombrero Art Border
        /// </summary>
        sombrero,

        /// <summary>
        /// Southwest-themed Art Border
        /// </summary>
        southwest,

        /// <summary>
        /// Stars Art Border
        /// </summary>
        stars,

        /// <summary>
        /// Stars On Top Art Border
        /// </summary>
        starsTop,

        /// <summary>
        /// 3-D Stars Art Border
        /// </summary>
        stars3d,

        /// <summary>
        /// Stars Art Border
        /// </summary>
        starsBlack,

        /// <summary>
        /// Stars With Shadows Art Border
        /// </summary>
        starsShadowed,

        /// <summary>
        /// Sun Art Border
        /// </summary>
        sun,

        /// <summary>
        /// Whirligig Art Border
        /// </summary>
        swirligig,

        /// <summary>
        /// Torn Paper Art Border
        /// </summary>
        tornPaper,

        /// <summary>
        /// Black Torn Paper Art Border
        /// </summary>
        tornPaperBlack,

        /// <summary>
        /// Tree Art Border
        /// </summary>
        trees,

        /// <summary>
        /// Triangle Art Border
        /// </summary>
        triangleParty,

        /// <summary>
        /// Triangles Art Border
        /// </summary>
        triangles,

        /// <summary>
        /// Tribal Art Border One
        /// </summary>
        tribal1,

        /// <summary>
        /// Tribal Art Border Two
        /// </summary>
        tribal2,

        /// <summary>
        /// Tribal Art Border Three
        /// </summary>
        tribal3,

        /// <summary>
        /// Tribal Art Border Four
        /// </summary>
        tribal4,

        /// <summary>
        /// Tribal Art Border Five
        /// </summary>
        tribal5,

        /// <summary>
        /// Tribal Art Border Six
        /// </summary>
        tribal6,

        /// <summary>
        /// Twisted Lines Art Border
        /// </summary>
        twistedLines1,

        /// <summary>
        /// Twisted Lines Art Border
        /// </summary>
        twistedLines2,

        /// <summary>
        /// Vine Art Border
        /// </summary>
        vine,

        /// <summary>
        /// Wavy Line Art Border
        /// </summary>
        waveline,

        /// <summary>
        /// Weaving Angles Art Border
        /// </summary>
        weavingAngles,

        /// <summary>
        /// Weaving Braid Art Border
        /// </summary>
        weavingBraid,

        /// <summary>
        /// Weaving Ribbon Art Border
        /// </summary>
        weavingRibbon,

        /// <summary>
        /// Weaving Strips Art Border
        /// </summary>
        weavingStrips,

        /// <summary>
        /// White Flowers Art Border
        /// </summary>
        whiteFlowers,

        /// <summary>
        /// Woodwork Art Border
        /// </summary>
        woodwork,

        /// <summary>
        /// Crisscross Art Border
        /// </summary>
        xIllusions,

        /// <summary>
        /// Triangle Art Border
        /// </summary>
        zanyTriangles,

        /// <summary>
        /// Zigzag Art Border
        /// </summary>
        zigZag,

        /// <summary>
        /// Zigzag stitch
        /// </summary>
        zigZagStitch,
    }
}