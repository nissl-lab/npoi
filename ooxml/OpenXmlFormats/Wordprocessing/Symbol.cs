using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Wordprocessing
{
    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Num
    {

        private CT_DecimalNumber abstractNumIdField;

        private List<CT_NumLvl> lvlOverrideField;

        private string numIdField;

        public CT_Num()
        {
            this.lvlOverrideField = new List<CT_NumLvl>();
            this.abstractNumIdField = new CT_DecimalNumber();
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_DecimalNumber abstractNumId
        {
            get
            {
                return this.abstractNumIdField;
            }
            set
            {
                this.abstractNumIdField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute("lvlOverride", Order = 1)]
        public List<CT_NumLvl> lvlOverride
        {
            get
            {
                return this.lvlOverrideField;
            }
            set
            {
                this.lvlOverrideField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string numId
        {
            get
            {
                return this.numIdField;
            }
            set
            {
                this.numIdField = value;
            }
        }
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_NumLvl
    {

        private CT_DecimalNumber startOverrideField;

        private CT_Lvl lvlField;

        private string ilvlField;

        public CT_NumLvl()
        {
            this.lvlField = new CT_Lvl();
            this.startOverrideField = new CT_DecimalNumber();
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_DecimalNumber startOverride
        {
            get
            {
                return this.startOverrideField;
            }
            set
            {
                this.startOverrideField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
        public CT_Lvl lvl
        {
            get
            {
                return this.lvlField;
            }
            set
            {
                this.lvlField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string ilvl
        {
            get
            {
                return this.ilvlField;
            }
            set
            {
                this.ilvlField = value;
            }
        }
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute("numbering", Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = false)]
    public class CT_Numbering
    {

        private List<CT_NumPicBullet> numPicBulletField;

        private List<CT_AbstractNum> abstractNumField;

        private List<CT_Num> numField;

        private CT_DecimalNumber numIdMacAtCleanupField;

        public CT_Numbering()
        {
            this.numIdMacAtCleanupField = new CT_DecimalNumber();
            this.numField = new List<CT_Num>();
            this.abstractNumField = new List<CT_AbstractNum>();
            this.numPicBulletField = new List<CT_NumPicBullet>();
        }

        [System.Xml.Serialization.XmlElementAttribute("numPicBullet", Order = 0)]
        public List<CT_NumPicBullet> numPicBullet
        {
            get
            {
                return this.numPicBulletField;
            }
            set
            {
                this.numPicBulletField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute("abstractNum", Order = 1)]
        public List<CT_AbstractNum> abstractNum
        {
            get
            {
                return this.abstractNumField;
            }
            set
            {
                this.abstractNumField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute("num", Order = 2)]
        public List<CT_Num> num
        {
            get
            {
                return this.numField;
            }
            set
            {
                this.numField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 3)]
        public CT_DecimalNumber numIdMacAtCleanup
        {
            get
            {
                return this.numIdMacAtCleanupField;
            }
            set
            {
                this.numIdMacAtCleanupField = value;
            }
        }
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_NumPicBullet
    {

        private CT_Picture pictField;

        private string numPicBulletIdField;

        public CT_NumPicBullet()
        {
            this.pictField = new CT_Picture();
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_Picture pict
        {
            get
            {
                return this.pictField;
            }
            set
            {
                this.pictField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string numPicBulletId
        {
            get
            {
                return this.numPicBulletIdField;
            }
            set
            {
                this.numPicBulletIdField = value;
            }
        }
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_NumFmt
    {

        private ST_NumberFormat valField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_NumberFormat val
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
    public enum ST_NumberFormat
    {

        /// <remarks/>
        [XmlEnum("decimal")]
        @decimal,

        /// <remarks/>
        upperRoman,

        /// <remarks/>
        lowerRoman,

        /// <remarks/>
        upperLetter,

        /// <remarks/>
        lowerLetter,

        /// <remarks/>
        ordinal,

        /// <remarks/>
        cardinalText,

        /// <remarks/>
        ordinalText,

        /// <remarks/>
        hex,

        /// <remarks/>
        chicago,

        /// <remarks/>
        ideographDigital,

        /// <remarks/>
        japaneseCounting,

        /// <remarks/>
        aiueo,

        /// <remarks/>
        iroha,

        /// <remarks/>
        decimalFullWidth,

        /// <remarks/>
        decimalHalfWidth,

        /// <remarks/>
        japaneseLegal,

        /// <remarks/>
        japaneseDigitalTenThousand,

        /// <remarks/>
        decimalEnclosedCircle,

        /// <remarks/>
        decimalFullWidth2,

        /// <remarks/>
        aiueoFullWidth,

        /// <remarks/>
        irohaFullWidth,

        /// <remarks/>
        decimalZero,

        /// <remarks/>
        bullet,

        /// <remarks/>
        ganada,

        /// <remarks/>
        chosung,

        /// <remarks/>
        decimalEnclosedFullstop,

        /// <remarks/>
        decimalEnclosedParen,

        /// <remarks/>
        decimalEnclosedCircleChinese,

        /// <remarks/>
        ideographEnclosedCircle,

        /// <remarks/>
        ideographTraditional,

        /// <remarks/>
        ideographZodiac,

        /// <remarks/>
        ideographZodiacTraditional,

        /// <remarks/>
        taiwaneseCounting,

        /// <remarks/>
        ideographLegalTraditional,

        /// <remarks/>
        taiwaneseCountingThousand,

        /// <remarks/>
        taiwaneseDigital,

        /// <remarks/>
        chineseCounting,

        /// <remarks/>
        chineseLegalSimplified,

        /// <remarks/>
        chineseCountingThousand,

        /// <remarks/>
        koreanDigital,

        /// <remarks/>
        koreanCounting,

        /// <remarks/>
        koreanLegal,

        /// <remarks/>
        koreanDigital2,

        /// <remarks/>
        vietnameseCounting,

        /// <remarks/>
        russianLower,

        /// <remarks/>
        russianUpper,

        /// <remarks/>
        none,

        /// <remarks/>
        numberInDash,

        /// <remarks/>
        hebrew1,

        /// <remarks/>
        hebrew2,

        /// <remarks/>
        arabicAlpha,

        /// <remarks/>
        arabicAbjad,

        /// <remarks/>
        hindiVowels,

        /// <remarks/>
        hindiConsonants,

        /// <remarks/>
        hindiNumbers,

        /// <remarks/>
        hindiCounting,

        /// <remarks/>
        thaiLetters,

        /// <remarks/>
        thaiNumbers,

        /// <remarks/>
        thaiCounting,
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_NumRestart
    {

        private ST_RestartNumber valField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_RestartNumber val
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
    public enum ST_RestartNumber
    {

        /// <remarks/>
        continuous,

        /// <remarks/>
        eachSect,

        /// <remarks/>
        eachPage,
    }
    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_NumPr
    {

        private CT_DecimalNumber ilvlField;

        private CT_DecimalNumber numIdField;

        private CT_TrackChangeNumbering numberingChangeField;

        private CT_TrackChange insField;

        public CT_NumPr()
        {
            this.insField = new CT_TrackChange();
            this.numberingChangeField = new CT_TrackChangeNumbering();
            this.numIdField = new CT_DecimalNumber();
            this.ilvlField = new CT_DecimalNumber();
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_DecimalNumber ilvl
        {
            get
            {
                return this.ilvlField;
            }
            set
            {
                this.ilvlField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
        public CT_DecimalNumber numId
        {
            get
            {
                return this.numIdField;
            }
            set
            {
                this.numIdField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 2)]
        public CT_TrackChangeNumbering numberingChange
        {
            get
            {
                return this.numberingChangeField;
            }
            set
            {
                this.numberingChangeField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 3)]
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
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Sym
    {

        private string fontField;

        private byte[] charField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string font
        {
            get
            {
                return this.fontField;
            }
            set
            {
                this.fontField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] @char
        {
            get
            {
                return this.charField;
            }
            set
            {
                this.charField = value;
            }
        }
    }

    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_AbstractNum
    {

        private CT_LongHexNumber nsidField;

        private CT_MultiLevelType multiLevelTypeField;

        private CT_LongHexNumber tmplField;

        private CT_String nameField;

        private CT_String styleLinkField;

        private CT_String numStyleLinkField;

        private List<CT_Lvl> lvlField;

        private string abstractNumIdField;

        public CT_AbstractNum()
        {
            this.lvlField = new List<CT_Lvl>();
            this.numStyleLinkField = new CT_String();
            this.styleLinkField = new CT_String();
            this.nameField = new CT_String();
            this.tmplField = new CT_LongHexNumber();
            this.multiLevelTypeField = new CT_MultiLevelType();
            this.nsidField = new CT_LongHexNumber();
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_LongHexNumber nsid
        {
            get
            {
                return this.nsidField;
            }
            set
            {
                this.nsidField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
        public CT_MultiLevelType multiLevelType
        {
            get
            {
                return this.multiLevelTypeField;
            }
            set
            {
                this.multiLevelTypeField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 2)]
        public CT_LongHexNumber tmpl
        {
            get
            {
                return this.tmplField;
            }
            set
            {
                this.tmplField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 3)]
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

        [System.Xml.Serialization.XmlElementAttribute(Order = 4)]
        public CT_String styleLink
        {
            get
            {
                return this.styleLinkField;
            }
            set
            {
                this.styleLinkField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 5)]
        public CT_String numStyleLink
        {
            get
            {
                return this.numStyleLinkField;
            }
            set
            {
                this.numStyleLinkField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute("lvl", Order = 6)]
        public List<CT_Lvl> lvl
        {
            get
            {
                return this.lvlField;
            }
            set
            {
                this.lvlField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string abstractNumId
        {
            get
            {
                return this.abstractNumIdField;
            }
            set
            {
                this.abstractNumIdField = value;
            }
        }
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_MultiLevelType
    {

        private ST_MultiLevelType valField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_MultiLevelType val
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
    public enum ST_MultiLevelType
    {

        /// <remarks/>
        singleLevel,

        /// <remarks/>
        multilevel,

        /// <remarks/>
        hybridMultilevel,
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Lvl
    {

        private CT_DecimalNumber startField;

        private CT_NumFmt numFmtField;

        private CT_DecimalNumber lvlRestartField;

        private CT_String pStyleField;

        private CT_OnOff isLglField;

        private CT_LevelSuffix suffField;

        private CT_LevelText lvlTextField;

        private CT_DecimalNumber lvlPicBulletIdField;

        private CT_LvlLegacy legacyField;

        private CT_Jc lvlJcField;

        private CT_PPr pPrField;

        private CT_RPr rPrField;

        private string ilvlField;

        private byte[] tplcField;

        private ST_OnOff tentativeField;

        private bool tentativeFieldSpecified;

        public CT_Lvl()
        {
            this.rPrField = new CT_RPr();
            this.pPrField = new CT_PPr();
            this.lvlJcField = new CT_Jc();
            this.legacyField = new CT_LvlLegacy();
            this.lvlPicBulletIdField = new CT_DecimalNumber();
            this.lvlTextField = new CT_LevelText();
            this.suffField = new CT_LevelSuffix();
            this.isLglField = new CT_OnOff();
            this.pStyleField = new CT_String();
            this.lvlRestartField = new CT_DecimalNumber();
            this.numFmtField = new CT_NumFmt();
            this.startField = new CT_DecimalNumber();
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_DecimalNumber start
        {
            get
            {
                return this.startField;
            }
            set
            {
                this.startField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
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

        [System.Xml.Serialization.XmlElementAttribute(Order = 2)]
        public CT_DecimalNumber lvlRestart
        {
            get
            {
                return this.lvlRestartField;
            }
            set
            {
                this.lvlRestartField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 3)]
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

        [System.Xml.Serialization.XmlElementAttribute(Order = 4)]
        public CT_OnOff isLgl
        {
            get
            {
                return this.isLglField;
            }
            set
            {
                this.isLglField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 5)]
        public CT_LevelSuffix suff
        {
            get
            {
                return this.suffField;
            }
            set
            {
                this.suffField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 6)]
        public CT_LevelText lvlText
        {
            get
            {
                return this.lvlTextField;
            }
            set
            {
                this.lvlTextField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 7)]
        public CT_DecimalNumber lvlPicBulletId
        {
            get
            {
                return this.lvlPicBulletIdField;
            }
            set
            {
                this.lvlPicBulletIdField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 8)]
        public CT_LvlLegacy legacy
        {
            get
            {
                return this.legacyField;
            }
            set
            {
                this.legacyField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 9)]
        public CT_Jc lvlJc
        {
            get
            {
                return this.lvlJcField;
            }
            set
            {
                this.lvlJcField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 10)]
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

        [System.Xml.Serialization.XmlElementAttribute(Order = 11)]
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string ilvl
        {
            get
            {
                return this.ilvlField;
            }
            set
            {
                this.ilvlField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] tplc
        {
            get
            {
                return this.tplcField;
            }
            set
            {
                this.tplcField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff tentative
        {
            get
            {
                return this.tentativeField;
            }
            set
            {
                this.tentativeField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool tentativeSpecified
        {
            get
            {
                return this.tentativeFieldSpecified;
            }
            set
            {
                this.tentativeFieldSpecified = value;
            }
        }
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_LevelSuffix
    {

        private ST_LevelSuffix valField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_LevelSuffix val
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
    public enum ST_LevelSuffix
    {

        /// <remarks/>
        tab,

        /// <remarks/>
        space,

        /// <remarks/>
        nothing,
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_LevelText
    {

        private string valField;

        private ST_OnOff nullField;

        private bool nullFieldSpecified;

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
        public ST_OnOff @null
        {
            get
            {
                return this.nullField;
            }
            set
            {
                this.nullField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool nullSpecified
        {
            get
            {
                return this.nullFieldSpecified;
            }
            set
            {
                this.nullFieldSpecified = value;
            }
        }
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_LvlLegacy
    {

        private ST_OnOff legacyField;

        private bool legacyFieldSpecified;

        private ulong legacySpaceField;

        private bool legacySpaceFieldSpecified;

        private string legacyIndentField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff legacy
        {
            get
            {
                return this.legacyField;
            }
            set
            {
                this.legacyField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool legacySpecified
        {
            get
            {
                return this.legacyFieldSpecified;
            }
            set
            {
                this.legacyFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong legacySpace
        {
            get
            {
                return this.legacySpaceField;
            }
            set
            {
                this.legacySpaceField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool legacySpaceSpecified
        {
            get
            {
                return this.legacySpaceFieldSpecified;
            }
            set
            {
                this.legacySpaceFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string legacyIndent
        {
            get
            {
                return this.legacyIndentField;
            }
            set
            {
                this.legacyIndentField = value;
            }
        }
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_LsdException
    {

        private string nameField;

        private ST_OnOff lockedField;

        private bool lockedFieldSpecified;

        private string uiPriorityField;

        private ST_OnOff semiHiddenField;

        private bool semiHiddenFieldSpecified;

        private ST_OnOff unhideWhenUsedField;

        private bool unhideWhenUsedFieldSpecified;

        private ST_OnOff qFormatField;

        private bool qFormatFieldSpecified;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string name
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff locked
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool lockedSpecified
        {
            get
            {
                return this.lockedFieldSpecified;
            }
            set
            {
                this.lockedFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string uiPriority
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff semiHidden
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool semiHiddenSpecified
        {
            get
            {
                return this.semiHiddenFieldSpecified;
            }
            set
            {
                this.semiHiddenFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff unhideWhenUsed
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool unhideWhenUsedSpecified
        {
            get
            {
                return this.unhideWhenUsedFieldSpecified;
            }
            set
            {
                this.unhideWhenUsedFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff qFormat
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool qFormatSpecified
        {
            get
            {
                return this.qFormatFieldSpecified;
            }
            set
            {
                this.qFormatFieldSpecified = value;
            }
        }
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TrackChangeNumbering : CT_TrackChange
    {

        private string originalField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string original
        {
            get
            {
                return this.originalField;
            }
            set
            {
                this.originalField = value;
            }
        }
    }

}
