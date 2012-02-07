namespace NPOI.OpenXmlFormats.Dml {
    
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_TextParagraph {
        
        private CT_TextParagraphProperties pPrField;
        
        private CT_RegularTextRun[] rField;
        
        private CT_TextLineBreak[] brField;
        
        private CT_TextField[] fldField;
        
        private CT_TextCharacterProperties endParaRPrField;
        
        /// <remarks/>
        public CT_TextParagraphProperties pPr {
            get {
                return this.pPrField;
            }
            set {
                this.pPrField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("r")]
        public CT_RegularTextRun[] r {
            get {
                return this.rField;
            }
            set {
                this.rField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("br")]
        public CT_TextLineBreak[] br {
            get {
                return this.brField;
            }
            set {
                this.brField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("fld")]
        public CT_TextField[] fld {
            get {
                return this.fldField;
            }
            set {
                this.fldField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextCharacterProperties endParaRPr {
            get {
                return this.endParaRPrField;
            }
            set {
                this.endParaRPrField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextParagraphProperties {
        
        private CT_TextSpacing lnSpcField;
        
        private CT_TextSpacing spcBefField;
        
        private CT_TextSpacing spcAftField;
        
        private CT_TextBulletColorFollowText buClrTxField;
        
        private CT_Color buClrField;
        
        private CT_TextBulletSizeFollowText buSzTxField;
        
        private CT_TextBulletSizePercent buSzPctField;
        
        private CT_TextBulletSizePoint buSzPtsField;
        
        private CT_TextBulletTypefaceFollowText buFontTxField;
        
        private CT_TextFont buFontField;
        
        private CT_TextNoBullet buNoneField;
        
        private CT_TextAutonumberBullet buAutoNumField;
        
        private CT_TextCharBullet buCharField;
        
        private CT_TextBlipBullet buBlipField;
        
        private CT_TextTabStop[] tabLstField;
        
        private CT_TextCharacterProperties defRPrField;
        
        private CT_OfficeArtExtensionList extLstField;
        
        private int marLField;
        
        private bool marLFieldSpecified;
        
        private int marRField;
        
        private bool marRFieldSpecified;
        
        private int lvlField;
        
        private bool lvlFieldSpecified;
        
        private int indentField;
        
        private bool indentFieldSpecified;
        
        private ST_TextAlignType algnField;
        
        private bool algnFieldSpecified;
        
        private int defTabSzField;
        
        private bool defTabSzFieldSpecified;
        
        private bool rtlField;
        
        private bool rtlFieldSpecified;
        
        private bool eaLnBrkField;
        
        private bool eaLnBrkFieldSpecified;
        
        private ST_TextFontAlignType fontAlgnField;
        
        private bool fontAlgnFieldSpecified;
        
        private bool latinLnBrkField;
        
        private bool latinLnBrkFieldSpecified;
        
        private bool hangingPunctField;
        
        private bool hangingPunctFieldSpecified;
        
        /// <remarks/>
        public CT_TextSpacing lnSpc {
            get {
                return this.lnSpcField;
            }
            set {
                this.lnSpcField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextSpacing spcBef {
            get {
                return this.spcBefField;
            }
            set {
                this.spcBefField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextSpacing spcAft {
            get {
                return this.spcAftField;
            }
            set {
                this.spcAftField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextBulletColorFollowText buClrTx {
            get {
                return this.buClrTxField;
            }
            set {
                this.buClrTxField = value;
            }
        }
        
        /// <remarks/>
        public CT_Color buClr {
            get {
                return this.buClrField;
            }
            set {
                this.buClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextBulletSizeFollowText buSzTx {
            get {
                return this.buSzTxField;
            }
            set {
                this.buSzTxField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextBulletSizePercent buSzPct {
            get {
                return this.buSzPctField;
            }
            set {
                this.buSzPctField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextBulletSizePoint buSzPts {
            get {
                return this.buSzPtsField;
            }
            set {
                this.buSzPtsField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextBulletTypefaceFollowText buFontTx {
            get {
                return this.buFontTxField;
            }
            set {
                this.buFontTxField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextFont buFont {
            get {
                return this.buFontField;
            }
            set {
                this.buFontField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextNoBullet buNone {
            get {
                return this.buNoneField;
            }
            set {
                this.buNoneField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextAutonumberBullet buAutoNum {
            get {
                return this.buAutoNumField;
            }
            set {
                this.buAutoNumField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextCharBullet buChar {
            get {
                return this.buCharField;
            }
            set {
                this.buCharField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextBlipBullet buBlip {
            get {
                return this.buBlipField;
            }
            set {
                this.buBlipField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("tab", IsNullable=false)]
        public CT_TextTabStop[] tabLst {
            get {
                return this.tabLstField;
            }
            set {
                this.tabLstField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextCharacterProperties defRPr {
            get {
                return this.defRPrField;
            }
            set {
                this.defRPrField = value;
            }
        }
        
        /// <remarks/>
        public CT_OfficeArtExtensionList extLst {
            get {
                return this.extLstField;
            }
            set {
                this.extLstField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int marL {
            get {
                return this.marLField;
            }
            set {
                this.marLField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool marLSpecified {
            get {
                return this.marLFieldSpecified;
            }
            set {
                this.marLFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int marR {
            get {
                return this.marRField;
            }
            set {
                this.marRField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool marRSpecified {
            get {
                return this.marRFieldSpecified;
            }
            set {
                this.marRFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int lvl {
            get {
                return this.lvlField;
            }
            set {
                this.lvlField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool lvlSpecified {
            get {
                return this.lvlFieldSpecified;
            }
            set {
                this.lvlFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int indent {
            get {
                return this.indentField;
            }
            set {
                this.indentField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool indentSpecified {
            get {
                return this.indentFieldSpecified;
            }
            set {
                this.indentFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_TextAlignType algn {
            get {
                return this.algnField;
            }
            set {
                this.algnField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool algnSpecified {
            get {
                return this.algnFieldSpecified;
            }
            set {
                this.algnFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int defTabSz {
            get {
                return this.defTabSzField;
            }
            set {
                this.defTabSzField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool defTabSzSpecified {
            get {
                return this.defTabSzFieldSpecified;
            }
            set {
                this.defTabSzFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool rtl {
            get {
                return this.rtlField;
            }
            set {
                this.rtlField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool rtlSpecified {
            get {
                return this.rtlFieldSpecified;
            }
            set {
                this.rtlFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool eaLnBrk {
            get {
                return this.eaLnBrkField;
            }
            set {
                this.eaLnBrkField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool eaLnBrkSpecified {
            get {
                return this.eaLnBrkFieldSpecified;
            }
            set {
                this.eaLnBrkFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_TextFontAlignType fontAlgn {
            get {
                return this.fontAlgnField;
            }
            set {
                this.fontAlgnField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool fontAlgnSpecified {
            get {
                return this.fontAlgnFieldSpecified;
            }
            set {
                this.fontAlgnFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool latinLnBrk {
            get {
                return this.latinLnBrkField;
            }
            set {
                this.latinLnBrkField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool latinLnBrkSpecified {
            get {
                return this.latinLnBrkFieldSpecified;
            }
            set {
                this.latinLnBrkFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool hangingPunct {
            get {
                return this.hangingPunctField;
            }
            set {
                this.hangingPunctField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool hangingPunctSpecified {
            get {
                return this.hangingPunctFieldSpecified;
            }
            set {
                this.hangingPunctFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextSpacing {
        
        private object itemField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("spcPct", typeof(CT_TextSpacingPercent))]
        [System.Xml.Serialization.XmlElementAttribute("spcPts", typeof(CT_TextSpacingPoint))]
        public object Item {
            get {
                return this.itemField;
            }
            set {
                this.itemField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextSpacingPercent {
        
        private int valField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int val {
            get {
                return this.valField;
            }
            set {
                this.valField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_FlatText {
        
        private long zField;
        
        public CT_FlatText() {
            this.zField = ((long)(0));
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(long), "0")]
        public long z {
            get {
                return this.zField;
            }
            set {
                this.zField = value;
            }
        }
    }
    
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextField {
        
        private CT_TextCharacterProperties rPrField;
        
        private CT_TextParagraphProperties pPrField;
        
        private string tField;
        
        private string idField;
        
        private string typeField;
        
        /// <remarks/>
        public CT_TextCharacterProperties rPr {
            get {
                return this.rPrField;
            }
            set {
                this.rPrField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextParagraphProperties pPr {
            get {
                return this.pPrField;
            }
            set {
                this.pPrField = value;
            }
        }
        
        /// <remarks/>
        public string t {
            get {
                return this.tField;
            }
            set {
                this.tField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="token")]
        public string id {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string type {
            get {
                return this.typeField;
            }
            set {
                this.typeField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextCharacterProperties {
        
        private CT_LineProperties lnField;
        
        private CT_NoFillProperties noFillField;
        
        private CT_SolidColorFillProperties solidFillField;
        
        private CT_GradientFillProperties gradFillField;
        
        private CT_BlipFillProperties blipFillField;
        
        private CT_PatternFillProperties pattFillField;
        
        private CT_GroupFillProperties grpFillField;
        
        private CT_EffectList effectLstField;
        
        private CT_EffectContainer effectDagField;
        
        private CT_Color highlightField;
        
        private CT_TextUnderlineLineFollowText uLnTxField;
        
        private CT_LineProperties uLnField;
        
        private CT_TextUnderlineFillFollowText uFillTxField;
        
        private CT_TextUnderlineFillGroupWrapper uFillField;
        
        private CT_TextFont latinField;
        
        private CT_TextFont eaField;
        
        private CT_TextFont csField;
        
        private CT_TextFont symField;
        
        private CT_Hyperlink hlinkClickField;
        
        private CT_Hyperlink hlinkMouseOverField;
        
        private CT_OfficeArtExtensionList extLstField;
        
        private bool kumimojiField;
        
        private bool kumimojiFieldSpecified;
        
        private string langField;
        
        private string altLangField;
        
        private int szField;
        
        private bool szFieldSpecified;
        
        private bool bField;
        
        private bool bFieldSpecified;
        
        private bool iField;
        
        private bool iFieldSpecified;
        
        private ST_TextUnderlineType uField;
        
        private bool uFieldSpecified;
        
        private ST_TextStrikeType strikeField;
        
        private bool strikeFieldSpecified;
        
        private int kernField;
        
        private bool kernFieldSpecified;
        
        private ST_TextCapsType capField;
        
        private bool capFieldSpecified;
        
        private int spcField;
        
        private bool spcFieldSpecified;
        
        private bool normalizeHField;
        
        private bool normalizeHFieldSpecified;
        
        private int baselineField;
        
        private bool baselineFieldSpecified;
        
        private bool noProofField;
        
        private bool noProofFieldSpecified;
        
        private bool dirtyField;
        
        private bool errField;
        
        private bool smtCleanField;
        
        private uint smtIdField;
        
        private string bmkField;
        
        public CT_TextCharacterProperties() {
            this.dirtyField = true;
            this.errField = false;
            this.smtCleanField = true;
            this.smtIdField = ((uint)(0));
        }
        
        /// <remarks/>
        public CT_LineProperties ln {
            get {
                return this.lnField;
            }
            set {
                this.lnField = value;
            }
        }
        
        /// <remarks/>
        public CT_NoFillProperties noFill {
            get {
                return this.noFillField;
            }
            set {
                this.noFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_SolidColorFillProperties solidFill {
            get {
                return this.solidFillField;
            }
            set {
                this.solidFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_GradientFillProperties gradFill {
            get {
                return this.gradFillField;
            }
            set {
                this.gradFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_BlipFillProperties blipFill {
            get {
                return this.blipFillField;
            }
            set {
                this.blipFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_PatternFillProperties pattFill {
            get {
                return this.pattFillField;
            }
            set {
                this.pattFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_GroupFillProperties grpFill {
            get {
                return this.grpFillField;
            }
            set {
                this.grpFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_EffectList effectLst {
            get {
                return this.effectLstField;
            }
            set {
                this.effectLstField = value;
            }
        }
        
        /// <remarks/>
        public CT_EffectContainer effectDag {
            get {
                return this.effectDagField;
            }
            set {
                this.effectDagField = value;
            }
        }
        
        /// <remarks/>
        public CT_Color highlight {
            get {
                return this.highlightField;
            }
            set {
                this.highlightField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextUnderlineLineFollowText uLnTx {
            get {
                return this.uLnTxField;
            }
            set {
                this.uLnTxField = value;
            }
        }
        
        /// <remarks/>
        public CT_LineProperties uLn {
            get {
                return this.uLnField;
            }
            set {
                this.uLnField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextUnderlineFillFollowText uFillTx {
            get {
                return this.uFillTxField;
            }
            set {
                this.uFillTxField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextUnderlineFillGroupWrapper uFill {
            get {
                return this.uFillField;
            }
            set {
                this.uFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextFont latin {
            get {
                return this.latinField;
            }
            set {
                this.latinField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextFont ea {
            get {
                return this.eaField;
            }
            set {
                this.eaField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextFont cs {
            get {
                return this.csField;
            }
            set {
                this.csField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextFont sym {
            get {
                return this.symField;
            }
            set {
                this.symField = value;
            }
        }
        
        /// <remarks/>
        public CT_Hyperlink hlinkClick {
            get {
                return this.hlinkClickField;
            }
            set {
                this.hlinkClickField = value;
            }
        }
        
        /// <remarks/>
        public CT_Hyperlink hlinkMouseOver {
            get {
                return this.hlinkMouseOverField;
            }
            set {
                this.hlinkMouseOverField = value;
            }
        }
        
        /// <remarks/>
        public CT_OfficeArtExtensionList extLst {
            get {
                return this.extLstField;
            }
            set {
                this.extLstField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool kumimoji {
            get {
                return this.kumimojiField;
            }
            set {
                this.kumimojiField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool kumimojiSpecified {
            get {
                return this.kumimojiFieldSpecified;
            }
            set {
                this.kumimojiFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string lang {
            get {
                return this.langField;
            }
            set {
                this.langField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string altLang {
            get {
                return this.altLangField;
            }
            set {
                this.altLangField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int sz {
            get {
                return this.szField;
            }
            set {
                this.szField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool szSpecified {
            get {
                return this.szFieldSpecified;
            }
            set {
                this.szFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool b {
            get {
                return this.bField;
            }
            set {
                this.bField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool bSpecified {
            get {
                return this.bFieldSpecified;
            }
            set {
                this.bFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool i {
            get {
                return this.iField;
            }
            set {
                this.iField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool iSpecified {
            get {
                return this.iFieldSpecified;
            }
            set {
                this.iFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_TextUnderlineType u {
            get {
                return this.uField;
            }
            set {
                this.uField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool uSpecified {
            get {
                return this.uFieldSpecified;
            }
            set {
                this.uFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_TextStrikeType strike {
            get {
                return this.strikeField;
            }
            set {
                this.strikeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool strikeSpecified {
            get {
                return this.strikeFieldSpecified;
            }
            set {
                this.strikeFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int kern {
            get {
                return this.kernField;
            }
            set {
                this.kernField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool kernSpecified {
            get {
                return this.kernFieldSpecified;
            }
            set {
                this.kernFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_TextCapsType cap {
            get {
                return this.capField;
            }
            set {
                this.capField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool capSpecified {
            get {
                return this.capFieldSpecified;
            }
            set {
                this.capFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int spc {
            get {
                return this.spcField;
            }
            set {
                this.spcField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool spcSpecified {
            get {
                return this.spcFieldSpecified;
            }
            set {
                this.spcFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool normalizeH {
            get {
                return this.normalizeHField;
            }
            set {
                this.normalizeHField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool normalizeHSpecified {
            get {
                return this.normalizeHFieldSpecified;
            }
            set {
                this.normalizeHFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int baseline {
            get {
                return this.baselineField;
            }
            set {
                this.baselineField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool baselineSpecified {
            get {
                return this.baselineFieldSpecified;
            }
            set {
                this.baselineFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool noProof {
            get {
                return this.noProofField;
            }
            set {
                this.noProofField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool noProofSpecified {
            get {
                return this.noProofFieldSpecified;
            }
            set {
                this.noProofFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dirty {
            get {
                return this.dirtyField;
            }
            set {
                this.dirtyField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool err {
            get {
                return this.errField;
            }
            set {
                this.errField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool smtClean {
            get {
                return this.smtCleanField;
            }
            set {
                this.smtCleanField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint smtId {
            get {
                return this.smtIdField;
            }
            set {
                this.smtIdField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string bmk {
            get {
                return this.bmkField;
            }
            set {
                this.bmkField = value;
            }
        }
    }
    
        /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextUnderlineLineFollowText {
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextUnderlineFillFollowText {
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextUnderlineFillGroupWrapper {
        
        private object itemField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("blipFill", typeof(CT_BlipFillProperties))]
        [System.Xml.Serialization.XmlElementAttribute("gradFill", typeof(CT_GradientFillProperties))]
        [System.Xml.Serialization.XmlElementAttribute("grpFill", typeof(CT_GroupFillProperties))]
        [System.Xml.Serialization.XmlElementAttribute("noFill", typeof(CT_NoFillProperties))]
        [System.Xml.Serialization.XmlElementAttribute("pattFill", typeof(CT_PatternFillProperties))]
        [System.Xml.Serialization.XmlElementAttribute("solidFill", typeof(CT_SolidColorFillProperties))]
        public object Item {
            get {
                return this.itemField;
            }
            set {
                this.itemField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_TextUnderlineType {
        
        /// <remarks/>
        none,
        
        /// <remarks/>
        words,
        
        /// <remarks/>
        sng,
        
        /// <remarks/>
        dbl,
        
        /// <remarks/>
        heavy,
        
        /// <remarks/>
        dotted,
        
        /// <remarks/>
        dottedHeavy,
        
        /// <remarks/>
        dash,
        
        /// <remarks/>
        dashHeavy,
        
        /// <remarks/>
        dashLong,
        
        /// <remarks/>
        dashLongHeavy,
        
        /// <remarks/>
        dotDash,
        
        /// <remarks/>
        dotDashHeavy,
        
        /// <remarks/>
        dotDotDash,
        
        /// <remarks/>
        dotDotDashHeavy,
        
        /// <remarks/>
        wavy,
        
        /// <remarks/>
        wavyHeavy,
        
        /// <remarks/>
        wavyDbl,
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_TextStrikeType {
        
        /// <remarks/>
        noStrike,
        
        /// <remarks/>
        sngStrike,
        
        /// <remarks/>
        dblStrike,
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_TextCapsType {
        
        /// <remarks/>
        none,
        
        /// <remarks/>
        small,
        
        /// <remarks/>
        all,
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextLineBreak {
        
        private CT_TextCharacterProperties rPrField;
        
        /// <remarks/>
        public CT_TextCharacterProperties rPr {
            get {
                return this.rPrField;
            }
            set {
                this.rPrField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_RegularTextRun {
        
        private CT_TextCharacterProperties rPrField;
        
        private string tField;
        
        /// <remarks/>
        public CT_TextCharacterProperties rPr {
            get {
                return this.rPrField;
            }
            set {
                this.rPrField = value;
            }
        }
        
        /// <remarks/>
        public string t {
            get {
                return this.tField;
            }
            set {
                this.tField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextTabStop {
        
        private int posField;
        
        private bool posFieldSpecified;
        
        private ST_TextTabAlignType algnField;
        
        private bool algnFieldSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int pos {
            get {
                return this.posField;
            }
            set {
                this.posField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool posSpecified {
            get {
                return this.posFieldSpecified;
            }
            set {
                this.posFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_TextTabAlignType algn {
            get {
                return this.algnField;
            }
            set {
                this.algnField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool algnSpecified {
            get {
                return this.algnFieldSpecified;
            }
            set {
                this.algnFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_TextTabAlignType {
        
        /// <remarks/>
        l,
        
        /// <remarks/>
        ctr,
        
        /// <remarks/>
        r,
        
        /// <remarks/>
        dec,
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextBlipBullet {
        
        private CT_Blip blipField;
        
        /// <remarks/>
        public CT_Blip blip {
            get {
                return this.blipField;
            }
            set {
                this.blipField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextCharBullet {
        
        private string charField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string @char {
            get {
                return this.charField;
            }
            set {
                this.charField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextAutonumberBullet {
        
        private ST_TextAutonumberScheme typeField;
        
        private int startAtField;
        
        public CT_TextAutonumberBullet() {
            this.startAtField = 1;
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_TextAutonumberScheme type {
            get {
                return this.typeField;
            }
            set {
                this.typeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(1)]
        public int startAt {
            get {
                return this.startAtField;
            }
            set {
                this.startAtField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_TextAutonumberScheme {
        
        /// <remarks/>
        alphaLcParenBoth,
        
        /// <remarks/>
        alphaUcParenBoth,
        
        /// <remarks/>
        alphaLcParenR,
        
        /// <remarks/>
        alphaUcParenR,
        
        /// <remarks/>
        alphaLcPeriod,
        
        /// <remarks/>
        alphaUcPeriod,
        
        /// <remarks/>
        arabicParenBoth,
        
        /// <remarks/>
        arabicParenR,
        
        /// <remarks/>
        arabicPeriod,
        
        /// <remarks/>
        arabicPlain,
        
        /// <remarks/>
        romanLcParenBoth,
        
        /// <remarks/>
        romanUcParenBoth,
        
        /// <remarks/>
        romanLcParenR,
        
        /// <remarks/>
        romanUcParenR,
        
        /// <remarks/>
        romanLcPeriod,
        
        /// <remarks/>
        romanUcPeriod,
        
        /// <remarks/>
        circleNumDbPlain,
        
        /// <remarks/>
        circleNumWdBlackPlain,
        
        /// <remarks/>
        circleNumWdWhitePlain,
        
        /// <remarks/>
        arabicDbPeriod,
        
        /// <remarks/>
        arabicDbPlain,
        
        /// <remarks/>
        ea1ChsPeriod,
        
        /// <remarks/>
        ea1ChsPlain,
        
        /// <remarks/>
        ea1ChtPeriod,
        
        /// <remarks/>
        ea1ChtPlain,
        
        /// <remarks/>
        ea1JpnChsDbPeriod,
        
        /// <remarks/>
        ea1JpnKorPlain,
        
        /// <remarks/>
        ea1JpnKorPeriod,
        
        /// <remarks/>
        arabic1Minus,
        
        /// <remarks/>
        arabic2Minus,
        
        /// <remarks/>
        hebrew2Minus,
        
        /// <remarks/>
        thaiAlphaPeriod,
        
        /// <remarks/>
        thaiAlphaParenR,
        
        /// <remarks/>
        thaiAlphaParenBoth,
        
        /// <remarks/>
        thaiNumPeriod,
        
        /// <remarks/>
        thaiNumParenR,
        
        /// <remarks/>
        thaiNumParenBoth,
        
        /// <remarks/>
        hindiAlphaPeriod,
        
        /// <remarks/>
        hindiNumPeriod,
        
        /// <remarks/>
        hindiNumParenR,
        
        /// <remarks/>
        hindiAlpha1Period,
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextNoBullet {
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextBulletTypefaceFollowText {
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextBulletSizePoint {
        
        private int valField;
        
        private bool valFieldSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int val {
            get {
                return this.valField;
            }
            set {
                this.valField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool valSpecified {
            get {
                return this.valFieldSpecified;
            }
            set {
                this.valFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextBulletSizePercent {
        
        private int valField;
        
        private bool valFieldSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int val {
            get {
                return this.valField;
            }
            set {
                this.valField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool valSpecified {
            get {
                return this.valFieldSpecified;
            }
            set {
                this.valFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextBulletSizeFollowText {
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextBulletColorFollowText {
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextSpacingPoint {
        
        private int valField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int val {
            get {
                return this.valField;
            }
            set {
                this.valField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_TextAlignType {
        
        /// <remarks/>
        l,
        
        /// <remarks/>
        ctr,
        
        /// <remarks/>
        r,
        
        /// <remarks/>
        just,
        
        /// <remarks/>
        justLow,
        
        /// <remarks/>
        dist,
        
        /// <remarks/>
        thaiDist,
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_TextFontAlignType {
        
        /// <remarks/>
        auto,
        
        /// <remarks/>
        t,
        
        /// <remarks/>
        ctr,
        
        /// <remarks/>
        @base,
        
        /// <remarks/>
        b,
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_TextAnchoringType {
        
        /// <remarks/>
        t,
        
        /// <remarks/>
        ctr,
        
        /// <remarks/>
        b,
        
        /// <remarks/>
        just,
        
        /// <remarks/>
        dist,
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_TextVertOverflowType {
        
        /// <remarks/>
        overflow,
        
        /// <remarks/>
        ellipsis,
        
        /// <remarks/>
        clip,
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_TextHorzOverflowType {
        
        /// <remarks/>
        overflow,
        
        /// <remarks/>
        clip,
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_TextVerticalType {
        
        /// <remarks/>
        horz,
        
        /// <remarks/>
        vert,
        
        /// <remarks/>
        vert270,
        
        /// <remarks/>
        wordArtVert,
        
        /// <remarks/>
        eaVert,
        
        /// <remarks/>
        mongolianVert,
        
        /// <remarks/>
        wordArtVertRtl,
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_TextWrappingType {
        
        /// <remarks/>
        none,
        
        /// <remarks/>
        square,
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_TextListStyle {
        
        private CT_TextParagraphProperties defPPrField;
        
        private CT_TextParagraphProperties lvl1pPrField;
        
        private CT_TextParagraphProperties lvl2pPrField;
        
        private CT_TextParagraphProperties lvl3pPrField;
        
        private CT_TextParagraphProperties lvl4pPrField;
        
        private CT_TextParagraphProperties lvl5pPrField;
        
        private CT_TextParagraphProperties lvl6pPrField;
        
        private CT_TextParagraphProperties lvl7pPrField;
        
        private CT_TextParagraphProperties lvl8pPrField;
        
        private CT_TextParagraphProperties lvl9pPrField;
        
        private CT_OfficeArtExtensionList extLstField;
        
        /// <remarks/>
        public CT_TextParagraphProperties defPPr {
            get {
                return this.defPPrField;
            }
            set {
                this.defPPrField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextParagraphProperties lvl1pPr {
            get {
                return this.lvl1pPrField;
            }
            set {
                this.lvl1pPrField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextParagraphProperties lvl2pPr {
            get {
                return this.lvl2pPrField;
            }
            set {
                this.lvl2pPrField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextParagraphProperties lvl3pPr {
            get {
                return this.lvl3pPrField;
            }
            set {
                this.lvl3pPrField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextParagraphProperties lvl4pPr {
            get {
                return this.lvl4pPrField;
            }
            set {
                this.lvl4pPrField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextParagraphProperties lvl5pPr {
            get {
                return this.lvl5pPrField;
            }
            set {
                this.lvl5pPrField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextParagraphProperties lvl6pPr {
            get {
                return this.lvl6pPrField;
            }
            set {
                this.lvl6pPrField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextParagraphProperties lvl7pPr {
            get {
                return this.lvl7pPrField;
            }
            set {
                this.lvl7pPrField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextParagraphProperties lvl8pPr {
            get {
                return this.lvl8pPrField;
            }
            set {
                this.lvl8pPrField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextParagraphProperties lvl9pPr {
            get {
                return this.lvl9pPrField;
            }
            set {
                this.lvl9pPrField = value;
            }
        }
        
        /// <remarks/>
        public CT_OfficeArtExtensionList extLst {
            get {
                return this.extLstField;
            }
            set {
                this.extLstField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_TextNormalAutofit {
        
        private int fontScaleField;
        
        private int lnSpcReductionField;
        
        public CT_TextNormalAutofit() {
            this.fontScaleField = 100000;
            this.lnSpcReductionField = 0;
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(100000)]
        public int fontScale {
            get {
                return this.fontScaleField;
            }
            set {
                this.fontScaleField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(0)]
        public int lnSpcReduction {
            get {
                return this.lnSpcReductionField;
            }
            set {
                this.lnSpcReductionField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_TextShapeAutofit {
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_TextNoAutofit {
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_TextBodyProperties {
        
        private CT_PresetTextShape prstTxWarpField;
        
        private CT_TextNoAutofit noAutofitField;
        
        private CT_TextNormalAutofit normAutofitField;
        
        private CT_TextShapeAutofit spAutoFitField;
        
        private CT_Scene3D scene3dField;
        
        private CT_Shape3D sp3dField;
        
        private CT_FlatText flatTxField;
        
        private CT_OfficeArtExtensionList extLstField;
        
        private int rotField;
        
        private bool rotFieldSpecified;
        
        private bool spcFirstLastParaField;
        
        private bool spcFirstLastParaFieldSpecified;
        
        private ST_TextVertOverflowType vertOverflowField;
        
        private bool vertOverflowFieldSpecified;
        
        private ST_TextHorzOverflowType horzOverflowField;
        
        private bool horzOverflowFieldSpecified;
        
        private ST_TextVerticalType vertField;
        
        private bool vertFieldSpecified;
        
        private ST_TextWrappingType wrapField;
        
        private bool wrapFieldSpecified;
        
        private int lInsField;
        
        private bool lInsFieldSpecified;
        
        private int tInsField;
        
        private bool tInsFieldSpecified;
        
        private int rInsField;
        
        private bool rInsFieldSpecified;
        
        private int bInsField;
        
        private bool bInsFieldSpecified;
        
        private int numColField;
        
        private bool numColFieldSpecified;
        
        private int spcColField;
        
        private bool spcColFieldSpecified;
        
        private bool rtlColField;
        
        private bool rtlColFieldSpecified;
        
        private bool fromWordArtField;
        
        private bool fromWordArtFieldSpecified;
        
        private ST_TextAnchoringType anchorField;
        
        private bool anchorFieldSpecified;
        
        private bool anchorCtrField;
        
        private bool anchorCtrFieldSpecified;
        
        private bool forceAAField;
        
        private bool forceAAFieldSpecified;
        
        private bool uprightField;
        
        private bool compatLnSpcField;
        
        private bool compatLnSpcFieldSpecified;
        
        public CT_TextBodyProperties() {
            this.uprightField = false;
        }
        
        /// <remarks/>
        public CT_PresetTextShape prstTxWarp {
            get {
                return this.prstTxWarpField;
            }
            set {
                this.prstTxWarpField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextNoAutofit noAutofit {
            get {
                return this.noAutofitField;
            }
            set {
                this.noAutofitField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextNormalAutofit normAutofit {
            get {
                return this.normAutofitField;
            }
            set {
                this.normAutofitField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextShapeAutofit spAutoFit {
            get {
                return this.spAutoFitField;
            }
            set {
                this.spAutoFitField = value;
            }
        }
        
        /// <remarks/>
        public CT_Scene3D scene3d {
            get {
                return this.scene3dField;
            }
            set {
                this.scene3dField = value;
            }
        }
        
        /// <remarks/>
        public CT_Shape3D sp3d {
            get {
                return this.sp3dField;
            }
            set {
                this.sp3dField = value;
            }
        }
        
        /// <remarks/>
        public CT_FlatText flatTx {
            get {
                return this.flatTxField;
            }
            set {
                this.flatTxField = value;
            }
        }
        
        /// <remarks/>
        public CT_OfficeArtExtensionList extLst {
            get {
                return this.extLstField;
            }
            set {
                this.extLstField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int rot {
            get {
                return this.rotField;
            }
            set {
                this.rotField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool rotSpecified {
            get {
                return this.rotFieldSpecified;
            }
            set {
                this.rotFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool spcFirstLastPara {
            get {
                return this.spcFirstLastParaField;
            }
            set {
                this.spcFirstLastParaField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool spcFirstLastParaSpecified {
            get {
                return this.spcFirstLastParaFieldSpecified;
            }
            set {
                this.spcFirstLastParaFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_TextVertOverflowType vertOverflow {
            get {
                return this.vertOverflowField;
            }
            set {
                this.vertOverflowField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool vertOverflowSpecified {
            get {
                return this.vertOverflowFieldSpecified;
            }
            set {
                this.vertOverflowFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_TextHorzOverflowType horzOverflow {
            get {
                return this.horzOverflowField;
            }
            set {
                this.horzOverflowField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool horzOverflowSpecified {
            get {
                return this.horzOverflowFieldSpecified;
            }
            set {
                this.horzOverflowFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_TextVerticalType vert {
            get {
                return this.vertField;
            }
            set {
                this.vertField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool vertSpecified {
            get {
                return this.vertFieldSpecified;
            }
            set {
                this.vertFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_TextWrappingType wrap {
            get {
                return this.wrapField;
            }
            set {
                this.wrapField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool wrapSpecified {
            get {
                return this.wrapFieldSpecified;
            }
            set {
                this.wrapFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int lIns {
            get {
                return this.lInsField;
            }
            set {
                this.lInsField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool lInsSpecified {
            get {
                return this.lInsFieldSpecified;
            }
            set {
                this.lInsFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int tIns {
            get {
                return this.tInsField;
            }
            set {
                this.tInsField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool tInsSpecified {
            get {
                return this.tInsFieldSpecified;
            }
            set {
                this.tInsFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int rIns {
            get {
                return this.rInsField;
            }
            set {
                this.rInsField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool rInsSpecified {
            get {
                return this.rInsFieldSpecified;
            }
            set {
                this.rInsFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int bIns {
            get {
                return this.bInsField;
            }
            set {
                this.bInsField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool bInsSpecified {
            get {
                return this.bInsFieldSpecified;
            }
            set {
                this.bInsFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int numCol {
            get {
                return this.numColField;
            }
            set {
                this.numColField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool numColSpecified {
            get {
                return this.numColFieldSpecified;
            }
            set {
                this.numColFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int spcCol {
            get {
                return this.spcColField;
            }
            set {
                this.spcColField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool spcColSpecified {
            get {
                return this.spcColFieldSpecified;
            }
            set {
                this.spcColFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool rtlCol {
            get {
                return this.rtlColField;
            }
            set {
                this.rtlColField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool rtlColSpecified {
            get {
                return this.rtlColFieldSpecified;
            }
            set {
                this.rtlColFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool fromWordArt {
            get {
                return this.fromWordArtField;
            }
            set {
                this.fromWordArtField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool fromWordArtSpecified {
            get {
                return this.fromWordArtFieldSpecified;
            }
            set {
                this.fromWordArtFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_TextAnchoringType anchor {
            get {
                return this.anchorField;
            }
            set {
                this.anchorField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool anchorSpecified {
            get {
                return this.anchorFieldSpecified;
            }
            set {
                this.anchorFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool anchorCtr {
            get {
                return this.anchorCtrField;
            }
            set {
                this.anchorCtrField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool anchorCtrSpecified {
            get {
                return this.anchorCtrFieldSpecified;
            }
            set {
                this.anchorCtrFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool forceAA {
            get {
                return this.forceAAField;
            }
            set {
                this.forceAAField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool forceAASpecified {
            get {
                return this.forceAAFieldSpecified;
            }
            set {
                this.forceAAFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool upright {
            get {
                return this.uprightField;
            }
            set {
                this.uprightField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool compatLnSpc {
            get {
                return this.compatLnSpcField;
            }
            set {
                this.compatLnSpcField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool compatLnSpcSpecified {
            get {
                return this.compatLnSpcFieldSpecified;
            }
            set {
                this.compatLnSpcFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_TextBody {
        
        private CT_TextBodyProperties bodyPrField;
        
        private CT_TextListStyle lstStyleField;
        
        private CT_TextParagraph[] pField;
        
        /// <remarks/>
        public CT_TextBodyProperties bodyPr {
            get {
                return this.bodyPrField;
            }
            set {
                this.bodyPrField = value;
            }
        }
        
        /// <remarks/>
        public CT_TextListStyle lstStyle {
            get {
                return this.lstStyleField;
            }
            set {
                this.lstStyleField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("p")]
        public CT_TextParagraph[] p {
            get {
                return this.pField;
            }
            set {
                this.pField = value;
            }
        }
    }
}
