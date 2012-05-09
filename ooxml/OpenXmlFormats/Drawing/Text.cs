
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Dml 
{
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_TextParagraph {
        
        private CT_TextParagraphProperties pPrField;
        
        private List<CT_RegularTextRun> rField;
        
        private List<CT_TextLineBreak> brField;
        
        private List<CT_TextField> fldField;
        
        private CT_TextCharacterProperties endParaRPrField;

        public CT_RegularTextRun AddNewR()
        {
            if(this.rField==null)
            this.rField = new List<CT_RegularTextRun>();
            CT_RegularTextRun rtr = new CT_RegularTextRun();
            this.rField.Add(rtr);
            return rtr;
        }
        public CT_TextParagraphProperties AddNewPPr()
        {
            this.pPrField = new CT_TextParagraphProperties();
            return this.pPrField;    
        }
        public CT_TextCharacterProperties AddNewEndParaRPr()
        {
            this.endParaRPrField=new CT_TextCharacterProperties();
            return this.endParaRPrField;
        }
    
        public CT_TextParagraphProperties pPr {
            get {
                return this.pPrField;
            }
            set {
                this.pPrField = value;
            }
        }
        
    
        [XmlElement("r")]
        public List<CT_RegularTextRun> r
        {
            get {
                return this.rField;
            }
            set {
                this.rField = value;
            }
        }
        
    
        [XmlElement("br")]
        public List<CT_TextLineBreak> br {
            get {
                return this.brField;
            }
            set {
                this.brField = value;
            }
        }
        
    
        [XmlElement("fld")]
        public List<CT_TextField> fld {
            get {
                return this.fldField;
            }
            set {
                this.fldField = value;
            }
        }
        
    
        public CT_TextCharacterProperties endParaRPr {
            get {
                return this.endParaRPrField;
            }
            set {
                this.endParaRPrField = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
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
        
    
        public CT_TextSpacing lnSpc {
            get {
                return this.lnSpcField;
            }
            set {
                this.lnSpcField = value;
            }
        }
        
    
        public CT_TextSpacing spcBef {
            get {
                return this.spcBefField;
            }
            set {
                this.spcBefField = value;
            }
        }
        
    
        public CT_TextSpacing spcAft {
            get {
                return this.spcAftField;
            }
            set {
                this.spcAftField = value;
            }
        }
        
    
        public CT_TextBulletColorFollowText buClrTx {
            get {
                return this.buClrTxField;
            }
            set {
                this.buClrTxField = value;
            }
        }
        
    
        public CT_Color buClr {
            get {
                return this.buClrField;
            }
            set {
                this.buClrField = value;
            }
        }
        
    
        public CT_TextBulletSizeFollowText buSzTx {
            get {
                return this.buSzTxField;
            }
            set {
                this.buSzTxField = value;
            }
        }
        
    
        public CT_TextBulletSizePercent buSzPct {
            get {
                return this.buSzPctField;
            }
            set {
                this.buSzPctField = value;
            }
        }
        
    
        public CT_TextBulletSizePoint buSzPts {
            get {
                return this.buSzPtsField;
            }
            set {
                this.buSzPtsField = value;
            }
        }
        
    
        public CT_TextBulletTypefaceFollowText buFontTx {
            get {
                return this.buFontTxField;
            }
            set {
                this.buFontTxField = value;
            }
        }
        
    
        public CT_TextFont buFont {
            get {
                return this.buFontField;
            }
            set {
                this.buFontField = value;
            }
        }
        
    
        public CT_TextNoBullet buNone {
            get {
                return this.buNoneField;
            }
            set {
                this.buNoneField = value;
            }
        }
        
    
        public CT_TextAutonumberBullet buAutoNum {
            get {
                return this.buAutoNumField;
            }
            set {
                this.buAutoNumField = value;
            }
        }
        
    
        public CT_TextCharBullet buChar {
            get {
                return this.buCharField;
            }
            set {
                this.buCharField = value;
            }
        }
        
    
        public CT_TextBlipBullet buBlip {
            get {
                return this.buBlipField;
            }
            set {
                this.buBlipField = value;
            }
        }
        
    
        [XmlArrayItem("tab", IsNullable=false)]
        public CT_TextTabStop[] tabLst {
            get {
                return this.tabLstField;
            }
            set {
                this.tabLstField = value;
            }
        }
        
    
        public CT_TextCharacterProperties defRPr {
            get {
                return this.defRPrField;
            }
            set {
                this.defRPrField = value;
            }
        }
        
    
        public CT_OfficeArtExtensionList extLst {
            get {
                return this.extLstField;
            }
            set {
                this.extLstField = value;
            }
        }
        
    
        [XmlAttribute]
        public int marL {
            get {
                return this.marLField;
            }
            set {
                this.marLField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool marLSpecified {
            get {
                return this.marLFieldSpecified;
            }
            set {
                this.marLFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public int marR {
            get {
                return this.marRField;
            }
            set {
                this.marRField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool marRSpecified {
            get {
                return this.marRFieldSpecified;
            }
            set {
                this.marRFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public int lvl {
            get {
                return this.lvlField;
            }
            set {
                this.lvlField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool lvlSpecified {
            get {
                return this.lvlFieldSpecified;
            }
            set {
                this.lvlFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public int indent {
            get {
                return this.indentField;
            }
            set {
                this.indentField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool indentSpecified {
            get {
                return this.indentFieldSpecified;
            }
            set {
                this.indentFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TextAlignType algn {
            get {
                return this.algnField;
            }
            set {
                this.algnField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool algnSpecified {
            get {
                return this.algnFieldSpecified;
            }
            set {
                this.algnFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public int defTabSz {
            get {
                return this.defTabSzField;
            }
            set {
                this.defTabSzField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool defTabSzSpecified {
            get {
                return this.defTabSzFieldSpecified;
            }
            set {
                this.defTabSzFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public bool rtl {
            get {
                return this.rtlField;
            }
            set {
                this.rtlField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool rtlSpecified {
            get {
                return this.rtlFieldSpecified;
            }
            set {
                this.rtlFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public bool eaLnBrk {
            get {
                return this.eaLnBrkField;
            }
            set {
                this.eaLnBrkField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool eaLnBrkSpecified {
            get {
                return this.eaLnBrkFieldSpecified;
            }
            set {
                this.eaLnBrkFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TextFontAlignType fontAlgn {
            get {
                return this.fontAlgnField;
            }
            set {
                this.fontAlgnField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool fontAlgnSpecified {
            get {
                return this.fontAlgnFieldSpecified;
            }
            set {
                this.fontAlgnFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public bool latinLnBrk {
            get {
                return this.latinLnBrkField;
            }
            set {
                this.latinLnBrkField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool latinLnBrkSpecified {
            get {
                return this.latinLnBrkFieldSpecified;
            }
            set {
                this.latinLnBrkFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public bool hangingPunct {
            get {
                return this.hangingPunctField;
            }
            set {
                this.hangingPunctField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool hangingPunctSpecified {
            get {
                return this.hangingPunctFieldSpecified;
            }
            set {
                this.hangingPunctFieldSpecified = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextSpacing {
        
        private object itemField;
        
    
        [XmlElement("spcPct", typeof(CT_TextSpacingPercent))]
        [XmlElement("spcPts", typeof(CT_TextSpacingPoint))]
        public object Item {
            get {
                return this.itemField;
            }
            set {
                this.itemField = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextSpacingPercent {
        
        private int valField;
        
    
        [XmlAttribute]
        public int val {
            get {
                return this.valField;
            }
            set {
                this.valField = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_FlatText {
        
        private long zField;
        
        public CT_FlatText() {
            this.zField = ((long)(0));
        }
        
    
        [XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long z {
            get {
                return this.zField;
            }
            set {
                this.zField = value;
            }
        }
    }
    
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextField {
        
        private CT_TextCharacterProperties rPrField;
        
        private CT_TextParagraphProperties pPrField;
        
        private string tField;
        
        private string idField;
        
        private string typeField;
        
    
        public CT_TextCharacterProperties rPr {
            get {
                return this.rPrField;
            }
            set {
                this.rPrField = value;
            }
        }
        
    
        public CT_TextParagraphProperties pPr {
            get {
                return this.pPrField;
            }
            set {
                this.pPrField = value;
            }
        }
        
    
        public string t {
            get {
                return this.tField;
            }
            set {
                this.tField = value;
            }
        }
        
    
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        //        [XmlAttribute(DataType="token")]
        public string id {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
        }
        
    
        [XmlAttribute]
        public string type {
            get {
                return this.typeField;
            }
            set {
                this.typeField = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
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
        public CT_TextFont AddNewLatin()
        {
            throw new NotImplementedException();
        }
    
        public CT_LineProperties ln {
            get {
                return this.lnField;
            }
            set {
                this.lnField = value;
            }
        }
        
    
        public CT_NoFillProperties noFill {
            get {
                return this.noFillField;
            }
            set {
                this.noFillField = value;
            }
        }
        
    
        public CT_SolidColorFillProperties solidFill {
            get {
                return this.solidFillField;
            }
            set {
                this.solidFillField = value;
            }
        }
        
    
        public CT_GradientFillProperties gradFill {
            get {
                return this.gradFillField;
            }
            set {
                this.gradFillField = value;
            }
        }
        
    
        public CT_BlipFillProperties blipFill {
            get {
                return this.blipFillField;
            }
            set {
                this.blipFillField = value;
            }
        }
        
    
        public CT_PatternFillProperties pattFill {
            get {
                return this.pattFillField;
            }
            set {
                this.pattFillField = value;
            }
        }
        
    
        public CT_GroupFillProperties grpFill {
            get {
                return this.grpFillField;
            }
            set {
                this.grpFillField = value;
            }
        }
        
    
        public CT_EffectList effectLst {
            get {
                return this.effectLstField;
            }
            set {
                this.effectLstField = value;
            }
        }
        
    
        public CT_EffectContainer effectDag {
            get {
                return this.effectDagField;
            }
            set {
                this.effectDagField = value;
            }
        }
        
    
        public CT_Color highlight {
            get {
                return this.highlightField;
            }
            set {
                this.highlightField = value;
            }
        }
        
    
        public CT_TextUnderlineLineFollowText uLnTx {
            get {
                return this.uLnTxField;
            }
            set {
                this.uLnTxField = value;
            }
        }
        
    
        public CT_LineProperties uLn {
            get {
                return this.uLnField;
            }
            set {
                this.uLnField = value;
            }
        }
        
    
        public CT_TextUnderlineFillFollowText uFillTx {
            get {
                return this.uFillTxField;
            }
            set {
                this.uFillTxField = value;
            }
        }
        
    
        public CT_TextUnderlineFillGroupWrapper uFill {
            get {
                return this.uFillField;
            }
            set {
                this.uFillField = value;
            }
        }
        
    
        public CT_TextFont latin {
            get {
                return this.latinField;
            }
            set {
                this.latinField = value;
            }
        }
        
    
        public CT_TextFont ea {
            get {
                return this.eaField;
            }
            set {
                this.eaField = value;
            }
        }
        
    
        public CT_TextFont cs {
            get {
                return this.csField;
            }
            set {
                this.csField = value;
            }
        }
        
    
        public CT_TextFont sym {
            get {
                return this.symField;
            }
            set {
                this.symField = value;
            }
        }
        
    
        public CT_Hyperlink hlinkClick {
            get {
                return this.hlinkClickField;
            }
            set {
                this.hlinkClickField = value;
            }
        }
        
    
        public CT_Hyperlink hlinkMouseOver {
            get {
                return this.hlinkMouseOverField;
            }
            set {
                this.hlinkMouseOverField = value;
            }
        }
        
    
        public CT_OfficeArtExtensionList extLst {
            get {
                return this.extLstField;
            }
            set {
                this.extLstField = value;
            }
        }
        
    
        [XmlAttribute]
        public bool kumimoji {
            get {
                return this.kumimojiField;
            }
            set {
                this.kumimojiField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool kumimojiSpecified {
            get {
                return this.kumimojiFieldSpecified;
            }
            set {
                this.kumimojiFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public string lang {
            get {
                return this.langField;
            }
            set {
                this.langField = value;
            }
        }
        
    
        [XmlAttribute]
        public string altLang {
            get {
                return this.altLangField;
            }
            set {
                this.altLangField = value;
            }
        }
        
    
        [XmlAttribute]
        public int sz {
            get {
                return this.szField;
            }
            set {
                this.szField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool szSpecified {
            get {
                return this.szFieldSpecified;
            }
            set {
                this.szFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public bool b {
            get {
                return this.bField;
            }
            set {
                this.bField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool bSpecified {
            get {
                return this.bFieldSpecified;
            }
            set {
                this.bFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public bool i {
            get {
                return this.iField;
            }
            set {
                this.iField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool iSpecified {
            get {
                return this.iFieldSpecified;
            }
            set {
                this.iFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TextUnderlineType u {
            get {
                return this.uField;
            }
            set {
                this.uField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool uSpecified {
            get {
                return this.uFieldSpecified;
            }
            set {
                this.uFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TextStrikeType strike {
            get {
                return this.strikeField;
            }
            set {
                this.strikeField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool strikeSpecified {
            get {
                return this.strikeFieldSpecified;
            }
            set {
                this.strikeFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public int kern {
            get {
                return this.kernField;
            }
            set {
                this.kernField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool kernSpecified {
            get {
                return this.kernFieldSpecified;
            }
            set {
                this.kernFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TextCapsType cap {
            get {
                return this.capField;
            }
            set {
                this.capField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool capSpecified {
            get {
                return this.capFieldSpecified;
            }
            set {
                this.capFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public int spc {
            get {
                return this.spcField;
            }
            set {
                this.spcField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool spcSpecified {
            get {
                return this.spcFieldSpecified;
            }
            set {
                this.spcFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public bool normalizeH {
            get {
                return this.normalizeHField;
            }
            set {
                this.normalizeHField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool normalizeHSpecified {
            get {
                return this.normalizeHFieldSpecified;
            }
            set {
                this.normalizeHFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public int baseline {
            get {
                return this.baselineField;
            }
            set {
                this.baselineField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool baselineSpecified {
            get {
                return this.baselineFieldSpecified;
            }
            set {
                this.baselineFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public bool noProof {
            get {
                return this.noProofField;
            }
            set {
                this.noProofField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool noProofSpecified {
            get {
                return this.noProofFieldSpecified;
            }
            set {
                this.noProofFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValue(true)]
        public bool dirty {
            get {
                return this.dirtyField;
            }
            set {
                this.dirtyField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValue(false)]
        public bool err {
            get {
                return this.errField;
            }
            set {
                this.errField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValue(true)]
        public bool smtClean {
            get {
                return this.smtCleanField;
            }
            set {
                this.smtCleanField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValue(typeof(uint), "0")]
        public uint smtId {
            get {
                return this.smtIdField;
            }
            set {
                this.smtIdField = value;
            }
        }
        
    
        [XmlAttribute]
        public string bmk {
            get {
                return this.bmkField;
            }
            set {
                this.bmkField = value;
            }
        }
    }
    
    
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextUnderlineLineFollowText {
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextUnderlineFillFollowText {
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextUnderlineFillGroupWrapper {
        
        private object itemField;
        
    
        [XmlElement("blipFill", typeof(CT_BlipFillProperties))]
        [XmlElement("gradFill", typeof(CT_GradientFillProperties))]
        [XmlElement("grpFill", typeof(CT_GroupFillProperties))]
        [XmlElement("noFill", typeof(CT_NoFillProperties))]
        [XmlElement("pattFill", typeof(CT_PatternFillProperties))]
        [XmlElement("solidFill", typeof(CT_SolidColorFillProperties))]
        public object Item {
            get {
                return this.itemField;
            }
            set {
                this.itemField = value;
            }
        }
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_TextUnderlineType {
        
    
        none,
        
    
        words,
        
    
        sng,
        
    
        dbl,
        
    
        heavy,
        
    
        dotted,
        
    
        dottedHeavy,
        
    
        dash,
        
    
        dashHeavy,
        
    
        dashLong,
        
    
        dashLongHeavy,
        
    
        dotDash,
        
    
        dotDashHeavy,
        
    
        dotDotDash,
        
    
        dotDotDashHeavy,
        
    
        wavy,
        
    
        wavyHeavy,
        
    
        wavyDbl,
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_TextStrikeType {
        
    
        noStrike,
        
    
        sngStrike,
        
    
        dblStrike,
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_TextCapsType {
        
    
        none,
        
    
        small,
        
    
        all,
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextLineBreak {
        
        private CT_TextCharacterProperties rPrField;
        
    
        public CT_TextCharacterProperties rPr {
            get {
                return this.rPrField;
            }
            set {
                this.rPrField = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_RegularTextRun {
        
        private CT_TextCharacterProperties rPrField;
        
        private string tField;

        public CT_TextCharacterProperties AddNewRPr()
        {
            this.rPrField = new CT_TextCharacterProperties();
            return rPrField;
        }
    
        public CT_TextCharacterProperties rPr {
            get {
                return this.rPrField;
            }
            set {
                this.rPrField = value;
            }
        }
        
    
        public string t {
            get {
                return this.tField;
            }
            set {
                this.tField = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextTabStop {
        
        private int posField;
        
        private bool posFieldSpecified;
        
        private ST_TextTabAlignType algnField;
        
        private bool algnFieldSpecified;
        
    
        [XmlAttribute]
        public int pos {
            get {
                return this.posField;
            }
            set {
                this.posField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool posSpecified {
            get {
                return this.posFieldSpecified;
            }
            set {
                this.posFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TextTabAlignType algn {
            get {
                return this.algnField;
            }
            set {
                this.algnField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool algnSpecified {
            get {
                return this.algnFieldSpecified;
            }
            set {
                this.algnFieldSpecified = value;
            }
        }
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_TextTabAlignType {
        
    
        l,
        
    
        ctr,
        
    
        r,
        
    
        dec,
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextBlipBullet {
        
        private CT_Blip blipField;
        
    
        public CT_Blip blip {
            get {
                return this.blipField;
            }
            set {
                this.blipField = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextCharBullet {
        
        private string charField;
        
    
        [XmlAttribute]
        public string @char {
            get {
                return this.charField;
            }
            set {
                this.charField = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextAutonumberBullet {
        
        private ST_TextAutonumberScheme typeField;
        
        private int startAtField;
        
        public CT_TextAutonumberBullet() {
            this.startAtField = 1;
        }
        
    
        [XmlAttribute]
        public ST_TextAutonumberScheme type {
            get {
                return this.typeField;
            }
            set {
                this.typeField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValue(1)]
        public int startAt {
            get {
                return this.startAtField;
            }
            set {
                this.startAtField = value;
            }
        }
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_TextAutonumberScheme {
        
    
        alphaLcParenBoth,
        
    
        alphaUcParenBoth,
        
    
        alphaLcParenR,
        
    
        alphaUcParenR,
        
    
        alphaLcPeriod,
        
    
        alphaUcPeriod,
        
    
        arabicParenBoth,
        
    
        arabicParenR,
        
    
        arabicPeriod,
        
    
        arabicPlain,
        
    
        romanLcParenBoth,
        
    
        romanUcParenBoth,
        
    
        romanLcParenR,
        
    
        romanUcParenR,
        
    
        romanLcPeriod,
        
    
        romanUcPeriod,
        
    
        circleNumDbPlain,
        
    
        circleNumWdBlackPlain,
        
    
        circleNumWdWhitePlain,
        
    
        arabicDbPeriod,
        
    
        arabicDbPlain,
        
    
        ea1ChsPeriod,
        
    
        ea1ChsPlain,
        
    
        ea1ChtPeriod,
        
    
        ea1ChtPlain,
        
    
        ea1JpnChsDbPeriod,
        
    
        ea1JpnKorPlain,
        
    
        ea1JpnKorPeriod,
        
    
        arabic1Minus,
        
    
        arabic2Minus,
        
    
        hebrew2Minus,
        
    
        thaiAlphaPeriod,
        
    
        thaiAlphaParenR,
        
    
        thaiAlphaParenBoth,
        
    
        thaiNumPeriod,
        
    
        thaiNumParenR,
        
    
        thaiNumParenBoth,
        
    
        hindiAlphaPeriod,
        
    
        hindiNumPeriod,
        
    
        hindiNumParenR,
        
    
        hindiAlpha1Period,
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextNoBullet {
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextBulletTypefaceFollowText {
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextBulletSizePoint {
        
        private int valField;
        
        private bool valFieldSpecified;
        
    
        [XmlAttribute]
        public int val {
            get {
                return this.valField;
            }
            set {
                this.valField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool valSpecified {
            get {
                return this.valFieldSpecified;
            }
            set {
                this.valFieldSpecified = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextBulletSizePercent {
        
        private int valField;
        
        private bool valFieldSpecified;
        
    
        [XmlAttribute]
        public int val {
            get {
                return this.valField;
            }
            set {
                this.valField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool valSpecified {
            get {
                return this.valFieldSpecified;
            }
            set {
                this.valFieldSpecified = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextBulletSizeFollowText {
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextBulletColorFollowText {
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_TextSpacingPoint {
        
        private int valField;
        
    
        [XmlAttribute]
        public int val {
            get {
                return this.valField;
            }
            set {
                this.valField = value;
            }
        }
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_TextAlignType {
        
    
        l,
        
    
        ctr,
        
    
        r,
        
    
        just,
        
    
        justLow,
        
    
        dist,
        
    
        thaiDist,
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_TextFontAlignType {
        
    
        auto,
        
    
        t,
        
    
        ctr,
        
    
        @base,
        
    
        b,
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_TextAnchoringType {
        
    
        t,
        
    
        ctr,
        
    
        b,
        
    
        just,
        
    
        dist,
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_TextVertOverflowType {
        
    
        overflow,
        
    
        ellipsis,
        
    
        clip,
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_TextHorzOverflowType {
        
    
        overflow,
        
    
        clip,
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_TextVerticalType {
        
    
        horz,
        
    
        vert,
        
    
        vert270,
        
    
        wordArtVert,
        
    
        eaVert,
        
    
        mongolianVert,
        
    
        wordArtVertRtl,
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_TextWrappingType {
        
    
        none,
        
    
        square,
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
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
        
    
        public CT_TextParagraphProperties defPPr {
            get {
                return this.defPPrField;
            }
            set {
                this.defPPrField = value;
            }
        }
        
    
        public CT_TextParagraphProperties lvl1pPr {
            get {
                return this.lvl1pPrField;
            }
            set {
                this.lvl1pPrField = value;
            }
        }
        
    
        public CT_TextParagraphProperties lvl2pPr {
            get {
                return this.lvl2pPrField;
            }
            set {
                this.lvl2pPrField = value;
            }
        }
        
    
        public CT_TextParagraphProperties lvl3pPr {
            get {
                return this.lvl3pPrField;
            }
            set {
                this.lvl3pPrField = value;
            }
        }
        
    
        public CT_TextParagraphProperties lvl4pPr {
            get {
                return this.lvl4pPrField;
            }
            set {
                this.lvl4pPrField = value;
            }
        }
        
    
        public CT_TextParagraphProperties lvl5pPr {
            get {
                return this.lvl5pPrField;
            }
            set {
                this.lvl5pPrField = value;
            }
        }
        
    
        public CT_TextParagraphProperties lvl6pPr {
            get {
                return this.lvl6pPrField;
            }
            set {
                this.lvl6pPrField = value;
            }
        }
        
    
        public CT_TextParagraphProperties lvl7pPr {
            get {
                return this.lvl7pPrField;
            }
            set {
                this.lvl7pPrField = value;
            }
        }
        
    
        public CT_TextParagraphProperties lvl8pPr {
            get {
                return this.lvl8pPrField;
            }
            set {
                this.lvl8pPrField = value;
            }
        }
        
    
        public CT_TextParagraphProperties lvl9pPr {
            get {
                return this.lvl9pPrField;
            }
            set {
                this.lvl9pPrField = value;
            }
        }
        
    
        public CT_OfficeArtExtensionList extLst {
            get {
                return this.extLstField;
            }
            set {
                this.extLstField = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_TextNormalAutofit {
        
        private int fontScaleField;
        
        private int lnSpcReductionField;
        
        public CT_TextNormalAutofit() {
            this.fontScaleField = 100000;
            this.lnSpcReductionField = 0;
        }
        
    
        [XmlAttribute]
        [DefaultValue(100000)]
        public int fontScale {
            get {
                return this.fontScaleField;
            }
            set {
                this.fontScaleField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValue(0)]
        public int lnSpcReduction {
            get {
                return this.lnSpcReductionField;
            }
            set {
                this.lnSpcReductionField = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_TextShapeAutofit {
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_TextNoAutofit {
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
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
        
    
        public CT_PresetTextShape prstTxWarp {
            get {
                return this.prstTxWarpField;
            }
            set {
                this.prstTxWarpField = value;
            }
        }
        
    
        public CT_TextNoAutofit noAutofit {
            get {
                return this.noAutofitField;
            }
            set {
                this.noAutofitField = value;
            }
        }
        
    
        public CT_TextNormalAutofit normAutofit {
            get {
                return this.normAutofitField;
            }
            set {
                this.normAutofitField = value;
            }
        }
        
    
        public CT_TextShapeAutofit spAutoFit {
            get {
                return this.spAutoFitField;
            }
            set {
                this.spAutoFitField = value;
            }
        }
        
    
        public CT_Scene3D scene3d {
            get {
                return this.scene3dField;
            }
            set {
                this.scene3dField = value;
            }
        }
        
    
        public CT_Shape3D sp3d {
            get {
                return this.sp3dField;
            }
            set {
                this.sp3dField = value;
            }
        }
        
    
        public CT_FlatText flatTx {
            get {
                return this.flatTxField;
            }
            set {
                this.flatTxField = value;
            }
        }
        
    
        public CT_OfficeArtExtensionList extLst {
            get {
                return this.extLstField;
            }
            set {
                this.extLstField = value;
            }
        }
        
    
        [XmlAttribute]
        public int rot {
            get {
                return this.rotField;
            }
            set {
                this.rotField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool rotSpecified {
            get {
                return this.rotFieldSpecified;
            }
            set {
                this.rotFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public bool spcFirstLastPara {
            get {
                return this.spcFirstLastParaField;
            }
            set {
                this.spcFirstLastParaField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool spcFirstLastParaSpecified {
            get {
                return this.spcFirstLastParaFieldSpecified;
            }
            set {
                this.spcFirstLastParaFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TextVertOverflowType vertOverflow {
            get {
                return this.vertOverflowField;
            }
            set {
                this.vertOverflowField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool vertOverflowSpecified {
            get {
                return this.vertOverflowFieldSpecified;
            }
            set {
                this.vertOverflowFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TextHorzOverflowType horzOverflow {
            get {
                return this.horzOverflowField;
            }
            set {
                this.horzOverflowField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool horzOverflowSpecified {
            get {
                return this.horzOverflowFieldSpecified;
            }
            set {
                this.horzOverflowFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TextVerticalType vert {
            get {
                return this.vertField;
            }
            set {
                this.vertField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool vertSpecified {
            get {
                return this.vertFieldSpecified;
            }
            set {
                this.vertFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TextWrappingType wrap {
            get {
                return this.wrapField;
            }
            set {
                this.wrapField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool wrapSpecified {
            get {
                return this.wrapFieldSpecified;
            }
            set {
                this.wrapFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public int lIns {
            get {
                return this.lInsField;
            }
            set {
                this.lInsField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool lInsSpecified {
            get {
                return this.lInsFieldSpecified;
            }
            set {
                this.lInsFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public int tIns {
            get {
                return this.tInsField;
            }
            set {
                this.tInsField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool tInsSpecified {
            get {
                return this.tInsFieldSpecified;
            }
            set {
                this.tInsFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public int rIns {
            get {
                return this.rInsField;
            }
            set {
                this.rInsField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool rInsSpecified {
            get {
                return this.rInsFieldSpecified;
            }
            set {
                this.rInsFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public int bIns {
            get {
                return this.bInsField;
            }
            set {
                this.bInsField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool bInsSpecified {
            get {
                return this.bInsFieldSpecified;
            }
            set {
                this.bInsFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public int numCol {
            get {
                return this.numColField;
            }
            set {
                this.numColField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool numColSpecified {
            get {
                return this.numColFieldSpecified;
            }
            set {
                this.numColFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public int spcCol {
            get {
                return this.spcColField;
            }
            set {
                this.spcColField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool spcColSpecified {
            get {
                return this.spcColFieldSpecified;
            }
            set {
                this.spcColFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public bool rtlCol {
            get {
                return this.rtlColField;
            }
            set {
                this.rtlColField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool rtlColSpecified {
            get {
                return this.rtlColFieldSpecified;
            }
            set {
                this.rtlColFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public bool fromWordArt {
            get {
                return this.fromWordArtField;
            }
            set {
                this.fromWordArtField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool fromWordArtSpecified {
            get {
                return this.fromWordArtFieldSpecified;
            }
            set {
                this.fromWordArtFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TextAnchoringType anchor {
            get {
                return this.anchorField;
            }
            set {
                this.anchorField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool anchorSpecified {
            get {
                return this.anchorFieldSpecified;
            }
            set {
                this.anchorFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public bool anchorCtr {
            get {
                return this.anchorCtrField;
            }
            set {
                this.anchorCtrField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool anchorCtrSpecified {
            get {
                return this.anchorCtrFieldSpecified;
            }
            set {
                this.anchorCtrFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public bool forceAA {
            get {
                return this.forceAAField;
            }
            set {
                this.forceAAField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool forceAASpecified {
            get {
                return this.forceAAFieldSpecified;
            }
            set {
                this.forceAAFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValue(false)]
        public bool upright {
            get {
                return this.uprightField;
            }
            set {
                this.uprightField = value;
            }
        }
        
    
        [XmlAttribute]
        public bool compatLnSpc {
            get {
                return this.compatLnSpcField;
            }
            set {
                this.compatLnSpcField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool compatLnSpcSpecified {
            get {
                return this.compatLnSpcFieldSpecified;
            }
            set {
                this.compatLnSpcFieldSpecified = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_TextBody {
        
        private CT_TextBodyProperties bodyPrField;
        
        private CT_TextListStyle lstStyleField;
        
        private List<CT_TextParagraph> pField;


        public void SetPArray(CT_TextParagraph[] array)
        {
            throw new NotImplementedException();
        }
        public CT_TextParagraph AddNewP()
        {
            if (this.pField == null)
                pField = new List<CT_TextParagraph>();
            CT_TextParagraph tp = new CT_TextParagraph();
            pField.Add(tp);
            return tp;
        }
        public CT_TextBodyProperties AddNewBodyPr()
        {
            this.bodyPrField = new CT_TextBodyProperties();
            return this.bodyPrField;
        }
        public CT_TextListStyle AddNewLstStyle()
        {
                this.lstStyleField=new CT_TextListStyle();
            return this.lstStyleField;   
        }
    
        public CT_TextBodyProperties bodyPr {
            get {
                return this.bodyPrField;
            }
            set {
                this.bodyPrField = value;
            }
        }
        
    
        public CT_TextListStyle lstStyle {
            get {
                return this.lstStyleField;
            }
            set {
                this.lstStyleField = value;
            }
        }
        
    
        [XmlElement("p")]
        public List<CT_TextParagraph> p {
            get {
                return this.pField;
            }
            set {
                this.pField = value;
            }
        }
    }
}
