
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

        public object SizeOfRArray()
        {
            throw new NotImplementedException();
        }
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_TextBodyProperties
    {

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

        public CT_TextBodyProperties()
        {
            this.uprightField = false;
        }

        public CT_PresetTextShape prstTxWarp
        {
            get
            {
                return this.prstTxWarpField;
            }
            set
            {
                this.prstTxWarpField = value;
            }
        }

        public CT_TextNoAutofit noAutofit
        {
            get
            {
                return this.noAutofitField;
            }
            set
            {
                this.noAutofitField = value;
            }
        }

        public CT_TextNormalAutofit normAutofit
        {
            get
            {
                return this.normAutofitField;
            }
            set
            {
                this.normAutofitField = value;
            }
        }
        public CT_TextShapeAutofit spAutoFit
        {
            get
            {
                return this.spAutoFitField;
            }
            set
            {
                this.spAutoFitField = value;
            }
        }


        public CT_Scene3D scene3d
        {
            get
            {
                return this.scene3dField;
            }
            set
            {
                this.scene3dField = value;
            }
        }


        public CT_Shape3D sp3d
        {
            get
            {
                return this.sp3dField;
            }
            set
            {
                this.sp3dField = value;
            }
        }


        public CT_FlatText flatTx
        {
            get
            {
                return this.flatTxField;
            }
            set
            {
                this.flatTxField = value;
            }
        }


        public CT_OfficeArtExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }


        [XmlAttribute]
        public int rot
        {
            get
            {
                return this.rotField;
            }
            set
            {
                this.rotField = value;
            }
        }


        [XmlIgnore]
        public bool rotSpecified
        {
            get
            {
                return this.rotFieldSpecified;
            }
            set
            {
                this.rotFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public bool spcFirstLastPara
        {
            get
            {
                return this.spcFirstLastParaField;
            }
            set
            {
                this.spcFirstLastParaField = value;
            }
        }


        [XmlIgnore]
        public bool spcFirstLastParaSpecified
        {
            get
            {
                return this.spcFirstLastParaFieldSpecified;
            }
            set
            {
                this.spcFirstLastParaFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public ST_TextVertOverflowType vertOverflow
        {
            get
            {
                return this.vertOverflowField;
            }
            set
            {
                this.vertOverflowField = value;
            }
        }


        [XmlIgnore]
        public bool vertOverflowSpecified
        {
            get
            {
                return this.vertOverflowFieldSpecified;
            }
            set
            {
                this.vertOverflowFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public ST_TextHorzOverflowType horzOverflow
        {
            get
            {
                return this.horzOverflowField;
            }
            set
            {
                this.horzOverflowField = value;
            }
        }


        [XmlIgnore]
        public bool horzOverflowSpecified
        {
            get
            {
                return this.horzOverflowFieldSpecified;
            }
            set
            {
                this.horzOverflowFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public ST_TextVerticalType vert
        {
            get
            {
                return this.vertField;
            }
            set
            {
                this.vertField = value;
            }
        }


        [XmlIgnore]
        public bool vertSpecified
        {
            get
            {
                return this.vertFieldSpecified;
            }
            set
            {
                this.vertFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public ST_TextWrappingType wrap
        {
            get
            {
                return this.wrapField;
            }
            set
            {
                this.wrapField = value;
            }
        }


        [XmlIgnore]
        public bool wrapSpecified
        {
            get
            {
                return this.wrapFieldSpecified;
            }
            set
            {
                this.wrapFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public int lIns
        {
            get
            {
                return this.lInsField;
            }
            set
            {
                this.lInsField = value;
            }
        }


        [XmlIgnore]
        public bool lInsSpecified
        {
            get
            {
                return this.lInsFieldSpecified;
            }
            set
            {
                this.lInsFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public int tIns
        {
            get
            {
                return this.tInsField;
            }
            set
            {
                this.tInsField = value;
            }
        }


        [XmlIgnore]
        public bool tInsSpecified
        {
            get
            {
                return this.tInsFieldSpecified;
            }
            set
            {
                this.tInsFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public int rIns
        {
            get
            {
                return this.rInsField;
            }
            set
            {
                this.rInsField = value;
            }
        }


        [XmlIgnore]
        public bool rInsSpecified
        {
            get
            {
                return this.rInsFieldSpecified;
            }
            set
            {
                this.rInsFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public int bIns
        {
            get
            {
                return this.bInsField;
            }
            set
            {
                this.bInsField = value;
            }
        }


        [XmlIgnore]
        public bool bInsSpecified
        {
            get
            {
                return this.bInsFieldSpecified;
            }
            set
            {
                this.bInsFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public int numCol
        {
            get
            {
                return this.numColField;
            }
            set
            {
                this.numColField = value;
            }
        }


        [XmlIgnore]
        public bool numColSpecified
        {
            get
            {
                return this.numColFieldSpecified;
            }
            set
            {
                this.numColFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public int spcCol
        {
            get
            {
                return this.spcColField;
            }
            set
            {
                this.spcColField = value;
            }
        }


        [XmlIgnore]
        public bool spcColSpecified
        {
            get
            {
                return this.spcColFieldSpecified;
            }
            set
            {
                this.spcColFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public bool rtlCol
        {
            get
            {
                return this.rtlColField;
            }
            set
            {
                this.rtlColField = value;
            }
        }


        [XmlIgnore]
        public bool rtlColSpecified
        {
            get
            {
                return this.rtlColFieldSpecified;
            }
            set
            {
                this.rtlColFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public bool fromWordArt
        {
            get
            {
                return this.fromWordArtField;
            }
            set
            {
                this.fromWordArtField = value;
            }
        }


        [XmlIgnore]
        public bool fromWordArtSpecified
        {
            get
            {
                return this.fromWordArtFieldSpecified;
            }
            set
            {
                this.fromWordArtFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public ST_TextAnchoringType anchor
        {
            get
            {
                return this.anchorField;
            }
            set
            {
                this.anchorField = value;
            }
        }


        [XmlIgnore]
        public bool anchorSpecified
        {
            get
            {
                return this.anchorFieldSpecified;
            }
            set
            {
                this.anchorFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public bool anchorCtr
        {
            get
            {
                return this.anchorCtrField;
            }
            set
            {
                this.anchorCtrField = value;
            }
        }


        [XmlIgnore]
        public bool anchorCtrSpecified
        {
            get
            {
                return this.anchorCtrFieldSpecified;
            }
            set
            {
                this.anchorCtrFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public bool forceAA
        {
            get
            {
                return this.forceAAField;
            }
            set
            {
                this.forceAAField = value;
            }
        }


        [XmlIgnore]
        public bool forceAASpecified
        {
            get
            {
                return this.forceAAFieldSpecified;
            }
            set
            {
                this.forceAAFieldSpecified = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(false)]
        public bool upright
        {
            get
            {
                return this.uprightField;
            }
            set
            {
                this.uprightField = value;
            }
        }


        [XmlAttribute]
        public bool compatLnSpc
        {
            get
            {
                return this.compatLnSpcField;
            }
            set
            {
                this.compatLnSpcField = value;
            }
        }


        [XmlIgnore]
        public bool compatLnSpcSpecified
        {
            get
            {
                return this.compatLnSpcFieldSpecified;
            }
            set
            {
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
            pField = new List<CT_TextParagraph>(array);
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
