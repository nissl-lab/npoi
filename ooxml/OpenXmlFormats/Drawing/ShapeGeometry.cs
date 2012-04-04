using System.Collections.Generic;
namespace NPOI.OpenXmlFormats.Dml {
    
    
    /// <remarks/>
    [System.Serializable]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_ShapeType {
        
        /// <remarks/>
        line,
        
        /// <remarks/>
        lineInv,
        
        /// <remarks/>
        triangle,
        
        /// <remarks/>
        rtTriangle,
        
        /// <remarks/>
        rect,
        
        /// <remarks/>
        diamond,
        
        /// <remarks/>
        parallelogram,
        
        /// <remarks/>
        trapezoid,
        
        /// <remarks/>
        nonIsoscelesTrapezoid,
        
        /// <remarks/>
        pentagon,
        
        /// <remarks/>
        hexagon,
        
        /// <remarks/>
        heptagon,
        
        /// <remarks/>
        octagon,
        
        /// <remarks/>
        decagon,
        
        /// <remarks/>
        dodecagon,
        
        /// <remarks/>
        star4,
        
        /// <remarks/>
        star5,
        
        /// <remarks/>
        star6,
        
        /// <remarks/>
        star7,
        
        /// <remarks/>
        star8,
        
        /// <remarks/>
        star10,
        
        /// <remarks/>
        star12,
        
        /// <remarks/>
        star16,
        
        /// <remarks/>
        star24,
        
        /// <remarks/>
        star32,
        
        /// <remarks/>
        roundRect,
        
        /// <remarks/>
        round1Rect,
        
        /// <remarks/>
        round2SameRect,
        
        /// <remarks/>
        round2DiagRect,
        
        /// <remarks/>
        snipRoundRect,
        
        /// <remarks/>
        snip1Rect,
        
        /// <remarks/>
        snip2SameRect,
        
        /// <remarks/>
        snip2DiagRect,
        
        /// <remarks/>
        plaque,
        
        /// <remarks/>
        ellipse,
        
        /// <remarks/>
        teardrop,
        
        /// <remarks/>
        homePlate,
        
        /// <remarks/>
        chevron,
        
        /// <remarks/>
        pieWedge,
        
        /// <remarks/>
        pie,
        
        /// <remarks/>
        blockArc,
        
        /// <remarks/>
        donut,
        
        /// <remarks/>
        noSmoking,
        
        /// <remarks/>
        rightArrow,
        
        /// <remarks/>
        leftArrow,
        
        /// <remarks/>
        upArrow,
        
        /// <remarks/>
        downArrow,
        
        /// <remarks/>
        stripedRightArrow,
        
        /// <remarks/>
        notchedRightArrow,
        
        /// <remarks/>
        bentUpArrow,
        
        /// <remarks/>
        leftRightArrow,
        
        /// <remarks/>
        upDownArrow,
        
        /// <remarks/>
        leftUpArrow,
        
        /// <remarks/>
        leftRightUpArrow,
        
        /// <remarks/>
        quadArrow,
        
        /// <remarks/>
        leftArrowCallout,
        
        /// <remarks/>
        rightArrowCallout,
        
        /// <remarks/>
        upArrowCallout,
        
        /// <remarks/>
        downArrowCallout,
        
        /// <remarks/>
        leftRightArrowCallout,
        
        /// <remarks/>
        upDownArrowCallout,
        
        /// <remarks/>
        quadArrowCallout,
        
        /// <remarks/>
        bentArrow,
        
        /// <remarks/>
        uturnArrow,
        
        /// <remarks/>
        circularArrow,
        
        /// <remarks/>
        leftCircularArrow,
        
        /// <remarks/>
        leftRightCircularArrow,
        
        /// <remarks/>
        curvedRightArrow,
        
        /// <remarks/>
        curvedLeftArrow,
        
        /// <remarks/>
        curvedUpArrow,
        
        /// <remarks/>
        curvedDownArrow,
        
        /// <remarks/>
        swooshArrow,
        
        /// <remarks/>
        cube,
        
        /// <remarks/>
        can,
        
        /// <remarks/>
        lightningBolt,
        
        /// <remarks/>
        heart,
        
        /// <remarks/>
        sun,
        
        /// <remarks/>
        moon,
        
        /// <remarks/>
        smileyFace,
        
        /// <remarks/>
        irregularSeal1,
        
        /// <remarks/>
        irregularSeal2,
        
        /// <remarks/>
        foldedCorner,
        
        /// <remarks/>
        bevel,
        
        /// <remarks/>
        frame,
        
        /// <remarks/>
        halfFrame,
        
        /// <remarks/>
        corner,
        
        /// <remarks/>
        diagStripe,
        
        /// <remarks/>
        chord,
        
        /// <remarks/>
        arc,
        
        /// <remarks/>
        leftBracket,
        
        /// <remarks/>
        rightBracket,
        
        /// <remarks/>
        leftBrace,
        
        /// <remarks/>
        rightBrace,
        
        /// <remarks/>
        bracketPair,
        
        /// <remarks/>
        bracePair,
        
        /// <remarks/>
        straightConnector1,
        
        /// <remarks/>
        bentConnector2,
        
        /// <remarks/>
        bentConnector3,
        
        /// <remarks/>
        bentConnector4,
        
        /// <remarks/>
        bentConnector5,
        
        /// <remarks/>
        curvedConnector2,
        
        /// <remarks/>
        curvedConnector3,
        
        /// <remarks/>
        curvedConnector4,
        
        /// <remarks/>
        curvedConnector5,
        
        /// <remarks/>
        callout1,
        
        /// <remarks/>
        callout2,
        
        /// <remarks/>
        callout3,
        
        /// <remarks/>
        accentCallout1,
        
        /// <remarks/>
        accentCallout2,
        
        /// <remarks/>
        accentCallout3,
        
        /// <remarks/>
        borderCallout1,
        
        /// <remarks/>
        borderCallout2,
        
        /// <remarks/>
        borderCallout3,
        
        /// <remarks/>
        accentBorderCallout1,
        
        /// <remarks/>
        accentBorderCallout2,
        
        /// <remarks/>
        accentBorderCallout3,
        
        /// <remarks/>
        wedgeRectCallout,
        
        /// <remarks/>
        wedgeRoundRectCallout,
        
        /// <remarks/>
        wedgeEllipseCallout,
        
        /// <remarks/>
        cloudCallout,
        
        /// <remarks/>
        cloud,
        
        /// <remarks/>
        ribbon,
        
        /// <remarks/>
        ribbon2,
        
        /// <remarks/>
        ellipseRibbon,
        
        /// <remarks/>
        ellipseRibbon2,
        
        /// <remarks/>
        leftRightRibbon,
        
        /// <remarks/>
        verticalScroll,
        
        /// <remarks/>
        horizontalScroll,
        
        /// <remarks/>
        wave,
        
        /// <remarks/>
        doubleWave,
        
        /// <remarks/>
        plus,
        
        /// <remarks/>
        flowChartProcess,
        
        /// <remarks/>
        flowChartDecision,
        
        /// <remarks/>
        flowChartInputOutput,
        
        /// <remarks/>
        flowChartPredefinedProcess,
        
        /// <remarks/>
        flowChartInternalStorage,
        
        /// <remarks/>
        flowChartDocument,
        
        /// <remarks/>
        flowChartMultidocument,
        
        /// <remarks/>
        flowChartTerminator,
        
        /// <remarks/>
        flowChartPreparation,
        
        /// <remarks/>
        flowChartManualInput,
        
        /// <remarks/>
        flowChartManualOperation,
        
        /// <remarks/>
        flowChartConnector,
        
        /// <remarks/>
        flowChartPunchedCard,
        
        /// <remarks/>
        flowChartPunchedTape,
        
        /// <remarks/>
        flowChartSummingJunction,
        
        /// <remarks/>
        flowChartOr,
        
        /// <remarks/>
        flowChartCollate,
        
        /// <remarks/>
        flowChartSort,
        
        /// <remarks/>
        flowChartExtract,
        
        /// <remarks/>
        flowChartMerge,
        
        /// <remarks/>
        flowChartOfflineStorage,
        
        /// <remarks/>
        flowChartOnlineStorage,
        
        /// <remarks/>
        flowChartMagneticTape,
        
        /// <remarks/>
        flowChartMagneticDisk,
        
        /// <remarks/>
        flowChartMagneticDrum,
        
        /// <remarks/>
        flowChartDisplay,
        
        /// <remarks/>
        flowChartDelay,
        
        /// <remarks/>
        flowChartAlternateProcess,
        
        /// <remarks/>
        flowChartOffpageConnector,
        
        /// <remarks/>
        actionButtonBlank,
        
        /// <remarks/>
        actionButtonHome,
        
        /// <remarks/>
        actionButtonHelp,
        
        /// <remarks/>
        actionButtonInformation,
        
        /// <remarks/>
        actionButtonForwardNext,
        
        /// <remarks/>
        actionButtonBackPrevious,
        
        /// <remarks/>
        actionButtonEnd,
        
        /// <remarks/>
        actionButtonBeginning,
        
        /// <remarks/>
        actionButtonReturn,
        
        /// <remarks/>
        actionButtonDocument,
        
        /// <remarks/>
        actionButtonSound,
        
        /// <remarks/>
        actionButtonMovie,
        
        /// <remarks/>
        gear6,
        
        /// <remarks/>
        gear9,
        
        /// <remarks/>
        funnel,
        
        /// <remarks/>
        mathPlus,
        
        /// <remarks/>
        mathMinus,
        
        /// <remarks/>
        mathMultiply,
        
        /// <remarks/>
        mathDivide,
        
        /// <remarks/>
        mathEqual,
        
        /// <remarks/>
        mathNotEqual,
        
        /// <remarks/>
        cornerTabs,
        
        /// <remarks/>
        squareTabs,
        
        /// <remarks/>
        plaqueTabs,
        
        /// <remarks/>
        chartX,
        
        /// <remarks/>
        chartStar,
        
        /// <remarks/>
        chartPlus,
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_TextShapeType {
        
        /// <remarks/>
        textNoShape,
        
        /// <remarks/>
        textPlain,
        
        /// <remarks/>
        textStop,
        
        /// <remarks/>
        textTriangle,
        
        /// <remarks/>
        textTriangleInverted,
        
        /// <remarks/>
        textChevron,
        
        /// <remarks/>
        textChevronInverted,
        
        /// <remarks/>
        textRingInside,
        
        /// <remarks/>
        textRingOutside,
        
        /// <remarks/>
        textArchUp,
        
        /// <remarks/>
        textArchDown,
        
        /// <remarks/>
        textCircle,
        
        /// <remarks/>
        textButton,
        
        /// <remarks/>
        textArchUpPour,
        
        /// <remarks/>
        textArchDownPour,
        
        /// <remarks/>
        textCirclePour,
        
        /// <remarks/>
        textButtonPour,
        
        /// <remarks/>
        textCurveUp,
        
        /// <remarks/>
        textCurveDown,
        
        /// <remarks/>
        textCanUp,
        
        /// <remarks/>
        textCanDown,
        
        /// <remarks/>
        textWave1,
        
        /// <remarks/>
        textWave2,
        
        /// <remarks/>
        textDoubleWave1,
        
        /// <remarks/>
        textWave4,
        
        /// <remarks/>
        textInflate,
        
        /// <remarks/>
        textDeflate,
        
        /// <remarks/>
        textInflateBottom,
        
        /// <remarks/>
        textDeflateBottom,
        
        /// <remarks/>
        textInflateTop,
        
        /// <remarks/>
        textDeflateTop,
        
        /// <remarks/>
        textDeflateInflate,
        
        /// <remarks/>
        textDeflateInflateDeflate,
        
        /// <remarks/>
        textFadeRight,
        
        /// <remarks/>
        textFadeLeft,
        
        /// <remarks/>
        textFadeUp,
        
        /// <remarks/>
        textFadeDown,
        
        /// <remarks/>
        textSlantUp,
        
        /// <remarks/>
        textSlantDown,
        
        /// <remarks/>
        textCascadeUp,
        
        /// <remarks/>
        textCascadeDown,
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GeomGuide {
        
        private string nameField;
        
        private string fmlaField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="token")]
        public string name {
            get {
                return this.nameField;
            }
            set {
                this.nameField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string fmla {
            get {
                return this.fmlaField;
            }
            set {
                this.fmlaField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_Path2DCubicBezierTo {
        
        private CT_AdjPoint2D[] ptField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("pt")]
        public CT_AdjPoint2D[] pt {
            get {
                return this.ptField;
            }
            set {
                this.ptField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AdjPoint2D {
        
        private string xField;
        
        private string yField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string x {
            get {
                return this.xField;
            }
            set {
                this.xField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string y {
            get {
                return this.yField;
            }
            set {
                this.yField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_Path2DQuadBezierTo {
        
        private CT_AdjPoint2D[] ptField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("pt")]
        public CT_AdjPoint2D[] pt {
            get {
                return this.ptField;
            }
            set {
                this.ptField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GeomGuideList {
        
        private CT_GeomGuide[] gdField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("gd")]
        public CT_GeomGuide[] gd {
            get {
                return this.gdField;
            }
            set {
                this.gdField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GeomRect {
        
        private string lField;
        
        private string tField;
        
        private string rField;
        
        private string bField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string l {
            get {
                return this.lField;
            }
            set {
                this.lField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string t {
            get {
                return this.tField;
            }
            set {
                this.tField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string r {
            get {
                return this.rField;
            }
            set {
                this.rField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string b {
            get {
                return this.bField;
            }
            set {
                this.bField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_XYAdjustHandle {
        
        private CT_AdjPoint2D posField;
        
        private string gdRefXField;
        
        private string minXField;
        
        private string maxXField;
        
        private string gdRefYField;
        
        private string minYField;
        
        private string maxYField;
        
        /// <remarks/>
        public CT_AdjPoint2D pos {
            get {
                return this.posField;
            }
            set {
                this.posField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="token")]
        public string gdRefX {
            get {
                return this.gdRefXField;
            }
            set {
                this.gdRefXField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string minX {
            get {
                return this.minXField;
            }
            set {
                this.minXField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string maxX {
            get {
                return this.maxXField;
            }
            set {
                this.maxXField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="token")]
        public string gdRefY {
            get {
                return this.gdRefYField;
            }
            set {
                this.gdRefYField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string minY {
            get {
                return this.minYField;
            }
            set {
                this.minYField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string maxY {
            get {
                return this.maxYField;
            }
            set {
                this.maxYField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_PolarAdjustHandle {
        
        private CT_AdjPoint2D posField;
        
        private string gdRefRField;
        
        private string minRField;
        
        private string maxRField;
        
        private string gdRefAngField;
        
        private string minAngField;
        
        private string maxAngField;
        
        /// <remarks/>
        public CT_AdjPoint2D pos {
            get {
                return this.posField;
            }
            set {
                this.posField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="token")]
        public string gdRefR {
            get {
                return this.gdRefRField;
            }
            set {
                this.gdRefRField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string minR {
            get {
                return this.minRField;
            }
            set {
                this.minRField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string maxR {
            get {
                return this.maxRField;
            }
            set {
                this.maxRField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="token")]
        public string gdRefAng {
            get {
                return this.gdRefAngField;
            }
            set {
                this.gdRefAngField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string minAng {
            get {
                return this.minAngField;
            }
            set {
                this.minAngField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string maxAng {
            get {
                return this.maxAngField;
            }
            set {
                this.maxAngField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_ConnectionSite {
        
        private CT_AdjPoint2D posField;
        
        private string angField;
        
        /// <remarks/>
        public CT_AdjPoint2D pos {
            get {
                return this.posField;
            }
            set {
                this.posField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string ang {
            get {
                return this.angField;
            }
            set {
                this.angField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AdjustHandleList {
        
        private object[] itemsField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("ahPolar", typeof(CT_PolarAdjustHandle))]
        [System.Xml.Serialization.XmlElement("ahXY", typeof(CT_XYAdjustHandle))]
        public object[] Items {
            get {
                return this.itemsField;
            }
            set {
                this.itemsField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_ConnectionSiteList {
        
        private CT_ConnectionSite[] cxnField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("cxn")]
        public CT_ConnectionSite[] cxn {
            get {
                return this.cxnField;
            }
            set {
                this.cxnField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_Connection {
        
        private uint idField;
        
        private uint idxField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public uint id {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public uint idx {
            get {
                return this.idxField;
            }
            set {
                this.idxField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_Path2DMoveTo {
        
        private CT_AdjPoint2D ptField;
        
        /// <remarks/>
        public CT_AdjPoint2D pt {
            get {
                return this.ptField;
            }
            set {
                this.ptField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_Path2DLineTo {
        
        private CT_AdjPoint2D ptField;
        
        /// <remarks/>
        public CT_AdjPoint2D pt {
            get {
                return this.ptField;
            }
            set {
                this.ptField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_Path2DArcTo {
        
        private string wrField;
        
        private string hrField;
        
        private string stAngField;
        
        private string swAngField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string wR {
            get {
                return this.wrField;
            }
            set {
                this.wrField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string hR {
            get {
                return this.hrField;
            }
            set {
                this.hrField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string stAng {
            get {
                return this.stAngField;
            }
            set {
                this.stAngField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string swAng {
            get {
                return this.swAngField;
            }
            set {
                this.swAngField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_Path2DClose {
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_PathFillMode {
        
        /// <remarks/>
        none,
        
        /// <remarks/>
        norm,
        
        /// <remarks/>
        lighten,
        
        /// <remarks/>
        lightenLess,
        
        /// <remarks/>
        darken,
        
        /// <remarks/>
        darkenLess,
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_Path2D {
        
        private object[] itemsField;
        
        private ItemsChoiceType[] itemsElementNameField;
        
        private long wField;
        
        private long hField;
        
        private ST_PathFillMode fillField;
        
        private bool strokeField;
        
        private bool extrusionOkField;
        
        public CT_Path2D() {
            this.wField = ((long)(0));
            this.hField = ((long)(0));
            this.fillField = ST_PathFillMode.norm;
            this.strokeField = true;
            this.extrusionOkField = true;
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("arcTo", typeof(CT_Path2DArcTo))]
        [System.Xml.Serialization.XmlElement("close", typeof(CT_Path2DClose))]
        [System.Xml.Serialization.XmlElement("cubicBezTo", typeof(CT_Path2DCubicBezierTo))]
        [System.Xml.Serialization.XmlElement("lnTo", typeof(CT_Path2DLineTo))]
        [System.Xml.Serialization.XmlElement("moveTo", typeof(CT_Path2DMoveTo))]
        [System.Xml.Serialization.XmlElement("quadBezTo", typeof(CT_Path2DQuadBezierTo))]
        [System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemsElementName")]
        public object[] Items {
            get {
                return this.itemsField;
            }
            set {
                this.itemsField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("ItemsElementName")]
        [System.Xml.Serialization.XmlIgnore]
        public ItemsChoiceType[] ItemsElementName {
            get {
                return this.itemsElementNameField;
            }
            set {
                this.itemsElementNameField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [System.ComponentModel.DefaultValueAttribute(typeof(long), "0")]
        public long w {
            get {
                return this.wField;
            }
            set {
                this.wField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [System.ComponentModel.DefaultValueAttribute(typeof(long), "0")]
        public long h {
            get {
                return this.hField;
            }
            set {
                this.hField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [System.ComponentModel.DefaultValueAttribute(ST_PathFillMode.norm)]
        public ST_PathFillMode fill {
            get {
                return this.fillField;
            }
            set {
                this.fillField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool stroke {
            get {
                return this.strokeField;
            }
            set {
                this.strokeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool extrusionOk {
            get {
                return this.extrusionOkField;
            }
            set {
                this.extrusionOkField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IncludeInSchema=false)]
    public enum ItemsChoiceType {
        
        /// <remarks/>
        arcTo,
        
        /// <remarks/>
        close,
        
        /// <remarks/>
        cubicBezTo,
        
        /// <remarks/>
        lnTo,
        
        /// <remarks/>
        moveTo,
        
        /// <remarks/>
        quadBezTo,
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_Path2DList {
        
        private CT_Path2D[] pathField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("path")]
        public CT_Path2D[] path {
            get {
                return this.pathField;
            }
            set {
                this.pathField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_PresetGeometry2D {
        
        private List<CT_GeomGuide> avLstField;
        
        private ST_ShapeType prstField;

        public void AddNewAvLst()
        {
            this.avLstField = new List<CT_GeomGuide>();
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("gd", IsNullable=false)]
        public List<CT_GeomGuide> avLst {
            get {
                return this.avLstField;
            }
            set {
                this.avLstField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_ShapeType prst {
            get {
                return this.prstField;
            }
            set {
                this.prstField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_PresetTextShape {
        
        private CT_GeomGuide[] avLstField;
        
        private ST_TextShapeType prstField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("gd", IsNullable=false)]
        public CT_GeomGuide[] avLst {
            get {
                return this.avLstField;
            }
            set {
                this.avLstField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TextShapeType prst {
            get {
                return this.prstField;
            }
            set {
                this.prstField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_CustomGeometry2D {
        
        private CT_GeomGuide[] avLstField;
        
        private CT_GeomGuide[] gdLstField;
        
        private object[] ahLstField;
        
        private CT_ConnectionSite[] cxnLstField;
        
        private CT_GeomRect rectField;
        
        private CT_Path2D[] pathLstField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("gd", IsNullable=false)]
        public CT_GeomGuide[] avLst {
            get {
                return this.avLstField;
            }
            set {
                this.avLstField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("gd", IsNullable=false)]
        public CT_GeomGuide[] gdLst {
            get {
                return this.gdLstField;
            }
            set {
                this.gdLstField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("ahPolar", typeof(CT_PolarAdjustHandle), IsNullable=false)]
        [System.Xml.Serialization.XmlArrayItemAttribute("ahXY", typeof(CT_XYAdjustHandle), IsNullable=false)]
        public object[] ahLst {
            get {
                return this.ahLstField;
            }
            set {
                this.ahLstField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("cxn", IsNullable=false)]
        public CT_ConnectionSite[] cxnLst {
            get {
                return this.cxnLstField;
            }
            set {
                this.cxnLstField = value;
            }
        }
        
        /// <remarks/>
        public CT_GeomRect rect {
            get {
                return this.rectField;
            }
            set {
                this.rectField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("path", IsNullable=false)]
        public CT_Path2D[] pathLst {
            get {
                return this.pathLstField;
            }
            set {
                this.pathLstField = value;
            }
        }
    }
}
