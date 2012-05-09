using System.Collections.Generic;
using System.ComponentModel;
namespace NPOI.OpenXmlFormats.Dml {
    
    

    [System.Serializable]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_ShapeType {
        
    
        line,
        
    
        lineInv,
        
    
        triangle,
        
    
        rtTriangle,
        
    
        rect,
        
    
        diamond,
        
    
        parallelogram,
        
    
        trapezoid,
        
    
        nonIsoscelesTrapezoid,
        
    
        pentagon,
        
    
        hexagon,
        
    
        heptagon,
        
    
        octagon,
        
    
        decagon,
        
    
        dodecagon,
        
    
        star4,
        
    
        star5,
        
    
        star6,
        
    
        star7,
        
    
        star8,
        
    
        star10,
        
    
        star12,
        
    
        star16,
        
    
        star24,
        
    
        star32,
        
    
        roundRect,
        
    
        round1Rect,
        
    
        round2SameRect,
        
    
        round2DiagRect,
        
    
        snipRoundRect,
        
    
        snip1Rect,
        
    
        snip2SameRect,
        
    
        snip2DiagRect,
        
    
        plaque,
        
    
        ellipse,
        
    
        teardrop,
        
    
        homePlate,
        
    
        chevron,
        
    
        pieWedge,
        
    
        pie,
        
    
        blockArc,
        
    
        donut,
        
    
        noSmoking,
        
    
        rightArrow,
        
    
        leftArrow,
        
    
        upArrow,
        
    
        downArrow,
        
    
        stripedRightArrow,
        
    
        notchedRightArrow,
        
    
        bentUpArrow,
        
    
        leftRightArrow,
        
    
        upDownArrow,
        
    
        leftUpArrow,
        
    
        leftRightUpArrow,
        
    
        quadArrow,
        
    
        leftArrowCallout,
        
    
        rightArrowCallout,
        
    
        upArrowCallout,
        
    
        downArrowCallout,
        
    
        leftRightArrowCallout,
        
    
        upDownArrowCallout,
        
    
        quadArrowCallout,
        
    
        bentArrow,
        
    
        uturnArrow,
        
    
        circularArrow,
        
    
        leftCircularArrow,
        
    
        leftRightCircularArrow,
        
    
        curvedRightArrow,
        
    
        curvedLeftArrow,
        
    
        curvedUpArrow,
        
    
        curvedDownArrow,
        
    
        swooshArrow,
        
    
        cube,
        
    
        can,
        
    
        lightningBolt,
        
    
        heart,
        
    
        sun,
        
    
        moon,
        
    
        smileyFace,
        
    
        irregularSeal1,
        
    
        irregularSeal2,
        
    
        foldedCorner,
        
    
        bevel,
        
    
        frame,
        
    
        halfFrame,
        
    
        corner,
        
    
        diagStripe,
        
    
        chord,
        
    
        arc,
        
    
        leftBracket,
        
    
        rightBracket,
        
    
        leftBrace,
        
    
        rightBrace,
        
    
        bracketPair,
        
    
        bracePair,
        
    
        straightConnector1,
        
    
        bentConnector2,
        
    
        bentConnector3,
        
    
        bentConnector4,
        
    
        bentConnector5,
        
    
        curvedConnector2,
        
    
        curvedConnector3,
        
    
        curvedConnector4,
        
    
        curvedConnector5,
        
    
        callout1,
        
    
        callout2,
        
    
        callout3,
        
    
        accentCallout1,
        
    
        accentCallout2,
        
    
        accentCallout3,
        
    
        borderCallout1,
        
    
        borderCallout2,
        
    
        borderCallout3,
        
    
        accentBorderCallout1,
        
    
        accentBorderCallout2,
        
    
        accentBorderCallout3,
        
    
        wedgeRectCallout,
        
    
        wedgeRoundRectCallout,
        
    
        wedgeEllipseCallout,
        
    
        cloudCallout,
        
    
        cloud,
        
    
        ribbon,
        
    
        ribbon2,
        
    
        ellipseRibbon,
        
    
        ellipseRibbon2,
        
    
        leftRightRibbon,
        
    
        verticalScroll,
        
    
        horizontalScroll,
        
    
        wave,
        
    
        doubleWave,
        
    
        plus,
        
    
        flowChartProcess,
        
    
        flowChartDecision,
        
    
        flowChartInputOutput,
        
    
        flowChartPredefinedProcess,
        
    
        flowChartInternalStorage,
        
    
        flowChartDocument,
        
    
        flowChartMultidocument,
        
    
        flowChartTerminator,
        
    
        flowChartPreparation,
        
    
        flowChartManualInput,
        
    
        flowChartManualOperation,
        
    
        flowChartConnector,
        
    
        flowChartPunchedCard,
        
    
        flowChartPunchedTape,
        
    
        flowChartSummingJunction,
        
    
        flowChartOr,
        
    
        flowChartCollate,
        
    
        flowChartSort,
        
    
        flowChartExtract,
        
    
        flowChartMerge,
        
    
        flowChartOfflineStorage,
        
    
        flowChartOnlineStorage,
        
    
        flowChartMagneticTape,
        
    
        flowChartMagneticDisk,
        
    
        flowChartMagneticDrum,
        
    
        flowChartDisplay,
        
    
        flowChartDelay,
        
    
        flowChartAlternateProcess,
        
    
        flowChartOffpageConnector,
        
    
        actionButtonBlank,
        
    
        actionButtonHome,
        
    
        actionButtonHelp,
        
    
        actionButtonInformation,
        
    
        actionButtonForwardNext,
        
    
        actionButtonBackPrevious,
        
    
        actionButtonEnd,
        
    
        actionButtonBeginning,
        
    
        actionButtonReturn,
        
    
        actionButtonDocument,
        
    
        actionButtonSound,
        
    
        actionButtonMovie,
        
    
        gear6,
        
    
        gear9,
        
    
        funnel,
        
    
        mathPlus,
        
    
        mathMinus,
        
    
        mathMultiply,
        
    
        mathDivide,
        
    
        mathEqual,
        
    
        mathNotEqual,
        
    
        cornerTabs,
        
    
        squareTabs,
        
    
        plaqueTabs,
        
    
        chartX,
        
    
        chartStar,
        
    
        chartPlus,
    }
    

    [System.Serializable]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_TextShapeType {
        
    
        textNoShape,
        
    
        textPlain,
        
    
        textStop,
        
    
        textTriangle,
        
    
        textTriangleInverted,
        
    
        textChevron,
        
    
        textChevronInverted,
        
    
        textRingInside,
        
    
        textRingOutside,
        
    
        textArchUp,
        
    
        textArchDown,
        
    
        textCircle,
        
    
        textButton,
        
    
        textArchUpPour,
        
    
        textArchDownPour,
        
    
        textCirclePour,
        
    
        textButtonPour,
        
    
        textCurveUp,
        
    
        textCurveDown,
        
    
        textCanUp,
        
    
        textCanDown,
        
    
        textWave1,
        
    
        textWave2,
        
    
        textDoubleWave1,
        
    
        textWave4,
        
    
        textInflate,
        
    
        textDeflate,
        
    
        textInflateBottom,
        
    
        textDeflateBottom,
        
    
        textInflateTop,
        
    
        textDeflateTop,
        
    
        textDeflateInflate,
        
    
        textDeflateInflateDeflate,
        
    
        textFadeRight,
        
    
        textFadeLeft,
        
    
        textFadeUp,
        
    
        textFadeDown,
        
    
        textSlantUp,
        
    
        textSlantDown,
        
    
        textCascadeUp,
        
    
        textCascadeDown,
    }
    

    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GeomGuide {
        
        private string nameField;
        
        private string fmlaField;
        
    
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="token")]
        public string name {
            get {
                return this.nameField;
            }
            set {
                this.nameField = value;
            }
        }
        
    
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
    

    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_Path2DCubicBezierTo {
        
        private CT_AdjPoint2D[] ptField;
        
    
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
    

    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AdjPoint2D {
        
        private string xField;
        
        private string yField;
        
    
        [System.Xml.Serialization.XmlAttribute]
        public string x {
            get {
                return this.xField;
            }
            set {
                this.xField = value;
            }
        }
        
    
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
    

    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_Path2DQuadBezierTo {
        
        private CT_AdjPoint2D[] ptField;
        
    
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
    

    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GeomGuideList {
        
        private CT_GeomGuide[] gdField;
        
    
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
        
    
        [System.Xml.Serialization.XmlAttribute]
        public string l {
            get {
                return this.lField;
            }
            set {
                this.lField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttribute]
        public string t {
            get {
                return this.tField;
            }
            set {
                this.tField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttribute]
        public string r {
            get {
                return this.rField;
            }
            set {
                this.rField = value;
            }
        }
        
    
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
        
    
        public CT_AdjPoint2D pos {
            get {
                return this.posField;
            }
            set {
                this.posField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="token")]
        public string gdRefX {
            get {
                return this.gdRefXField;
            }
            set {
                this.gdRefXField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttribute]
        public string minX {
            get {
                return this.minXField;
            }
            set {
                this.minXField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttribute]
        public string maxX {
            get {
                return this.maxXField;
            }
            set {
                this.maxXField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="token")]
        public string gdRefY {
            get {
                return this.gdRefYField;
            }
            set {
                this.gdRefYField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttribute]
        public string minY {
            get {
                return this.minYField;
            }
            set {
                this.minYField = value;
            }
        }
        
    
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
        
    
        public CT_AdjPoint2D pos {
            get {
                return this.posField;
            }
            set {
                this.posField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="token")]
        public string gdRefR {
            get {
                return this.gdRefRField;
            }
            set {
                this.gdRefRField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttribute]
        public string minR {
            get {
                return this.minRField;
            }
            set {
                this.minRField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttribute]
        public string maxR {
            get {
                return this.maxRField;
            }
            set {
                this.maxRField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="token")]
        public string gdRefAng {
            get {
                return this.gdRefAngField;
            }
            set {
                this.gdRefAngField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttribute]
        public string minAng {
            get {
                return this.minAngField;
            }
            set {
                this.minAngField = value;
            }
        }
        
    
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
    

    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_ConnectionSite {
        
        private CT_AdjPoint2D posField;
        
        private string angField;
        
    
        public CT_AdjPoint2D pos {
            get {
                return this.posField;
            }
            set {
                this.posField = value;
            }
        }
        
    
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
    

    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AdjustHandleList {
        
        private object[] itemsField;
        
    
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
    

    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_ConnectionSiteList {
        
        private CT_ConnectionSite[] cxnField;
        
    
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
    

    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_Connection {
        
        private uint idField;
        
        private uint idxField;
        
    
        [System.Xml.Serialization.XmlAttribute]
        public uint id {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
        }
        
    
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
    

    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_Path2DMoveTo {
        
        private CT_AdjPoint2D ptField;
        
    
        public CT_AdjPoint2D pt {
            get {
                return this.ptField;
            }
            set {
                this.ptField = value;
            }
        }
    }
    

    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_Path2DLineTo {
        
        private CT_AdjPoint2D ptField;
        
    
        public CT_AdjPoint2D pt {
            get {
                return this.ptField;
            }
            set {
                this.ptField = value;
            }
        }
    }
    

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
        
    
        [System.Xml.Serialization.XmlAttribute]
        public string wR {
            get {
                return this.wrField;
            }
            set {
                this.wrField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttribute]
        public string hR {
            get {
                return this.hrField;
            }
            set {
                this.hrField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttribute]
        public string stAng {
            get {
                return this.stAngField;
            }
            set {
                this.stAngField = value;
            }
        }
        
    
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
    

    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_Path2DClose {
    }
    

    [System.Serializable]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_PathFillMode {
        
    
        none,
        
    
        norm,
        
    
        lighten,
        
    
        lightenLess,
        
    
        darken,
        
    
        darkenLess,
    }
    

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
        
    
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long w {
            get {
                return this.wField;
            }
            set {
                this.wField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long h {
            get {
                return this.hField;
            }
            set {
                this.hField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(ST_PathFillMode.norm)]
        public ST_PathFillMode fill {
            get {
                return this.fillField;
            }
            set {
                this.fillField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(true)]
        public bool stroke {
            get {
                return this.strokeField;
            }
            set {
                this.strokeField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(true)]
        public bool extrusionOk {
            get {
                return this.extrusionOkField;
            }
            set {
                this.extrusionOkField = value;
            }
        }
    }
    

    [System.Serializable]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IncludeInSchema=false)]
    public enum ItemsChoiceType {
        
    
        arcTo,
        
    
        close,
        
    
        cubicBezTo,
        
    
        lnTo,
        
    
        moveTo,
        
    
        quadBezTo,
    }
    

    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_Path2DList {
        
        private CT_Path2D[] pathField;
        
    
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
        
    
        [System.Xml.Serialization.XmlArrayItemAttribute("gd", IsNullable=false)]
        public List<CT_GeomGuide> avLst {
            get {
                return this.avLstField;
            }
            set {
                this.avLstField = value;
            }
        }
        
    
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
    

    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_PresetTextShape {
        
        private CT_GeomGuide[] avLstField;
        
        private ST_TextShapeType prstField;
        
    
        [System.Xml.Serialization.XmlArrayItemAttribute("gd", IsNullable=false)]
        public CT_GeomGuide[] avLst {
            get {
                return this.avLstField;
            }
            set {
                this.avLstField = value;
            }
        }
        
    
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
        
    
        [System.Xml.Serialization.XmlArrayItemAttribute("gd", IsNullable=false)]
        public CT_GeomGuide[] avLst {
            get {
                return this.avLstField;
            }
            set {
                this.avLstField = value;
            }
        }
        
    
        [System.Xml.Serialization.XmlArrayItemAttribute("gd", IsNullable=false)]
        public CT_GeomGuide[] gdLst {
            get {
                return this.gdLstField;
            }
            set {
                this.gdLstField = value;
            }
        }
        
    
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
        
    
        [System.Xml.Serialization.XmlArrayItemAttribute("cxn", IsNullable=false)]
        public CT_ConnectionSite[] cxnLst {
            get {
                return this.cxnLstField;
            }
            set {
                this.cxnLstField = value;
            }
        }
        
    
        public CT_GeomRect rect {
            get {
                return this.rectField;
            }
            set {
                this.rectField = value;
            }
        }
        
    
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
