using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Xml.Serialization;
using System.Diagnostics;
using System.Xml;
using NPOI.OpenXml4Net.Util;
using System.IO;

namespace NPOI.OpenXmlFormats.Dml
{



    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_ShapeType : int
    {
        none,

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


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_TextShapeType
    {


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


    [Serializable]
    [DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_GeomGuide
    {

        private string nameField;

        private string fmlaField;

        public static CT_GeomGuide Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_GeomGuide ctObj = new CT_GeomGuide();
            ctObj.name = XmlHelper.ReadString(node.Attributes["name"]);
            ctObj.fmla = XmlHelper.ReadString(node.Attributes["fmla"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "name", this.name);
            XmlHelper.WriteAttribute(sw, "fmla", this.fmla);
            sw.Write("/>");
        }

        [XmlAttribute(DataType = "token")]
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


        [XmlAttribute]
        public string fmla
        {
            get
            {
                return this.fmlaField;
            }
            set
            {
                this.fmlaField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_Path2DCubicBezierTo
    {

        private CT_AdjPoint2D[] ptField;


        [XmlElement("pt", Order = 0)]
        public CT_AdjPoint2D[] pt
        {
            get
            {
                return this.ptField;
            }
            set
            {
                this.ptField = value;
            }
        }
    }


    [Serializable]
    [DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_AdjPoint2D
    {

        private string xField;

        private string yField;

        public static CT_AdjPoint2D Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_AdjPoint2D ctObj = new CT_AdjPoint2D();
            ctObj.x = XmlHelper.ReadString(node.Attributes["x"]);
            ctObj.y = XmlHelper.ReadString(node.Attributes["y"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "x", this.x);
            XmlHelper.WriteAttribute(sw, "y", this.y);
            sw.Write(">");
            sw.Write(string.Format("</a:{0}>", nodeName));
        }

        [XmlAttribute]
        public string x
        {
            get
            {
                return this.xField;
            }
            set
            {
                this.xField = value;
            }
        }


        [XmlAttribute]
        public string y
        {
            get
            {
                return this.yField;
            }
            set
            {
                this.yField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_Path2DQuadBezierTo
    {

        private CT_AdjPoint2D[] ptField;


        [XmlElement("pt", Order = 0)]
        public CT_AdjPoint2D[] pt
        {
            get
            {
                return this.ptField;
            }
            set
            {
                this.ptField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_GeomGuideList
    {

        private List<CT_GeomGuide> gdField;


        [XmlElement("gd", Order = 0)]
        public List<CT_GeomGuide> gd
        {
            get
            {
                return this.gdField;
            }
            set
            {
                this.gdField = value;
            }
        }

        internal static CT_GeomGuideList Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_GeomGuideList avLst = new CT_GeomGuideList();
            avLst.gdField = new List<CT_GeomGuide>();
            if (node.ChildNodes != null)
            {
                foreach (XmlNode childNode in node.ChildNodes)
                {
                    if (childNode.LocalName == "gd")
                        avLst.gdField.Add(CT_GeomGuide.Parse(childNode, namespaceManager));
                }
            }
            return avLst;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write("<a:{0}>", nodeName);
            if (this.gdField != null)
            {
                foreach (CT_GeomGuide gg in gdField)
                {
                    gg.Write(sw, "gd");
                }
            }
            sw.Write("</a:{0}>", nodeName);

        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_GeomRect
    {

        private string lField;

        private string tField;

        private string rField;

        private string bField;

        public static CT_GeomRect Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_GeomRect ctObj = new CT_GeomRect();
            ctObj.l = XmlHelper.ReadString(node.Attributes["l"]);
            ctObj.t = XmlHelper.ReadString(node.Attributes["t"]);
            ctObj.r = XmlHelper.ReadString(node.Attributes["r"]);
            ctObj.b = XmlHelper.ReadString(node.Attributes["b"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "l", this.l);
            XmlHelper.WriteAttribute(sw, "t", this.t);
            XmlHelper.WriteAttribute(sw, "r", this.r);
            XmlHelper.WriteAttribute(sw, "b", this.b);
            sw.Write("/>");
        }

        [XmlAttribute]
        public string l
        {
            get
            {
                return this.lField;
            }
            set
            {
                this.lField = value;
            }
        }


        [XmlAttribute]
        public string t
        {
            get
            {
                return this.tField;
            }
            set
            {
                this.tField = value;
            }
        }


        [XmlAttribute]
        public string r
        {
            get
            {
                return this.rField;
            }
            set
            {
                this.rField = value;
            }
        }


        [XmlAttribute]
        public string b
        {
            get
            {
                return this.bField;
            }
            set
            {
                this.bField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_XYAdjustHandle
    {

        private CT_AdjPoint2D posField;

        private string gdRefXField;

        private string minXField;

        private string maxXField;

        private string gdRefYField;

        private string minYField;

        private string maxYField;

        [XmlElement(Order = 0)]
        public CT_AdjPoint2D pos
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


        [XmlAttribute(DataType = "token")]
        public string gdRefX
        {
            get
            {
                return this.gdRefXField;
            }
            set
            {
                this.gdRefXField = value;
            }
        }


        [XmlAttribute]
        public string minX
        {
            get
            {
                return this.minXField;
            }
            set
            {
                this.minXField = value;
            }
        }


        [XmlAttribute]
        public string maxX
        {
            get
            {
                return this.maxXField;
            }
            set
            {
                this.maxXField = value;
            }
        }


        [XmlAttribute(DataType = "token")]
        public string gdRefY
        {
            get
            {
                return this.gdRefYField;
            }
            set
            {
                this.gdRefYField = value;
            }
        }


        [XmlAttribute]
        public string minY
        {
            get
            {
                return this.minYField;
            }
            set
            {
                this.minYField = value;
            }
        }


        [XmlAttribute]
        public string maxY
        {
            get
            {
                return this.maxYField;
            }
            set
            {
                this.maxYField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_PolarAdjustHandle
    {

        private CT_AdjPoint2D posField;

        private string gdRefRField;

        private string minRField;

        private string maxRField;

        private string gdRefAngField;

        private string minAngField;

        private string maxAngField;

        [XmlElement(Order = 0)]
        public CT_AdjPoint2D pos
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


        [XmlAttribute(DataType = "token")]
        public string gdRefR
        {
            get
            {
                return this.gdRefRField;
            }
            set
            {
                this.gdRefRField = value;
            }
        }


        [XmlAttribute]
        public string minR
        {
            get
            {
                return this.minRField;
            }
            set
            {
                this.minRField = value;
            }
        }


        [XmlAttribute]
        public string maxR
        {
            get
            {
                return this.maxRField;
            }
            set
            {
                this.maxRField = value;
            }
        }


        [XmlAttribute(DataType = "token")]
        public string gdRefAng
        {
            get
            {
                return this.gdRefAngField;
            }
            set
            {
                this.gdRefAngField = value;
            }
        }


        [XmlAttribute]
        public string minAng
        {
            get
            {
                return this.minAngField;
            }
            set
            {
                this.minAngField = value;
            }
        }


        [XmlAttribute]
        public string maxAng
        {
            get
            {
                return this.maxAngField;
            }
            set
            {
                this.maxAngField = value;
            }
        }
    }


    [Serializable]
    [DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_ConnectionSite
    {

        private CT_AdjPoint2D posField;

        private string angField;
        public static CT_ConnectionSite Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ConnectionSite ctObj = new CT_ConnectionSite();
            ctObj.ang = XmlHelper.ReadString(node.Attributes["ang"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "pos")
                    ctObj.pos = CT_AdjPoint2D.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "ang", this.ang);
            sw.Write(">");
            if (this.pos != null)
                this.pos.Write(sw, "pos");
            sw.Write(string.Format("</a:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_AdjPoint2D pos
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


        [XmlAttribute]
        public string ang
        {
            get
            {
                return this.angField;
            }
            set
            {
                this.angField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_AdjustHandleList
    {

        private object[] itemsField;


        [XmlElement("ahPolar", typeof(CT_PolarAdjustHandle), Order = 0)]
        [XmlElement("ahXY", typeof(CT_XYAdjustHandle), Order = 0)]
        public object[] Items
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
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_ConnectionSiteList
    {

        private List<CT_ConnectionSite> cxnField;


        //[XmlElement("cxn", Order = 0)]
        public List<CT_ConnectionSite> cxn
        {
            get
            {
                return this.cxnField;
            }
            set
            {
                this.cxnField = value;
            }
        }
        internal static CT_ConnectionSiteList Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_ConnectionSiteList cxnLst = new CT_ConnectionSiteList();
            cxnLst.cxnField = new List<CT_ConnectionSite>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "cxn")
                    cxnLst.cxnField.Add(CT_ConnectionSite.Parse(childNode, namespaceManager));
            }
            return cxnLst;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write("<a:{0}>", nodeName);
            if (this.cxnField != null)
            {
                foreach (CT_ConnectionSite gg in cxnField)
                {
                    gg.Write(sw, "cxn");
                }
            }
            sw.Write("</a:{0}>", nodeName);

        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_Connection
    {

        private uint idField;

        private uint idxField;
        public static CT_Connection Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Connection ctObj = new CT_Connection();
            ctObj.id = XmlHelper.ReadUInt(node.Attributes["id"]);
            ctObj.idx = XmlHelper.ReadUInt(node.Attributes["idx"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "id", this.id);
            XmlHelper.WriteAttribute(sw, "idx", this.idx);
            sw.Write("/>");
        }


        [XmlAttribute]
        public uint id
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


        [XmlAttribute]
        public uint idx
        {
            get
            {
                return this.idxField;
            }
            set
            {
                this.idxField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_Path2DMoveTo
    {

        private CT_AdjPoint2D ptField;
        public CT_Path2DMoveTo()
        {
            this.ptField = new CT_AdjPoint2D();
        }
        [XmlElement(Order = 0)]
        public CT_AdjPoint2D pt
        {
            get
            {
                return this.ptField;
            }
            set
            {
                this.ptField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_Path2DLineTo
    {

        private CT_AdjPoint2D ptField;

        public CT_Path2DLineTo()
        {
            this.ptField = new CT_AdjPoint2D();
        }

        [XmlElement(Order = 0)]
        public CT_AdjPoint2D pt
        {
            get
            {
                return this.ptField;
            }
            set
            {
                this.ptField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_Path2DArcTo
    {

        private string wrField;

        private string hrField;

        private string stAngField;

        private string swAngField;


        [XmlAttribute]
        public string wR
        {
            get
            {
                return this.wrField;
            }
            set
            {
                this.wrField = value;
            }
        }


        [XmlAttribute]
        public string hR
        {
            get
            {
                return this.hrField;
            }
            set
            {
                this.hrField = value;
            }
        }


        [XmlAttribute]
        public string stAng
        {
            get
            {
                return this.stAngField;
            }
            set
            {
                this.stAngField = value;
            }
        }


        [XmlAttribute]
        public string swAng
        {
            get
            {
                return this.swAngField;
            }
            set
            {
                this.swAngField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_Path2DClose
    {
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_PathFillMode
    {


        none,


        norm,


        lighten,


        lightenLess,


        darken,


        darkenLess,
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_Path2D
    {

        //private object[] itemsField;

        //private ItemsChoiceType[] itemsElementNameField;

        private long wField;

        private long hField;

        private ST_PathFillMode fillField;

        private bool strokeField;

        private bool extrusionOkField;
        public static CT_Path2D Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Path2D ctObj = new CT_Path2D();
            ctObj.w = XmlHelper.ReadLong(node.Attributes["w"]);
            ctObj.h = XmlHelper.ReadLong(node.Attributes["h"]);
            if (node.Attributes["fill"] != null)
                ctObj.fill = (ST_PathFillMode)Enum.Parse(typeof(ST_PathFillMode), node.Attributes["fill"].Value);
            ctObj.stroke = XmlHelper.ReadBool(node.Attributes["stroke"]);
            ctObj.extrusionOk = XmlHelper.ReadBool(node.Attributes["extrusionOk"]);
            //foreach(XmlNode childNode in node.ChildNodes)
            //{
            //    if(childNode.LocalName == "ItemsElementName")
            //        ctObj.ItemsElementName = ItemsChoiceType[].Parse(childNode, namespaceManager);
            //}
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w", this.w);
            XmlHelper.WriteAttribute(sw, "h", this.h);
            XmlHelper.WriteAttribute(sw, "fill", this.fill.ToString());
            XmlHelper.WriteAttribute(sw, "stroke", this.stroke);
            XmlHelper.WriteAttribute(sw, "extrusionOk", this.extrusionOk);
            sw.Write(">");
            //if (this.ItemsElementName != null)
            //    this.ItemsElementName.Write(sw, "ItemsElementName");
            sw.Write(string.Format("</a:{0}>", nodeName));
        }

        public CT_Path2D()
        {
            this.wField = ((long)(0));
            this.hField = ((long)(0));
            this.fillField = ST_PathFillMode.norm;
            this.strokeField = true;
            this.extrusionOkField = true;
        }


        //[XmlElement("arcTo", typeof(CT_Path2DArcTo))]
        //[XmlElement("close", typeof(CT_Path2DClose))]
        //[XmlElement("cubicBezTo", typeof(CT_Path2DCubicBezierTo))]
        //[XmlElement("lnTo", typeof(CT_Path2DLineTo))]
        //[XmlElement("moveTo", typeof(CT_Path2DMoveTo))]
        //[XmlElement("quadBezTo", typeof(CT_Path2DQuadBezierTo))]
        //[XmlChoiceIdentifier("ItemsElementName")]
        //public object[] Items
        //{
        //    get
        //    {
        //        return this.itemsField;
        //    }
        //    set
        //    {
        //        this.itemsField = value;
        //    }
        //}


        //[XmlElement("ItemsElementName")]
        //[XmlIgnore]
        //public ItemsChoiceType[] ItemsElementName
        //{
        //    get
        //    {
        //        return this.itemsElementNameField;
        //    }
        //    set
        //    {
        //        this.itemsElementNameField = value;
        //    }
        //}


        [XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long w
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


        [XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long h
        {
            get
            {
                return this.hField;
            }
            set
            {
                this.hField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(ST_PathFillMode.norm)]
        public ST_PathFillMode fill
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


        [XmlAttribute]
        [DefaultValue(true)]
        public bool stroke
        {
            get
            {
                return this.strokeField;
            }
            set
            {
                this.strokeField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(true)]
        public bool extrusionOk
        {
            get
            {
                return this.extrusionOkField;
            }
            set
            {
                this.extrusionOkField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType
    {


        arcTo,


        close,


        cubicBezTo,


        lnTo,


        moveTo,


        quadBezTo,
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_Path2DList
    {

        private List<CT_Path2D> pathField;


        //[XmlElement("path", Order = 0)]
        public List<CT_Path2D> path
        {
            get
            {
                return this.pathField;
            }
            set
            {
                this.pathField = value;
            }
        }

        internal static CT_Path2DList Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_Path2DList pathList = new CT_Path2DList();
            pathList.path = new List<CT_Path2D>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "path")
                    pathList.pathField.Add(CT_Path2D.Parse(childNode, namespaceManager));
            }
            return pathList;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write("<a:{0}>", nodeName);
            if (this.pathField != null)
            {
                foreach (CT_Path2D gg in pathField)
                {
                    gg.Write(sw, "path");
                }
            }
            sw.Write("</a:{0}>", nodeName);

        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_PresetGeometry2D
    {

        private CT_GeomGuideList avLstField;

        private ST_ShapeType prstField;

        public CT_GeomGuideList AddNewAvLst()
        {
            this.avLstField = new CT_GeomGuideList();
            return this.avLstField;
        }

        public static CT_PresetGeometry2D Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PresetGeometry2D ctObj = new CT_PresetGeometry2D();
            if (node.Attributes["prst"] != null)
                ctObj.prst = (ST_ShapeType)Enum.Parse(typeof(ST_ShapeType), node.Attributes["prst"].Value);
            if (node.ChildNodes != null)
            {
                foreach (XmlNode childNode in node.ChildNodes)
                {
                    if (childNode.LocalName == "avLst")
                    {
                        ctObj.avLstField = CT_GeomGuideList.Parse(childNode, namespaceManager);
                    }
                }
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "prst", this.prst.ToString());
            sw.Write(">");
            if (this.avLst != null)
            {
                avLst.Write(sw, "avLst");
            }
            sw.Write(string.Format("</a:{0}>", nodeName));
        }



        [XmlElement(Order = 0)]
        public CT_GeomGuideList avLst
        {
            get
            {
                return this.avLstField;
            }
            set
            {
                this.avLstField =value;
            }
        }


        [XmlAttribute]
        public ST_ShapeType prst
        {
            get
            {
                return this.prstField;
            }
            set
            {
                this.prstField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_PresetTextShape
    {
        public static CT_PresetTextShape Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PresetTextShape ctObj = new CT_PresetTextShape();
            if (node.Attributes["prst"] != null)
                ctObj.prst = (ST_TextShapeType)Enum.Parse(typeof(ST_TextShapeType), node.Attributes["prst"].Value);
            ctObj.avLst = new List<CT_GeomGuide>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "avLst")
                    ctObj.avLst.Add(CT_GeomGuide.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "prst", this.prst.ToString());
            sw.Write(">");
            if (this.avLst != null)
            {
                foreach (CT_GeomGuide x in this.avLst)
                {
                    x.Write(sw, "avLst");
                }
            }
            sw.Write(string.Format("</a:{0}>", nodeName));
        }

        private List<CT_GeomGuide> avLstField;

        private ST_TextShapeType prstField;

        [XmlArray(Order = 0)]
        [XmlArrayItem("gd", IsNullable = false)]
        public List<CT_GeomGuide> avLst
        {
            get
            {
                return this.avLstField;
            }
            set
            {
                this.avLstField = value;
            }
        }


        [XmlAttribute]
        public ST_TextShapeType prst
        {
            get
            {
                return this.prstField;
            }
            set
            {
                this.prstField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_CustomGeometry2D
    {

        private CT_GeomGuideList avLstField;

        private CT_GeomGuideList gdLstField;

        private List<object> ahLstField;

        private CT_ConnectionSiteList cxnLstField;

        private CT_GeomRect rectField;

        private CT_Path2DList pathLstField;
        public static CT_CustomGeometry2D Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CustomGeometry2D ctObj = new CT_CustomGeometry2D();
            ctObj.ahLst = new List<Object>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "rect")
                    ctObj.rect = CT_GeomRect.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "avLst")
                    ctObj.avLst = CT_GeomGuideList.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "gdLst")
                    ctObj.gdLst = CT_GeomGuideList.Parse(childNode, namespaceManager);
                //else if (childNode.LocalName == "ahLst")
                //    ctObj.ahLst.Add(Object.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "cxnLst")
                    ctObj.cxnLst = CT_ConnectionSiteList.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pathLst")
                    ctObj.pathLst = CT_Path2DList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            sw.Write(">");
            if (this.rect != null)
                this.rect.Write(sw, "rect");
            if (this.avLst != null)
            {
                this.avLst.Write(sw, "avLst");
            }
            if (this.gdLst != null)
            {
                this.gdLst.Write(sw, "gdLst");
            }
            if (this.cxnLst != null)
            {
                this.cxnLstField.Write(sw, "cxnLst");
            }
            if (this.pathLst != null)
            {
                this.pathLstField.Write(sw, "pathLst");
            }
            sw.Write(string.Format("</a:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        //[XmlArrayItem("gd", IsNullable = false)]
        public CT_GeomGuideList avLst
        {
            get
            {
                return this.avLstField;
            }
            set
            {
                this.avLstField = value;
            }
        }

        [XmlElement(Order = 1)]
        //[XmlArrayItem("gd", IsNullable = false)]
        public CT_GeomGuideList gdLst
        {
            get
            {
                return this.gdLstField;
            }
            set
            {
                this.gdLstField = value;
            }
        }

        [XmlArray(Order = 2)]
        [XmlArrayItem("ahPolar", typeof(CT_PolarAdjustHandle), IsNullable = false)]
        [XmlArrayItem("ahXY", typeof(CT_XYAdjustHandle), IsNullable = false)]
        public List<object> ahLst
        {
            get
            {
                return this.ahLstField;
            }
            set
            {
                this.ahLstField = value;
            }
        }

        [XmlElement(Order = 3)]
        //[XmlArrayItem("cxn", IsNullable = false)]
        public CT_ConnectionSiteList cxnLst
        {
            get
            {
                return this.cxnLstField;
            }
            set
            {
                this.cxnLstField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_GeomRect rect
        {
            get
            {
                return this.rectField;
            }
            set
            {
                this.rectField = value;
            }
        }

        [XmlElement(Order = 5)]
        //[XmlArrayItem("path", IsNullable = false)]
        public CT_Path2DList pathLst
        {
            get
            {
                return this.pathLstField;
            }
            set
            {
                this.pathLstField = value;
            }
        }
    }
}
