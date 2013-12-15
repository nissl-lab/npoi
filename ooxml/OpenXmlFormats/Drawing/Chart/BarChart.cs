using NPOI.OpenXml4Net.Util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Dml.Chart
{


    [Serializable]

    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart", IsNullable = true)]
    public class CT_BarChart
    {

        private CT_BarDir barDirField;

        private CT_BarGrouping groupingField;

        private CT_Boolean varyColorsField;

        private List<CT_BarSer> serField;

        private CT_DLbls dLblsField;

        private CT_GapAmount gapWidthField;

        private CT_Overlap overlapField;

        private List<CT_ChartLines> serLinesField;

        private List<CT_UnsignedInt> axIdField;

        private List<CT_Extension> extLstField;

        public CT_BarChart()
        {
        }
        public static CT_BarChart Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_BarChart ctObj = new CT_BarChart();
            ctObj.ser = new List<CT_BarSer>();
            ctObj.serLines = new List<CT_ChartLines>();
            ctObj.axId = new List<CT_UnsignedInt>();
            ctObj.extLst = new List<CT_Extension>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "barDir")
                    ctObj.barDir = CT_BarDir.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "grouping")
                    ctObj.grouping = CT_BarGrouping.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "varyColors")
                    ctObj.varyColors = CT_Boolean.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dLbls")
                    ctObj.dLbls = CT_DLbls.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "gapWidth")
                    ctObj.gapWidth = CT_GapAmount.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "overlap")
                    ctObj.overlap = CT_Overlap.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "ser")
                    ctObj.ser.Add(CT_BarSer.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "serLines")
                    ctObj.serLines.Add(CT_ChartLines.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "axId")
                    ctObj.axId.Add(CT_UnsignedInt.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "extLst")
                    ctObj.extLst.Add(CT_Extension.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<c:{0}", nodeName));
            sw.Write(">");
            if (this.barDir != null)
                this.barDir.Write(sw, "barDir");
            if (this.grouping != null)
                this.grouping.Write(sw, "grouping");
            if (this.varyColors != null)
                this.varyColors.Write(sw, "varyColors");
            if (this.gapWidth != null)
                this.gapWidth.Write(sw, "gapWidth");
            if (this.overlap != null)
                this.overlap.Write(sw, "overlap");
            if (this.ser != null)
            {
                foreach (CT_BarSer x in this.ser)
                {
                    x.Write(sw, "ser");
                }
            }
            if (this.serLines != null)
            {
                foreach (CT_ChartLines x in this.serLines)
                {
                    x.Write(sw, "serLines");
                }
            }
            if (this.dLbls != null)
                this.dLbls.Write(sw, "dLbls");
            if (this.axId != null)
            {
                foreach (CT_UnsignedInt x in this.axId)
                {
                    x.Write(sw, "axId");
                }
            }
            if (this.extLst != null)
            {
                foreach (CT_Extension x in this.extLst)
                {
                    x.Write(sw, "extLst");
                }
            }
            sw.Write(string.Format("</c:{0}>", nodeName));
        }


        [XmlElement(Order = 0)]
        public CT_BarDir barDir
        {
            get
            {
                return this.barDirField;
            }
            set
            {
                this.barDirField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_BarGrouping grouping
        {
            get
            {
                return this.groupingField;
            }
            set
            {
                this.groupingField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_Boolean varyColors
        {
            get
            {
                return this.varyColorsField;
            }
            set
            {
                this.varyColorsField = value;
            }
        }

        [XmlElement("ser", Order = 3)]
        public List<CT_BarSer> ser
        {
            get
            {
                return this.serField;
            }
            set
            {
                this.serField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_DLbls dLbls
        {
            get
            {
                return this.dLblsField;
            }
            set
            {
                this.dLblsField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_GapAmount gapWidth
        {
            get
            {
                return this.gapWidthField;
            }
            set
            {
                this.gapWidthField = value;
            }
        }

        [XmlElement(Order = 6)]
        public CT_Overlap overlap
        {
            get
            {
                return this.overlapField;
            }
            set
            {
                this.overlapField = value;
            }
        }

        [XmlElement("serLines", Order = 7)]
        public List<CT_ChartLines> serLines
        {
            get
            {
                return this.serLinesField;
            }
            set
            {
                this.serLinesField = value;
            }
        }

        [XmlElement("axId", Order = 8)]
        public List<CT_UnsignedInt> axId
        {
            get
            {
                return this.axIdField;
            }
            set
            {
                this.axIdField = value;
            }
        }

        [XmlArray(Order = 9)]
        [XmlArrayItem("ext", IsNullable = false)]
        public List<CT_Extension> extLst
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
    }


    [Serializable]

    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart", IsNullable = true)]
    public class CT_BarDir
    {

        private ST_BarDir valField;

        public CT_BarDir()
        {
            this.valField = ST_BarDir.col;
        }
        public static CT_BarDir Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_BarDir ctObj = new CT_BarDir();
            if (node.Attributes["val"] != null)
                ctObj.val = (ST_BarDir)Enum.Parse(typeof(ST_BarDir), node.Attributes["val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<c:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</c:{0}>", nodeName));
        }

        [XmlAttribute]
        [DefaultValue(ST_BarDir.col)]
        public ST_BarDir val
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart")]
    public enum ST_BarDir
    {

        /// <remarks/>
        bar,

        /// <remarks/>
        col,
    }


    [Serializable]

    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart", IsNullable = true)]
    public class CT_BarGrouping
    {

        private ST_BarGrouping valField;

        public CT_BarGrouping()
        {
            this.valField = ST_BarGrouping.clustered;
        }
        public static CT_BarGrouping Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_BarGrouping ctObj = new CT_BarGrouping();
            if (node.Attributes["val"] != null)
                ctObj.val = (ST_BarGrouping)Enum.Parse(typeof(ST_BarGrouping), node.Attributes["val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<c:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</c:{0}>", nodeName));
        }

        [XmlAttribute]
        [DefaultValue(ST_BarGrouping.clustered)]
        public ST_BarGrouping val
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart")]
    public enum ST_BarGrouping
    {

        /// <remarks/>
        percentStacked,

        /// <remarks/>
        clustered,

        /// <remarks/>
        standard,

        /// <remarks/>
        stacked,
    }


    [Serializable]

    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart", IsNullable = true)]
    public class CT_BarSer
    {

        private CT_UnsignedInt idxField;

        private CT_UnsignedInt orderField;

        private CT_SerTx txField;

        private CT_ShapeProperties spPrField;

        private CT_Boolean invertIfNegativeField;

        private CT_PictureOptions pictureOptionsField;

        private List<CT_DPt> dPtField;

        private CT_DLbls dLblsField;

        private List<CT_Trendline> trendlineField;

        private CT_ErrBars errBarsField;

        private CT_AxDataSource catField;

        private CT_NumDataSource valField;

        private CT_Shape shapeField;

        private List<CT_Extension> extLstField;

        public CT_BarSer()
        {
        }
        public static CT_BarSer Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_BarSer ctObj = new CT_BarSer();
            ctObj.dPt = new List<CT_DPt>();
            ctObj.trendline = new List<CT_Trendline>();
            ctObj.extLst = new List<CT_Extension>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "idx")
                    ctObj.idx = CT_UnsignedInt.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "order")
                    ctObj.order = CT_UnsignedInt.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tx")
                    ctObj.tx = CT_SerTx.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "spPr")
                    ctObj.spPr = CT_ShapeProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "invertIfNegative")
                    ctObj.invertIfNegative = CT_Boolean.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pictureOptions")
                    ctObj.pictureOptions = CT_PictureOptions.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dLbls")
                    ctObj.dLbls = CT_DLbls.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "errBars")
                    ctObj.errBars = CT_ErrBars.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cat")
                    ctObj.cat = CT_AxDataSource.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "val")
                    ctObj.val = CT_NumDataSource.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "shape")
                    ctObj.shape = CT_Shape.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dPt")
                    ctObj.dPt.Add(CT_DPt.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "trendline")
                    ctObj.trendline.Add(CT_Trendline.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "extLst")
                    ctObj.extLst.Add(CT_Extension.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<c:{0}", nodeName));
            sw.Write(">");
            if (this.idx != null)
                this.idx.Write(sw, "idx");
            if (this.order != null)
                this.order.Write(sw, "order");
            if (this.tx != null)
                this.tx.Write(sw, "tx");
            if (this.spPr != null)
                this.spPr.Write(sw, "spPr");
            if (this.invertIfNegative != null)
                this.invertIfNegative.Write(sw, "invertIfNegative");
            if (this.pictureOptions != null)
                this.pictureOptions.Write(sw, "pictureOptions");
            if (this.dLbls != null)
                this.dLbls.Write(sw, "dLbls");
            if (this.errBars != null)
                this.errBars.Write(sw, "errBars");
            if (this.cat != null)
                this.cat.Write(sw, "cat");
            if (this.val != null)
                this.val.Write(sw, "val");
            if (this.shape != null)
                this.shape.Write(sw, "shape");
            if (this.dPt != null)
            {
                foreach (CT_DPt x in this.dPt)
                {
                    x.Write(sw, "dPt");
                }
            }
            if (this.trendline != null)
            {
                foreach (CT_Trendline x in this.trendline)
                {
                    x.Write(sw, "trendline");
                }
            }
            if (this.extLst != null)
            {
                foreach (CT_Extension x in this.extLst)
                {
                    x.Write(sw, "extLst");
                }
            }
            sw.Write(string.Format("</c:{0}>", nodeName));
        }


        [XmlElement(Order = 0)]
        public CT_UnsignedInt idx
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

        [XmlElement(Order = 1)]
        public CT_UnsignedInt order
        {
            get
            {
                return this.orderField;
            }
            set
            {
                this.orderField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_SerTx tx
        {
            get
            {
                return this.txField;
            }
            set
            {
                this.txField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_ShapeProperties spPr
        {
            get
            {
                return this.spPrField;
            }
            set
            {
                this.spPrField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_Boolean invertIfNegative
        {
            get
            {
                return this.invertIfNegativeField;
            }
            set
            {
                this.invertIfNegativeField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_PictureOptions pictureOptions
        {
            get
            {
                return this.pictureOptionsField;
            }
            set
            {
                this.pictureOptionsField = value;
            }
        }

        [XmlElement("dPt", Order = 6)]
        public List<CT_DPt> dPt
        {
            get
            {
                return this.dPtField;
            }
            set
            {
                this.dPtField = value;
            }
        }

        [XmlElement(Order = 7)]
        public CT_DLbls dLbls
        {
            get
            {
                return this.dLblsField;
            }
            set
            {
                this.dLblsField = value;
            }
        }

        [XmlElement("trendline", Order = 8)]
        public List<CT_Trendline> trendline
        {
            get
            {
                return this.trendlineField;
            }
            set
            {
                this.trendlineField = value;
            }
        }

        [XmlElement(Order = 9)]
        public CT_ErrBars errBars
        {
            get
            {
                return this.errBarsField;
            }
            set
            {
                this.errBarsField = value;
            }
        }

        [XmlElement(Order = 10)]
        public CT_AxDataSource cat
        {
            get
            {
                return this.catField;
            }
            set
            {
                this.catField = value;
            }
        }

        [XmlElement(Order = 11)]
        public CT_NumDataSource val
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

        [XmlElement(Order = 12)]
        public CT_Shape shape
        {
            get
            {
                return this.shapeField;
            }
            set
            {
                this.shapeField = value;
            }
        }

        [XmlElement(Order = 13)]
        public List<CT_Extension> extLst
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
    }


    [Serializable]

    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart", IsNullable = true)]
    public class CT_Bar3DChart
    {

        private CT_BarDir barDirField;

        private CT_BarGrouping groupingField;

        private CT_Boolean varyColorsField;

        private List<CT_BarSer> serField;

        private CT_DLbls dLblsField;

        private CT_GapAmount gapWidthField;

        private CT_GapAmount gapDepthField;

        private CT_Shape shapeField;

        private List<CT_UnsignedInt> axIdField;

        private List<CT_Extension> extLstField;

        public CT_Bar3DChart()
        {
            this.extLstField = new List<CT_Extension>();
            this.axIdField = new List<CT_UnsignedInt>();
            this.shapeField = new CT_Shape();
            this.gapDepthField = new CT_GapAmount();
            this.gapWidthField = new CT_GapAmount();
            this.dLblsField = new CT_DLbls();
            this.serField = new List<CT_BarSer>();
            this.varyColorsField = new CT_Boolean();
            this.groupingField = new CT_BarGrouping();
            this.barDirField = new CT_BarDir();
        }

        [XmlElement(Order = 0)]
        public CT_BarDir barDir
        {
            get
            {
                return this.barDirField;
            }
            set
            {
                this.barDirField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_BarGrouping grouping
        {
            get
            {
                return this.groupingField;
            }
            set
            {
                this.groupingField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_Boolean varyColors
        {
            get
            {
                return this.varyColorsField;
            }
            set
            {
                this.varyColorsField = value;
            }
        }

        [XmlElement("ser", Order = 3)]
        public List<CT_BarSer> ser
        {
            get
            {
                return this.serField;
            }
            set
            {
                this.serField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_DLbls dLbls
        {
            get
            {
                return this.dLblsField;
            }
            set
            {
                this.dLblsField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_GapAmount gapWidth
        {
            get
            {
                return this.gapWidthField;
            }
            set
            {
                this.gapWidthField = value;
            }
        }

        [XmlElement(Order = 6)]
        public CT_GapAmount gapDepth
        {
            get
            {
                return this.gapDepthField;
            }
            set
            {
                this.gapDepthField = value;
            }
        }

        [XmlElement(Order = 7)]
        public CT_Shape shape
        {
            get
            {
                return this.shapeField;
            }
            set
            {
                this.shapeField = value;
            }
        }

        [XmlElement("axId", Order = 8)]
        public List<CT_UnsignedInt> axId
        {
            get
            {
                return this.axIdField;
            }
            set
            {
                this.axIdField = value;
            }
        }

        [XmlArray(Order = 9)]
        [XmlArrayItem("ext", IsNullable = false)]
        public List<CT_Extension> extLst
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
        public static CT_Bar3DChart Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Bar3DChart ctObj = new CT_Bar3DChart();
            ctObj.ser = new List<CT_BarSer>();
            ctObj.axId = new List<CT_UnsignedInt>();
            ctObj.extLst = new List<CT_Extension>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "barDir")
                    ctObj.barDir = CT_BarDir.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "grouping")
                    ctObj.grouping = CT_BarGrouping.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "varyColors")
                    ctObj.varyColors = CT_Boolean.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dLbls")
                    ctObj.dLbls = CT_DLbls.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "gapWidth")
                    ctObj.gapWidth = CT_GapAmount.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "gapDepth")
                    ctObj.gapDepth = CT_GapAmount.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "shape")
                    ctObj.shape = CT_Shape.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "ser")
                    ctObj.ser.Add(CT_BarSer.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "axId")
                    ctObj.axId.Add(CT_UnsignedInt.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "extLst")
                    ctObj.extLst.Add(CT_Extension.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<c:{0}", nodeName));
            sw.Write(">");
            if (this.barDir != null)
                this.barDir.Write(sw, "barDir");
            if (this.grouping != null)
                this.grouping.Write(sw, "grouping");
            if (this.varyColors != null)
                this.varyColors.Write(sw, "varyColors");
            if (this.dLbls != null)
                this.dLbls.Write(sw, "dLbls");
            if (this.gapWidth != null)
                this.gapWidth.Write(sw, "gapWidth");
            if (this.gapDepth != null)
                this.gapDepth.Write(sw, "gapDepth");
            if (this.shape != null)
                this.shape.Write(sw, "shape");
            if (this.ser != null)
            {
                foreach (CT_BarSer x in this.ser)
                {
                    x.Write(sw, "ser");
                }
            }
            if (this.axId != null)
            {
                foreach (CT_UnsignedInt x in this.axId)
                {
                    x.Write(sw, "axId");
                }
            }
            if (this.extLst != null)
            {
                foreach (CT_Extension x in this.extLst)
                {
                    x.Write(sw, "extLst");
                }
            }
            sw.Write(string.Format("</c:{0}>", nodeName));
        }

    }

}
