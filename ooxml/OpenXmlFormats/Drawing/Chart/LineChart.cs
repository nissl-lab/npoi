using System;
using System.Collections.Generic;
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
    public class CT_LineSer
    {

        private CT_UnsignedInt idxField;

        private CT_UnsignedInt orderField;

        private CT_SerTx txField;

        private CT_ShapeProperties spPrField;

        private CT_Marker markerField;

        private List<CT_DPt> dPtField;

        private CT_DLbls dLblsField;

        private List<CT_Trendline> trendlineField;

        private CT_ErrBars errBarsField;

        private CT_AxDataSource catField;

        private CT_NumDataSource valField;

        private CT_Boolean smoothField;

        private List<CT_Extension> extLstField;

        public CT_LineSer()
        {
        }

        public static CT_LineSer Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_LineSer ctObj = new CT_LineSer();
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
                else if (childNode.LocalName == "marker")
                    ctObj.marker = CT_Marker.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dLbls")
                    ctObj.dLbls = CT_DLbls.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "errBars")
                    ctObj.errBars = CT_ErrBars.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cat")
                    ctObj.cat = CT_AxDataSource.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "val")
                    ctObj.val = CT_NumDataSource.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "smooth")
                    ctObj.smooth = CT_Boolean.Parse(childNode, namespaceManager);
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
            if (this.marker != null)
                this.marker.Write(sw, "marker");
            if (this.dLbls != null)
                this.dLbls.Write(sw, "dLbls");
            if (this.errBars != null)
                this.errBars.Write(sw, "errBars");
            if (this.cat != null)
                this.cat.Write(sw, "cat");
            if (this.val != null)
                this.val.Write(sw, "val");
            if (this.smooth != null)
                this.smooth.Write(sw, "smooth");
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
        public CT_Marker marker
        {
            get
            {
                return this.markerField;
            }
            set
            {
                this.markerField = value;
            }
        }

        [XmlElement("dPt", Order = 5)]
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

        [XmlElement(Order = 6)]
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

        [XmlElement("trendline", Order = 7)]
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

        [XmlElement(Order = 8)]
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

        [XmlElement(Order = 9)]
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

        [XmlElement(Order = 10)]
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

        [XmlElement(Order = 11)]
        public CT_Boolean smooth
        {
            get
            {
                return this.smoothField;
            }
            set
            {
                this.smoothField = value;
            }
        }

        [XmlElement(Order = 12)]
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

        public CT_UnsignedInt AddNewIdx()
        {
            this.idxField = new CT_UnsignedInt();
            return this.idxField;
        }

        public CT_UnsignedInt AddNewOrder()
        {
            this.orderField = new CT_UnsignedInt();
            return this.orderField;
        }

        public CT_Marker AddNewMarker()
        {
            this.markerField = new CT_Marker();
            return this.markerField;
        }

        public CT_AxDataSource AddNewCat()
        {
            this.catField = new CT_AxDataSource();
            return this.catField;
        }

        public CT_NumDataSource AddNewVal()
        {
            this.valField = new CT_NumDataSource();
            return this.valField;
        }
    }


    [Serializable]

    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart", IsNullable = true)]
    public class CT_Line3DChart
    {

        private CT_Grouping groupingField;

        private CT_Boolean varyColorsField;

        private List<CT_LineSer> serField;

        private CT_DLbls dLblsField;

        private CT_ChartLines dropLinesField;

        private CT_GapAmount gapDepthField;

        private List<CT_UnsignedInt> axIdField;

        private List<CT_Extension> extLstField;

        public CT_Line3DChart()
        {
            //this.extLstField = new List<CT_Extension>();
            //this.axIdField = new List<CT_UnsignedInt>();
            //this.gapDepthField = new CT_GapAmount();
            //this.dropLinesField = new CT_ChartLines();
            //this.dLblsField = new CT_DLbls();
            //this.serField = new List<CT_LineSer>();
            //this.varyColorsField = new CT_Boolean();
            //this.groupingField = new CT_Grouping();
        }

        [XmlElement(Order = 0)]
        public CT_Grouping grouping
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

        [XmlElement(Order = 1)]
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

        [XmlElement("ser", Order = 2)]
        public List<CT_LineSer> ser
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

        [XmlElement(Order = 3)]
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

        [XmlElement(Order = 4)]
        public CT_ChartLines dropLines
        {
            get
            {
                return this.dropLinesField;
            }
            set
            {
                this.dropLinesField = value;
            }
        }

        [XmlElement(Order = 5)]
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

        [XmlElement("axId", Order = 6)]
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

        [XmlElement(Order = 7)]
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
        public static CT_Line3DChart Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Line3DChart ctObj = new CT_Line3DChart();
            ctObj.ser = new List<CT_LineSer>();
            ctObj.axId = new List<CT_UnsignedInt>();
            ctObj.extLst = new List<CT_Extension>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "grouping")
                    ctObj.grouping = CT_Grouping.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "varyColors")
                    ctObj.varyColors = CT_Boolean.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dLbls")
                    ctObj.dLbls = CT_DLbls.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dropLines")
                    ctObj.dropLines = CT_ChartLines.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "gapDepth")
                    ctObj.gapDepth = CT_GapAmount.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "ser")
                    ctObj.ser.Add(CT_LineSer.Parse(childNode, namespaceManager));
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
            if (this.grouping != null)
                this.grouping.Write(sw, "grouping");
            if (this.varyColors != null)
                this.varyColors.Write(sw, "varyColors");
            if (this.dLbls != null)
                this.dLbls.Write(sw, "dLbls");
            if (this.dropLines != null)
                this.dropLines.Write(sw, "dropLines");
            if (this.gapDepth != null)
                this.gapDepth.Write(sw, "gapDepth");
            if (this.ser != null)
            {
                foreach (CT_LineSer x in this.ser)
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




    [Serializable]

    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart", IsNullable = true)]
    public class CT_UpDownBars
    {

        private CT_GapAmount gapWidthField;

        private CT_UpDownBar upBarsField;

        private CT_UpDownBar downBarsField;

        private List<CT_Extension> extLstField;

        public CT_UpDownBars()
        {
        }
        public static CT_UpDownBars Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_UpDownBars ctObj = new CT_UpDownBars();
            ctObj.extLst = new List<CT_Extension>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "gapWidth")
                    ctObj.gapWidth = CT_GapAmount.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "upBars")
                    ctObj.upBars = CT_UpDownBar.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "downBars")
                    ctObj.downBars = CT_UpDownBar.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "extLst")
                    ctObj.extLst.Add(CT_Extension.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<c:{0}", nodeName));
            sw.Write(">");
            if (this.gapWidth != null)
                this.gapWidth.Write(sw, "gapWidth");
            if (this.upBars != null)
                this.upBars.Write(sw, "upBars");
            if (this.downBars != null)
                this.downBars.Write(sw, "downBars");
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

        [XmlElement(Order = 1)]
        public CT_UpDownBar upBars
        {
            get
            {
                return this.upBarsField;
            }
            set
            {
                this.upBarsField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_UpDownBar downBars
        {
            get
            {
                return this.downBarsField;
            }
            set
            {
                this.downBarsField = value;
            }
        }

        [XmlElement(Order = 3)]
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
    public class CT_UpDownBar
    {

        private CT_ShapeProperties spPrField;
        public static CT_UpDownBar Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_UpDownBar ctObj = new CT_UpDownBar();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "spPr")
                    ctObj.spPr = CT_ShapeProperties.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<c:{0}", nodeName));
            sw.Write(">");
            if (this.spPr != null)
                this.spPr.Write(sw, "spPr");
            sw.Write(string.Format("</c:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
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
    }


    [Serializable]

    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart", IsNullable = true)]
    public class CT_StockChart
    {

        private List<CT_LineSer> serField;

        private CT_DLbls dLblsField;

        private CT_ChartLines dropLinesField;

        private CT_ChartLines hiLowLinesField;

        private CT_UpDownBars upDownBarsField;

        private List<CT_UnsignedInt> axIdField;

        private List<CT_Extension> extLstField;

        public CT_StockChart()
        {
        }
        public static CT_StockChart Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_StockChart ctObj = new CT_StockChart();
            ctObj.ser = new List<CT_LineSer>();
            ctObj.axId = new List<CT_UnsignedInt>();
            ctObj.extLst = new List<CT_Extension>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "dLbls")
                    ctObj.dLbls = CT_DLbls.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dropLines")
                    ctObj.dropLines = CT_ChartLines.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "hiLowLines")
                    ctObj.hiLowLines = CT_ChartLines.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "upDownBars")
                    ctObj.upDownBars = CT_UpDownBars.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "ser")
                    ctObj.ser.Add(CT_LineSer.Parse(childNode, namespaceManager));
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
            if (this.dLbls != null)
                this.dLbls.Write(sw, "dLbls");
            if (this.dropLines != null)
                this.dropLines.Write(sw, "dropLines");
            if (this.hiLowLines != null)
                this.hiLowLines.Write(sw, "hiLowLines");
            if (this.upDownBars != null)
                this.upDownBars.Write(sw, "upDownBars");
            if (this.ser != null)
            {
                foreach (CT_LineSer x in this.ser)
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


        [XmlElement("ser", Order = 0)]
        public List<CT_LineSer> ser
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

        [XmlElement(Order = 1)]
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

        [XmlElement(Order = 2)]
        public CT_ChartLines dropLines
        {
            get
            {
                return this.dropLinesField;
            }
            set
            {
                this.dropLinesField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_ChartLines hiLowLines
        {
            get
            {
                return this.hiLowLinesField;
            }
            set
            {
                this.hiLowLinesField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_UpDownBars upDownBars
        {
            get
            {
                return this.upDownBarsField;
            }
            set
            {
                this.upDownBarsField = value;
            }
        }

        [XmlElement("axId", Order = 5)]
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

        [XmlElement(Order = 6)]
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
    public class CT_LineChart
    {

        private CT_Grouping groupingField;

        private CT_Boolean varyColorsField;

        private List<CT_LineSer> serField;

        private CT_DLbls dLblsField;

        private CT_ChartLines dropLinesField;

        private CT_ChartLines hiLowLinesField;

        private CT_UpDownBars upDownBarsField;

        private CT_Boolean markerField;

        private CT_Boolean smoothField;

        private List<CT_UnsignedInt> axIdField;

        private List<CT_Extension> extLstField;

        public CT_LineChart()
        {
        }
        public static CT_LineChart Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_LineChart ctObj = new CT_LineChart();
            ctObj.ser = new List<CT_LineSer>();
            ctObj.axId = new List<CT_UnsignedInt>();
            ctObj.extLst = new List<CT_Extension>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "grouping")
                    ctObj.grouping = CT_Grouping.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "varyColors")
                    ctObj.varyColors = CT_Boolean.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dLbls")
                    ctObj.dLbls = CT_DLbls.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dropLines")
                    ctObj.dropLines = CT_ChartLines.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "hiLowLines")
                    ctObj.hiLowLines = CT_ChartLines.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "upDownBars")
                    ctObj.upDownBars = CT_UpDownBars.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "marker")
                    ctObj.marker = CT_Boolean.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "smooth")
                    ctObj.smooth = CT_Boolean.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "ser")
                    ctObj.ser.Add(CT_LineSer.Parse(childNode, namespaceManager));
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
            if (this.grouping != null)
                this.grouping.Write(sw, "grouping");
            if (this.varyColors != null)
                this.varyColors.Write(sw, "varyColors");
            if (this.ser != null)
            {
                foreach (CT_LineSer x in this.ser)
                {
                    x.Write(sw, "ser");
                }
            }
            if (this.dLbls != null)
                this.dLbls.Write(sw, "dLbls");
            if (this.dropLines != null)
                this.dropLines.Write(sw, "dropLines");
            if (this.hiLowLines != null)
                this.hiLowLines.Write(sw, "hiLowLines");
            if (this.upDownBars != null)
                this.upDownBars.Write(sw, "upDownBars");
            if (this.marker != null)
                this.marker.Write(sw, "marker");
            if (this.smooth != null)
                this.smooth.Write(sw, "smooth");

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
        public CT_Grouping grouping
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

        [XmlElement(Order = 1)]
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

        [XmlElement("ser", Order = 2)]
        public List<CT_LineSer> ser
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

        [XmlElement(Order = 3)]
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

        [XmlElement(Order = 4)]
        public CT_ChartLines dropLines
        {
            get
            {
                return this.dropLinesField;
            }
            set
            {
                this.dropLinesField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_ChartLines hiLowLines
        {
            get
            {
                return this.hiLowLinesField;
            }
            set
            {
                this.hiLowLinesField = value;
            }
        }

        [XmlElement(Order = 6)]
        public CT_UpDownBars upDownBars
        {
            get
            {
                return this.upDownBarsField;
            }
            set
            {
                this.upDownBarsField = value;
            }
        }

        [XmlElement(Order = 7)]
        public CT_Boolean marker
        {
            get
            {
                return this.markerField;
            }
            set
            {
                this.markerField = value;
            }
        }

        [XmlElement(Order = 8)]
        public CT_Boolean smooth
        {
            get
            {
                return this.smoothField;
            }
            set
            {
                this.smoothField = value;
            }
        }

        [XmlElement("axId", Order = 9)]
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

        [XmlElement(Order = 10)]
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
        public CT_Grouping AddNewGrouping()
        {
            this.groupingField = new CT_Grouping();
            return this.groupingField;
        }
        public CT_LineSer AddNewSer()
        {
            CT_LineSer newSer = new  CT_LineSer();
            if (this.serField == null)
            {
                this.serField = new List<CT_LineSer>();
            }
            this.serField.Add(newSer);
            return newSer;
        }

        public CT_Boolean AddNewVaryColors()
        {
            this.varyColorsField = new CT_Boolean();
            return this.varyColorsField;
        }

        public CT_UnsignedInt AddNewAxId()
        {
            CT_UnsignedInt si = new CT_UnsignedInt();
            if (this.axIdField == null)
                this.axIdField = new List<CT_UnsignedInt>();
            axIdField.Add(si);
            return si;
        }
    }


}
