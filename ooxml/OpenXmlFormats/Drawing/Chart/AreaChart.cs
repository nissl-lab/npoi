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
    public class CT_Area3DChart
    {

        private CT_Grouping groupingField;

        private CT_Boolean varyColorsField;

        private List<CT_AreaSer> serField;

        private CT_DLbls dLblsField;

        private CT_ChartLines dropLinesField;

        private CT_GapAmount gapDepthField;

        private List<CT_UnsignedInt> axIdField;

        private List<CT_Extension> extLstField;

        public CT_Area3DChart()
        {
        }
        public static CT_Area3DChart Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Area3DChart ctObj = new CT_Area3DChart();
            ctObj.ser = new List<CT_AreaSer>();
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
                    ctObj.ser.Add(CT_AreaSer.Parse(childNode, namespaceManager));
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
                foreach (CT_AreaSer x in this.ser)
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
        public List<CT_AreaSer> ser
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
    }


    [Serializable]

    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart", IsNullable = true)]
    public class CT_AreaSer
    {

        private CT_UnsignedInt idxField;

        private CT_UnsignedInt orderField;

        private CT_SerTx txField;

        private CT_ShapeProperties spPrField;

        private CT_PictureOptions pictureOptionsField;

        private List<CT_DPt> dPtField;

        private CT_DLbls dLblsField;

        private List<CT_Trendline> trendlineField;

        private List<CT_ErrBars> errBarsField;

        private CT_AxDataSource catField;

        private CT_NumDataSource valField;

        private List<CT_Extension> extLstField;
        public static CT_AreaSer Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_AreaSer ctObj = new CT_AreaSer();
            ctObj.dPt = new List<CT_DPt>();
            ctObj.trendline = new List<CT_Trendline>();
            ctObj.errBars = new List<CT_ErrBars>();
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
                else if (childNode.LocalName == "pictureOptions")
                    ctObj.pictureOptions = CT_PictureOptions.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dLbls")
                    ctObj.dLbls = CT_DLbls.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cat")
                    ctObj.cat = CT_AxDataSource.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "val")
                    ctObj.val = CT_NumDataSource.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dPt")
                    ctObj.dPt.Add(CT_DPt.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "trendline")
                    ctObj.trendline.Add(CT_Trendline.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "errBars")
                    ctObj.errBars.Add(CT_ErrBars.Parse(childNode, namespaceManager));
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
            if (this.pictureOptions != null)
                this.pictureOptions.Write(sw, "pictureOptions");
            if (this.dLbls != null)
                this.dLbls.Write(sw, "dLbls");
            if (this.cat != null)
                this.cat.Write(sw, "cat");
            if (this.val != null)
                this.val.Write(sw, "val");
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
            if (this.errBars != null)
            {
                foreach (CT_ErrBars x in this.errBars)
                {
                    x.Write(sw, "errBars");
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

        public CT_AreaSer()
        {
            //this.extLstField = new List<CT_Extension>();
            //this.valField = new CT_NumDataSource();
            //this.catField = new CT_AxDataSource();
            //this.errBarsField = new List<CT_ErrBars>();
            //this.trendlineField = new List<CT_Trendline>();
            //this.dLblsField = new CT_DLbls();
            //this.dPtField = new List<CT_DPt>();
            //this.pictureOptionsField = new CT_PictureOptions();
            //this.txField = new CT_SerTx();
            //this.orderField = new CT_UnsignedInt();
            //this.idxField = new CT_UnsignedInt();
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

        [XmlElement("errBars", Order = 8)]
        public List<CT_ErrBars> errBars
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
    public class CT_AreaChart
    {

        private CT_Grouping groupingField;

        private CT_Boolean varyColorsField;

        private List<CT_AreaSer> serField;

        private CT_DLbls dLblsField;

        private CT_ChartLines dropLinesField;

        private List<CT_UnsignedInt> axIdField;

        private List<CT_Extension> extLstField;

        public CT_AreaChart()
        {

        }
        public static CT_AreaChart Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_AreaChart ctObj = new CT_AreaChart();
            ctObj.ser = new List<CT_AreaSer>();
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
                else if (childNode.LocalName == "ser")
                    ctObj.ser.Add(CT_AreaSer.Parse(childNode, namespaceManager));
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
            if (this.dropLines != null)
                this.dropLines.Write(sw, "dropLines");
            if (this.ser != null)
            {
                foreach (CT_AreaSer x in this.ser)
                {
                    x.Write(sw, "ser");
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
        public List<CT_AreaSer> ser
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
}
