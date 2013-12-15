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
    public class CT_ScatterSer
    {

        private CT_UnsignedInt idxField;

        private CT_UnsignedInt orderField;

        private CT_SerTx txField;

        private CT_ShapeProperties spPrField;

        private CT_Marker markerField;

        private List<CT_DPt> dPtField;

        private CT_DLbls dLblsField;

        private List<CT_Trendline> trendlineField;

        private List<CT_ErrBars> errBarsField;

        private CT_AxDataSource xValField;

        private CT_NumDataSource yValField;

        private CT_Boolean smoothField;

        private List<CT_Extension> extLstField;

        public CT_ScatterSer()
        {

        }
        public static CT_ScatterSer Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ScatterSer ctObj = new CT_ScatterSer();
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
                else if (childNode.LocalName == "marker")
                    ctObj.marker = CT_Marker.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dLbls")
                    ctObj.dLbls = CT_DLbls.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "xVal")
                    ctObj.xVal = CT_AxDataSource.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "yVal")
                    ctObj.yVal = CT_NumDataSource.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "smooth")
                    ctObj.smooth = CT_Boolean.Parse(childNode, namespaceManager);
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
            if (this.marker != null)
                this.marker.Write(sw, "marker");
            if (this.dLbls != null)
                this.dLbls.Write(sw, "dLbls");
            if (this.xVal != null)
                this.xVal.Write(sw, "xVal");
            if (this.yVal != null)
                this.yVal.Write(sw, "yVal");
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


        public CT_UnsignedInt AddNewIdx()
        {
            this.idxField = new CT_UnsignedInt();
            return idxField;
        }
        public CT_UnsignedInt AddNewOrder()
        {
            this.orderField = new CT_UnsignedInt();
            return orderField;
        }
        public CT_AxDataSource AddNewXVal()
        {
            this.xValField = new CT_AxDataSource();
            return this.xValField;
        }
        public CT_NumDataSource AddNewYVal()
        {
            this.yValField = new CT_NumDataSource();
            return this.yValField;
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
        public CT_AxDataSource xVal
        {
            get
            {
                return this.xValField;
            }
            set
            {
                this.xValField = value;
            }
        }

        [XmlElement(Order = 10)]
        public CT_NumDataSource yVal
        {
            get
            {
                return this.yValField;
            }
            set
            {
                this.yValField = value;
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
    }


    [Serializable]

    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart", IsNullable = true)]
    public class CT_ScatterStyle
    {
        public static CT_ScatterStyle Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ScatterStyle ctObj = new CT_ScatterStyle();
            if (node.Attributes["val"] != null)
                ctObj.val = (ST_ScatterStyle)Enum.Parse(typeof(ST_ScatterStyle), node.Attributes["val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<c:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</c:{0}>", nodeName));
        }

        private ST_ScatterStyle valField;

        public CT_ScatterStyle()
        {
            this.valField = ST_ScatterStyle.marker;
        }

        [XmlAttribute]
        [DefaultValue(ST_ScatterStyle.marker)]
        public ST_ScatterStyle val
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
    public enum ST_ScatterStyle
    {

        /// <remarks/>
        none,

        /// <remarks/>
        line,

        /// <remarks/>
        lineMarker,

        /// <remarks/>
        marker,

        /// <remarks/>
        smooth,

        /// <remarks/>
        smoothMarker,
    }


    [Serializable]

    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart", IsNullable = true)]
    public class CT_ScatterChart
    {

        private CT_ScatterStyle scatterStyleField;

        private CT_Boolean varyColorsField;

        private List<CT_ScatterSer> serField;

        private CT_DLbls dLblsField;

        private List<CT_UnsignedInt> axIdField;

        private List<CT_Extension> extLstField;

        public CT_ScatterChart()
        {
            //this.extLstField = new List<CT_Extension>();
            //this.dLblsField = new CT_DLbls();
            //this.varyColorsField = new CT_Boolean();
        }
        public static CT_ScatterChart Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ScatterChart ctObj = new CT_ScatterChart();
            ctObj.ser = new List<CT_ScatterSer>();
            ctObj.axId = new List<CT_UnsignedInt>();
            ctObj.extLst = new List<CT_Extension>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "scatterStyle")
                    ctObj.scatterStyle = CT_ScatterStyle.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "varyColors")
                    ctObj.varyColors = CT_Boolean.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dLbls")
                    ctObj.dLbls = CT_DLbls.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "ser")
                    ctObj.ser.Add(CT_ScatterSer.Parse(childNode, namespaceManager));
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
            if (this.scatterStyle != null)
                this.scatterStyle.Write(sw, "scatterStyle");
            if (this.varyColors != null)
                this.varyColors.Write(sw, "varyColors");
            if (this.dLbls != null)
                this.dLbls.Write(sw, "dLbls");
            if (this.ser != null)
            {
                foreach (CT_ScatterSer x in this.ser)
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

        public CT_ScatterStyle AddNewScatterStyle()
        {
            this.scatterStyleField = new CT_ScatterStyle();
            return scatterStyleField;
        }
        public CT_UnsignedInt AddNewAxId()
        {
            if (this.axIdField == null)
                this.axIdField = new List<CT_UnsignedInt>();
            CT_UnsignedInt axIdItem = new CT_UnsignedInt();
            this.axIdField.Add(axIdItem);
            return axIdItem;
        }
        public CT_ScatterSer AddNewSer()
        {
            if (this.serField == null)
                this.serField = new List<CT_ScatterSer>();
            CT_ScatterSer ser = new CT_ScatterSer();
            this.serField.Add(ser);
            return ser;
        }
        [XmlElement(Order = 0)]
        public CT_ScatterStyle scatterStyle
        {
            get
            {
                return this.scatterStyleField;
            }
            set
            {
                this.scatterStyleField = value;
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
        public List<CT_ScatterSer> ser
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

        [XmlElement("axId", Order = 4)]
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

        [XmlElement(Order = 5)]
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
