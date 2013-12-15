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
    public class CT_SizeRepresents
    {

        private ST_SizeRepresents valField;

        public CT_SizeRepresents()
        {
            this.valField = ST_SizeRepresents.area;
        }
        public static CT_SizeRepresents Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SizeRepresents ctObj = new CT_SizeRepresents();
            if (node.Attributes["val"] != null)
                ctObj.val = (ST_SizeRepresents)Enum.Parse(typeof(ST_SizeRepresents), node.Attributes["val"].Value);
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
        [DefaultValue(ST_SizeRepresents.area)]
        public ST_SizeRepresents val
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
    public enum ST_SizeRepresents
    {

        /// <remarks/>
        area,

        /// <remarks/>
        w,
    }
    [Serializable]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart", IsNullable = true)]
    public class CT_BubbleScale
    {

        private uint valField;

        public CT_BubbleScale()
        {
            this.valField = ((uint)(100));
        }
        public static CT_BubbleScale Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_BubbleScale ctObj = new CT_BubbleScale();
            if (node.Attributes["val"] != null)
                ctObj.val = XmlHelper.ReadUInt(node.Attributes["val"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<c:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "val", this.val);
            sw.Write(">");
            sw.Write(string.Format("</c:{0}>", nodeName));
        }

        [XmlAttribute]
        [DefaultValue(typeof(uint), "100")]
        public uint val
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
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart", IsNullable = true)]
    public class CT_BubbleSer
    {

        private CT_UnsignedInt idxField;

        private CT_UnsignedInt orderField;

        private CT_SerTx txField;

        private CT_ShapeProperties spPrField;

        private CT_Boolean invertIfNegativeField;

        private List<CT_DPt> dPtField;

        private CT_DLbls dLblsField;

        private List<CT_Trendline> trendlineField;

        private List<CT_ErrBars> errBarsField;

        private CT_AxDataSource xValField;

        private CT_NumDataSource yValField;

        private CT_NumDataSource bubbleSizeField;

        private CT_Boolean bubble3DField;

        private List<CT_Extension> extLstField;

        public CT_BubbleSer()
        {
        }
        public static CT_BubbleSer Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_BubbleSer ctObj = new CT_BubbleSer();
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
                else if (childNode.LocalName == "invertIfNegative")
                    ctObj.invertIfNegative = CT_Boolean.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dLbls")
                    ctObj.dLbls = CT_DLbls.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "xVal")
                    ctObj.xVal = CT_AxDataSource.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "yVal")
                    ctObj.yVal = CT_NumDataSource.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bubbleSize")
                    ctObj.bubbleSize = CT_NumDataSource.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bubble3D")
                    ctObj.bubble3D = CT_Boolean.Parse(childNode, namespaceManager);
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
            if (this.invertIfNegative != null)
                this.invertIfNegative.Write(sw, "invertIfNegative");
            if (this.dLbls != null)
                this.dLbls.Write(sw, "dLbls");
            if (this.xVal != null)
                this.xVal.Write(sw, "xVal");
            if (this.yVal != null)
                this.yVal.Write(sw, "yVal");
            if (this.bubbleSize != null)
                this.bubbleSize.Write(sw, "bubbleSize");
            if (this.bubble3D != null)
                this.bubble3D.Write(sw, "bubble3D");
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
        public CT_NumDataSource bubbleSize
        {
            get
            {
                return this.bubbleSizeField;
            }
            set
            {
                this.bubbleSizeField = value;
            }
        }

        [XmlElement(Order = 12)]
        public CT_Boolean bubble3D
        {
            get
            {
                return this.bubble3DField;
            }
            set
            {
                this.bubble3DField = value;
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
    public class CT_BubbleChart
    {

        private CT_Boolean varyColorsField;

        private List<CT_BubbleSer> serField;

        private CT_DLbls dLblsField;

        private CT_Boolean bubble3DField;

        private CT_BubbleScale bubbleScaleField;

        private CT_Boolean showNegBubblesField;

        private CT_SizeRepresents sizeRepresentsField;

        private List<CT_UnsignedInt> axIdField;

        private List<CT_Extension> extLstField;

        public CT_BubbleChart()
        {

        }
        public static CT_BubbleChart Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_BubbleChart ctObj = new CT_BubbleChart();
            ctObj.ser = new List<CT_BubbleSer>();
            ctObj.axId = new List<CT_UnsignedInt>();
            ctObj.extLst = new List<CT_Extension>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "varyColors")
                    ctObj.varyColors = CT_Boolean.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dLbls")
                    ctObj.dLbls = CT_DLbls.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bubble3D")
                    ctObj.bubble3D = CT_Boolean.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bubbleScale")
                    ctObj.bubbleScale = CT_BubbleScale.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "showNegBubbles")
                    ctObj.showNegBubbles = CT_Boolean.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sizeRepresents")
                    ctObj.sizeRepresents = CT_SizeRepresents.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "ser")
                    ctObj.ser.Add(CT_BubbleSer.Parse(childNode, namespaceManager));
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
            if (this.varyColors != null)
                this.varyColors.Write(sw, "varyColors");
            if (this.dLbls != null)
                this.dLbls.Write(sw, "dLbls");
            if (this.bubble3D != null)
                this.bubble3D.Write(sw, "bubble3D");
            if (this.bubbleScale != null)
                this.bubbleScale.Write(sw, "bubbleScale");
            if (this.showNegBubbles != null)
                this.showNegBubbles.Write(sw, "showNegBubbles");
            if (this.sizeRepresents != null)
                this.sizeRepresents.Write(sw, "sizeRepresents");
            if (this.ser != null)
            {
                foreach (CT_BubbleSer x in this.ser)
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

        [XmlElement("ser", Order = 1)]
        public List<CT_BubbleSer> ser
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

        [XmlElement(Order = 2)]
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

        [XmlElement(Order = 3)]
        public CT_Boolean bubble3D
        {
            get
            {
                return this.bubble3DField;
            }
            set
            {
                this.bubble3DField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_BubbleScale bubbleScale
        {
            get
            {
                return this.bubbleScaleField;
            }
            set
            {
                this.bubbleScaleField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_Boolean showNegBubbles
        {
            get
            {
                return this.showNegBubblesField;
            }
            set
            {
                this.showNegBubblesField = value;
            }
        }

        [XmlElement(Order = 6)]
        public CT_SizeRepresents sizeRepresents
        {
            get
            {
                return this.sizeRepresentsField;
            }
            set
            {
                this.sizeRepresentsField = value;
            }
        }

        [XmlElement("axId", Order = 7)]
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

        [XmlElement(Order = 8)]
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
