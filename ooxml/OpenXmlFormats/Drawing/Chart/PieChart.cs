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
    public class CT_Pie3DChart
    {

        private CT_Boolean varyColorsField;

        private List<CT_PieSer> serField;

        private CT_DLbls dLblsField;

        private List<CT_Extension> extLstField;

        public CT_Pie3DChart()
        {
        }
        public static CT_Pie3DChart Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Pie3DChart ctObj = new CT_Pie3DChart();
            ctObj.ser = new List<CT_PieSer>();
            ctObj.extLst = new List<CT_Extension>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "varyColors")
                    ctObj.varyColors = CT_Boolean.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dLbls")
                    ctObj.dLbls = CT_DLbls.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "ser")
                    ctObj.ser.Add(CT_PieSer.Parse(childNode, namespaceManager));
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
            if (this.ser != null)
            {
                foreach (CT_PieSer x in this.ser)
                {
                    x.Write(sw, "ser");
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
        public List<CT_PieSer> ser
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
    public class CT_PieChart
    {

        private CT_Boolean varyColorsField;

        private List<CT_PieSer> serField;

        private CT_DLbls dLblsField;

        private CT_FirstSliceAng firstSliceAngField;

        private List<CT_Extension> extLstField;

        public CT_PieChart()
        {
        }
        public static CT_PieChart Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PieChart ctObj = new CT_PieChart();
            ctObj.ser = new List<CT_PieSer>();
            ctObj.extLst = new List<CT_Extension>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "varyColors")
                    ctObj.varyColors = CT_Boolean.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dLbls")
                    ctObj.dLbls = CT_DLbls.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "firstSliceAng")
                    ctObj.firstSliceAng = CT_FirstSliceAng.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "ser")
                    ctObj.ser.Add(CT_PieSer.Parse(childNode, namespaceManager));
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
            if (this.firstSliceAng != null)
                this.firstSliceAng.Write(sw, "firstSliceAng");
            if (this.ser != null)
            {
                foreach (CT_PieSer x in this.ser)
                {
                    x.Write(sw, "ser");
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
        public List<CT_PieSer> ser
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
        public CT_FirstSliceAng firstSliceAng
        {
            get
            {
                return this.firstSliceAngField;
            }
            set
            {
                this.firstSliceAngField = value;
            }
        }

        [XmlElement(Order = 4)]
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
    public class CT_PieSer
    {

        private CT_UnsignedInt idxField;

        private CT_UnsignedInt orderField;

        private CT_SerTx txField;

        private CT_ShapeProperties spPrField;

        private CT_UnsignedInt explosionField;

        private List<CT_DPt> dPtField;

        private CT_DLbls dLblsField;

        private CT_AxDataSource catField;

        private CT_NumDataSource valField;

        private List<CT_Extension> extLstField;

        public CT_PieSer()
        {

        }
        public static CT_PieSer Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PieSer ctObj = new CT_PieSer();
            ctObj.dPt = new List<CT_DPt>();
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
                else if (childNode.LocalName == "explosion")
                    ctObj.explosion = CT_UnsignedInt.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dLbls")
                    ctObj.dLbls = CT_DLbls.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cat")
                    ctObj.cat = CT_AxDataSource.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "val")
                    ctObj.val = CT_NumDataSource.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dPt")
                    ctObj.dPt.Add(CT_DPt.Parse(childNode, namespaceManager));
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
            if (this.explosion != null)
                this.explosion.Write(sw, "explosion");
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
        public CT_UnsignedInt explosion
        {
            get
            {
                return this.explosionField;
            }
            set
            {
                this.explosionField = value;
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

        [XmlElement(Order = 7)]
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

        [XmlElement(Order = 8)]
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
    public class CT_HoleSize
    {

        private byte valField;

        public CT_HoleSize()
        {
            this.valField = ((byte)(10));
        }
        public static CT_HoleSize Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_HoleSize ctObj = new CT_HoleSize();
            if (node.Attributes["val"] != null)
                ctObj.val = XmlHelper.ReadByte(node.Attributes["val"]);
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
        [DefaultValue(typeof(byte), "10")]
        public byte val
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
    public class CT_FirstSliceAng
    {

        private ushort valField;

        public CT_FirstSliceAng()
        {
            this.valField = ((ushort)(0));
        }
        public static CT_FirstSliceAng Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_FirstSliceAng ctObj = new CT_FirstSliceAng();
            if (node.Attributes["val"] != null)
                ctObj.val = XmlHelper.ReadUShort(node.Attributes["val"]);
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
        [DefaultValue(typeof(ushort), "0")]
        public ushort val
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
    public class CT_DoughnutChart
    {

        private CT_Boolean varyColorsField;

        private List<CT_PieSer> serField;

        private CT_DLbls dLblsField;

        private CT_FirstSliceAng firstSliceAngField;

        private CT_HoleSize holeSizeField;

        private List<CT_Extension> extLstField;

        public CT_DoughnutChart()
        {
        }
        public static CT_DoughnutChart Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_DoughnutChart ctObj = new CT_DoughnutChart();
            ctObj.ser = new List<CT_PieSer>();
            ctObj.extLst = new List<CT_Extension>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "varyColors")
                    ctObj.varyColors = CT_Boolean.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dLbls")
                    ctObj.dLbls = CT_DLbls.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "firstSliceAng")
                    ctObj.firstSliceAng = CT_FirstSliceAng.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "holeSize")
                    ctObj.holeSize = CT_HoleSize.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "ser")
                    ctObj.ser.Add(CT_PieSer.Parse(childNode, namespaceManager));
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
            if (this.firstSliceAng != null)
                this.firstSliceAng.Write(sw, "firstSliceAng");
            if (this.holeSize != null)
                this.holeSize.Write(sw, "holeSize");
            if (this.ser != null)
            {
                foreach (CT_PieSer x in this.ser)
                {
                    x.Write(sw, "ser");
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
        public List<CT_PieSer> ser
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
        public CT_FirstSliceAng firstSliceAng
        {
            get
            {
                return this.firstSliceAngField;
            }
            set
            {
                this.firstSliceAngField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_HoleSize holeSize
        {
            get
            {
                return this.holeSizeField;
            }
            set
            {
                this.holeSizeField = value;
            }
        }

        [XmlArray(Order = 5)]
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
    public class CT_OfPieType
    {

        private ST_OfPieType valField;

        public CT_OfPieType()
        {
            this.valField = ST_OfPieType.pie;
        }
        public static CT_OfPieType Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_OfPieType ctObj = new CT_OfPieType();
            if (node.Attributes["val"] != null)
                ctObj.val = (ST_OfPieType)Enum.Parse(typeof(ST_OfPieType), node.Attributes["val"].Value);
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
        [DefaultValue(ST_OfPieType.pie)]
        public ST_OfPieType val
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
    public enum ST_OfPieType
    {

        /// <remarks/>
        pie,

        /// <remarks/>
        bar,
    }


    [Serializable]

    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart", IsNullable = true)]
    public class CT_OfPieChart
    {

        private CT_OfPieType ofPieTypeField;

        private CT_Boolean varyColorsField;

        private List<CT_PieSer> serField;

        private CT_DLbls dLblsField;

        private CT_GapAmount gapWidthField;

        private CT_SplitType splitTypeField;

        private CT_Double splitPosField;

        private List<CT_UnsignedInt> custSplitField;

        private CT_SecondPieSize secondPieSizeField;

        private List<CT_ChartLines> serLinesField;

        private List<CT_Extension> extLstField;

        public CT_OfPieChart()
        {
        }

        public static CT_OfPieChart Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_OfPieChart ctObj = new CT_OfPieChart();
            ctObj.ser = new List<CT_PieSer>();
            ctObj.custSplit = new List<CT_UnsignedInt>();
            ctObj.serLines = new List<CT_ChartLines>();
            ctObj.extLst = new List<CT_Extension>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "ofPieType")
                    ctObj.ofPieType = CT_OfPieType.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "varyColors")
                    ctObj.varyColors = CT_Boolean.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dLbls")
                    ctObj.dLbls = CT_DLbls.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "gapWidth")
                    ctObj.gapWidth = CT_GapAmount.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "splitType")
                    ctObj.splitType = CT_SplitType.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "splitPos")
                    ctObj.splitPos = CT_Double.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "secondPieSize")
                    ctObj.secondPieSize = CT_SecondPieSize.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "ser")
                    ctObj.ser.Add(CT_PieSer.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "custSplit")
                    ctObj.custSplit.Add(CT_UnsignedInt.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "serLines")
                    ctObj.serLines.Add(CT_ChartLines.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "extLst")
                    ctObj.extLst.Add(CT_Extension.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<c:{0}", nodeName));
            sw.Write(">");
            if (this.ofPieType != null)
                this.ofPieType.Write(sw, "ofPieType");
            if (this.varyColors != null)
                this.varyColors.Write(sw, "varyColors");
            if (this.dLbls != null)
                this.dLbls.Write(sw, "dLbls");
            if (this.gapWidth != null)
                this.gapWidth.Write(sw, "gapWidth");
            if (this.splitType != null)
                this.splitType.Write(sw, "splitType");
            if (this.splitPos != null)
                this.splitPos.Write(sw, "splitPos");
            if (this.secondPieSize != null)
                this.secondPieSize.Write(sw, "secondPieSize");
            if (this.ser != null)
            {
                foreach (CT_PieSer x in this.ser)
                {
                    x.Write(sw, "ser");
                }
            }
            if (this.custSplit != null)
            {
                foreach (CT_UnsignedInt x in this.custSplit)
                {
                    x.Write(sw, "custSplit");
                }
            }
            if (this.serLines != null)
            {
                foreach (CT_ChartLines x in this.serLines)
                {
                    x.Write(sw, "serLines");
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
        public CT_OfPieType ofPieType
        {
            get
            {
                return this.ofPieTypeField;
            }
            set
            {
                this.ofPieTypeField = value;
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
        public List<CT_PieSer> ser
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

        [XmlElement(Order = 5)]
        public CT_SplitType splitType
        {
            get
            {
                return this.splitTypeField;
            }
            set
            {
                this.splitTypeField = value;
            }
        }

        [XmlElement(Order = 6)]
        public CT_Double splitPos
        {
            get
            {
                return this.splitPosField;
            }
            set
            {
                this.splitPosField = value;
            }
        }

        [XmlElement(Order = 7)]
        public List<CT_UnsignedInt> custSplit
        {
            get
            {
                return this.custSplitField;
            }
            set
            {
                this.custSplitField = value;
            }
        }

        [XmlElement(Order = 8)]
        public CT_SecondPieSize secondPieSize
        {
            get
            {
                return this.secondPieSizeField;
            }
            set
            {
                this.secondPieSizeField = value;
            }
        }

        [XmlElement("serLines", Order = 9)]
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

        [XmlArray(Order = 10)]
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

}
