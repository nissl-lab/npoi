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
    public class CT_Surface
    {

        private CT_UnsignedInt thicknessField;

        private CT_ShapeProperties spPrField;

        private CT_PictureOptions pictureOptionsField;

        private List<CT_Extension> extLstField;

        public CT_Surface()
        {
        }
        public static CT_Surface Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Surface ctObj = new CT_Surface();
            ctObj.extLst = new List<CT_Extension>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "thickness")
                    ctObj.thickness = CT_UnsignedInt.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "spPr")
                    ctObj.spPr = CT_ShapeProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pictureOptions")
                    ctObj.pictureOptions = CT_PictureOptions.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "extLst")
                    ctObj.extLst.Add(CT_Extension.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<c:{0}", nodeName));
            sw.Write(">");
            if (this.thickness != null)
                this.thickness.Write(sw, "thickness");
            if (this.spPr != null)
                this.spPr.Write(sw, "spPr");
            if (this.pictureOptions != null)
                this.pictureOptions.Write(sw, "pictureOptions");
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
        public CT_UnsignedInt thickness
        {
            get
            {
                return this.thicknessField;
            }
            set
            {
                this.thicknessField = value;
            }
        }

        [XmlElement(Order = 1)]
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

        [XmlElement(Order = 2)]
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
    public class CT_Surface3DChart
    {

        private CT_Boolean wireframeField;

        private List<CT_SurfaceSer> serField;

        private List<CT_BandFmt> bandFmtsField;

        private List<CT_UnsignedInt> axIdField;

        private List<CT_Extension> extLstField;

        public CT_Surface3DChart()
        {

        }

        public static CT_Surface3DChart Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Surface3DChart ctObj = new CT_Surface3DChart();
            ctObj.ser = new List<CT_SurfaceSer>();
            ctObj.bandFmts = new List<CT_BandFmt>();
            ctObj.axId = new List<CT_UnsignedInt>();
            ctObj.extLst = new List<CT_Extension>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "wireframe")
                    ctObj.wireframe = CT_Boolean.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "ser")
                    ctObj.ser.Add(CT_SurfaceSer.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "bandFmts")
                    ctObj.bandFmts.Add(CT_BandFmt.Parse(childNode, namespaceManager));
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
            if (this.wireframe != null)
                this.wireframe.Write(sw, "wireframe");
            if (this.ser != null)
            {
                foreach (CT_SurfaceSer x in this.ser)
                {
                    x.Write(sw, "ser");
                }
            }
            if (this.bandFmts != null)
            {
                foreach (CT_BandFmt x in this.bandFmts)
                {
                    x.Write(sw, "bandFmts");
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
        public CT_Boolean wireframe
        {
            get
            {
                return this.wireframeField;
            }
            set
            {
                this.wireframeField = value;
            }
        }

        [XmlElement("ser", Order = 1)]
        public List<CT_SurfaceSer> ser
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
        public List<CT_BandFmt> bandFmts
        {
            get
            {
                return this.bandFmtsField;
            }
            set
            {
                this.bandFmtsField = value;
            }
        }

        [XmlElement("axId", Order = 3)]
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
    public class CT_SurfaceSer
    {

        private CT_UnsignedInt idxField;

        private CT_UnsignedInt orderField;

        private CT_SerTx txField;

        private CT_ShapeProperties spPrField;

        private CT_AxDataSource catField;

        private CT_NumDataSource valField;

        private List<CT_Extension> extLstField;

        public CT_SurfaceSer()
        {

        }

        public static CT_SurfaceSer Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SurfaceSer ctObj = new CT_SurfaceSer();
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
                else if (childNode.LocalName == "cat")
                    ctObj.cat = CT_AxDataSource.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "val")
                    ctObj.val = CT_NumDataSource.Parse(childNode, namespaceManager);
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
            if (this.cat != null)
                this.cat.Write(sw, "cat");
            if (this.val != null)
                this.val.Write(sw, "val");
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

        [XmlElement(Order = 5)]
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
