using System;
using System.ComponentModel;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Dml.Picture
{
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/picture")]
    [XmlRoot("pic", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/picture", IsNullable = false)]
    // draw-pic:pic
    public class CT_Picture
    {
        private CT_PictureNonVisual nvPicPrField;        //  draw-pic 1..1 

        private CT_BlipFillProperties blipFillField;   //  draw-pic: 1..1 

        private CT_ShapeProperties spPrField;             //  draw-pic: 1..1 

        [XmlElement(Order = 0)]
        public CT_PictureNonVisual nvPicPr
        {
            get { return this.nvPicPrField; }
            set { this.nvPicPrField = value; }
        }

        [XmlElement(Order = 1)]
        public CT_BlipFillProperties blipFill
        {
            get { return this.blipFillField; }
            set { this.blipFillField = value; }
        }

        [XmlElement(Order = 2)]
        public CT_ShapeProperties spPr
        {
            get { return this.spPrField; }
            set { this.spPrField = value; }
        }

        public CT_PictureNonVisual AddNewNvPicPr()
        {
            nvPicPrField = new CT_PictureNonVisual();
            return this.nvPicPrField;
        }

        public CT_BlipFillProperties AddNewBlipFill()
        {
            blipFillField = new CT_BlipFillProperties();
            return this.blipFillField;
        }

        public CT_ShapeProperties AddNewSpPr()
        {
            spPrField = new CT_ShapeProperties();
            return this.spPrField;
        }
        public void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0} xmlns:pic=\"{1}\">", nodeName, "http://schemas.openxmlformats.org/drawingml/2006/picture"));
            if (this.nvPicPr != null)
            {
                this.nvPicPr.Write(sw, "pic:nvPicPr");
            }
            if (this.blipFill != null)
            {
                this.blipFill.Write(sw, "pic:blipFill");
            }
            if (this.spPr != null)
            {
                this.spPr.Write(sw, "pic:spPr");
            }
            sw.Write(string.Format("</{0}>",nodeName));
        }

        public static CT_Picture Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if(node == null) return null;
            CT_Picture ctObj = new CT_Picture();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if(childNode.LocalName == "nvPicPr")
                    ctObj.nvPicPrField = CT_PictureNonVisual.Parse(childNode, namespaceManager);
                else if(childNode.LocalName == "blipFill")
                    ctObj.blipFillField = CT_BlipFillProperties.Parse(childNode, namespaceManager);
                else if(childNode.LocalName == "spPrField")
                    ctObj.spPrField = CT_ShapeProperties.Parse(childNode, namespaceManager);
            }

            return ctObj;
        }
    }

    // see same class in different name space in SpeedsheetDrawing.cs
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/picture")]
    public class CT_PictureNonVisual
    {

        private CT_NonVisualDrawingProps cNvPrField = new CT_NonVisualDrawingProps(); // 1..1
        private CT_NonVisualPictureProperties cNvPicPrField = new CT_NonVisualPictureProperties(); // 1..1

        public CT_NonVisualDrawingProps AddNewCNvPr()
        {
            this.cNvPrField = new CT_NonVisualDrawingProps();
            return this.cNvPrField;
        }
        public CT_NonVisualPictureProperties AddNewCNvPicPr()
        {
            this.cNvPicPrField = new CT_NonVisualPictureProperties();
            return this.cNvPicPrField;
        }

        [XmlElement(Order = 0)]
        public CT_NonVisualDrawingProps cNvPr
        {
            get { return this.cNvPrField; }
            set { this.cNvPrField = value; }
        }


        [XmlElement(Order = 1)]
        public CT_NonVisualPictureProperties cNvPicPr
        {
            get { return this.cNvPicPrField; }
            set { this.cNvPicPrField = value; }
        }

        internal void Write(StreamWriter sw, string p)
        {
            sw.Write(string.Format("<{0}>",p));
            if (this.cNvPr!=null)
            {
                this.cNvPr.Write(sw, "cNvPr");
            }
            if (this.cNvPicPr != null)
            {
                this.cNvPicPr.Write(sw, "cNvPicPr");
            }
            sw.Write(string.Format("</{0}>", p));
        }

        public static CT_PictureNonVisual Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if(node == null) return null;
            CT_PictureNonVisual ctObj = new CT_PictureNonVisual();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if(childNode.LocalName == "cNvPr")
                    ctObj.cNvPrField = CT_NonVisualDrawingProps.Parse(childNode, namespaceManager);
                else if(childNode.LocalName == "cNvPicPr")
                    ctObj.cNvPicPrField = CT_NonVisualPictureProperties.Parse(childNode, namespaceManager);

            }

            return ctObj;
        }
    }

}
