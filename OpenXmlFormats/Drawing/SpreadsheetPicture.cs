using NPOI.OpenXml4Net.Util;
using System;
using System.ComponentModel;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Dml.Spreadsheet
{

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_Picture // empty interface: EG_ObjectChoices
    {
        private CT_PictureNonVisual nvPicPrField = new CT_PictureNonVisual();        //  draw-ssdraw: 1..1 
        private CT_BlipFillProperties blipFillField = new CT_BlipFillProperties();   //  draw-ssdraw: 1..1 
        private CT_ShapeProperties spPrField = new CT_ShapeProperties();             //  draw-ssdraw: 1..1 
        private CT_ShapeStyle styleField = null; // 0..1

        private string macroField = null;
        private bool fPublishedField = false;
        public static CT_Picture Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Picture ctObj = new CT_Picture();
            ctObj.macro = XmlHelper.ReadString(node.Attributes["macro"]);
            ctObj.fPublished = XmlHelper.ReadBool(node.Attributes["fPublished"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "nvPicPr")
                    ctObj.nvPicPr = CT_PictureNonVisual.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "blipFill")
                    ctObj.blipFill = CT_BlipFillProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "spPr")
                    ctObj.spPr = CT_ShapeProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "style")
                    ctObj.style = CT_ShapeStyle.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "macro", this.macro);
            if (this.fPublished)
                XmlHelper.WriteAttribute(sw, "fPublished", this.fPublished);
            sw.Write(">");
            if (this.nvPicPr != null)
                this.nvPicPr.Write(sw, "nvPicPr");
            if (this.blipFill != null)
                this.blipFill.Write(sw, "blipFill");
            if (this.spPr != null)
                this.spPr.Write(sw, "spPr");
            if (this.style != null)
                this.style.Write(sw, "style");
            sw.Write(string.Format("</xdr:{0}>", nodeName));
        }

        [XmlElement]
        public CT_PictureNonVisual nvPicPr
        {
            get { return this.nvPicPrField; }
            set { this.nvPicPrField = value; }
        }

        [XmlElement]
        public CT_BlipFillProperties blipFill
        {
            get { return this.blipFillField; }
            set { this.blipFillField = value; }
        }

        [XmlElement]
        public CT_ShapeProperties spPr
        {
            get { return this.spPrField; }
            set { this.spPrField = value; }
        }

        [XmlElement]
        public CT_ShapeStyle style
        {
            get
            {
                return this.styleField;
            }
            set
            {
                this.styleField = value;
            }
        }
        private bool styleSpecifiedField = false;
        [XmlIgnore]
        public bool styleSpecified
        {
            get { return styleSpecifiedField; }
            set { styleSpecifiedField = value; }
        }

        [XmlAttribute]
        public string macro
        {
            get { return macroField; }
            set { macroField = value; }
        }
        private bool macroSpecifiedField = false;
        [XmlIgnore]
        public bool macroSpecified
        {
            get { return macroSpecifiedField; }
            set { macroSpecifiedField = value; }
        }

        [XmlAttribute]
        public bool fPublished
        {
            get { return fPublishedField; }
            set { fPublishedField = value; }
        }
        private bool fPublishedSpecifiedField = false;
        [XmlIgnore]
        public bool fPublishedSpecified
        {
            get { return fPublishedSpecifiedField; }
            set { fPublishedSpecifiedField = value; }
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

        public void Set(CT_Picture pict)
        {
            this.nvPicPr = pict.nvPicPr;
            this.spPr = pict.spPr;
            this.macro = pict.macro;
            this.macroSpecified = this.macroSpecified;
            this.style = pict.style;
            this.styleSpecified = pict.styleSpecified;
            this.fPublished = pict.fPublished;
            this.fPublishedSpecified = pict.fPublishedSpecified;
            this.blipFill = pict.blipFill;


        }

    }

    // see same class in different name space in Picture.cs
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_PictureNonVisual
    {

        private CT_NonVisualDrawingProps cNvPrField = new CT_NonVisualDrawingProps(); // 1..1
        private CT_NonVisualPictureProperties cNvPicPrField = new CT_NonVisualPictureProperties(); // 1..1

        public static CT_PictureNonVisual Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PictureNonVisual ctObj = new CT_PictureNonVisual();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "cNvPr")
                    ctObj.cNvPr = CT_NonVisualDrawingProps.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cNvPicPr")
                    ctObj.cNvPicPr = CT_NonVisualPictureProperties.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}", nodeName));
            sw.Write(">");
            if (this.cNvPr != null)
                this.cNvPr.Write(sw, "cNvPr");
            if (this.cNvPicPr != null)
                this.cNvPicPr.Write(sw, "cNvPicPr");
            sw.Write(string.Format("</xdr:{0}>", nodeName));
        }
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

        [XmlElement]
        public CT_NonVisualDrawingProps cNvPr
        {
            get { return this.cNvPrField; }
            set { this.cNvPrField = value; }
        }


        [XmlElement]
        public CT_NonVisualPictureProperties cNvPicPr
        {
            get { return this.cNvPicPrField; }
            set { this.cNvPicPrField = value; }
        }

    }

}
