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

            CT_Picture ctShape = new CT_Picture();
            if (node.Attributes["macro"] != null)
                ctShape.macroField = node.Attributes["macro"].Value;
            if (node.Attributes["fPublished"] != null && node.Attributes["fPublished"].Value == "1")
                ctShape.fPublishedField = true;

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "nvPicPr")
                {
                    ctShape.nvPicPr = CT_PictureNonVisual.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "spPr")
                {
                    ctShape.spPr = CT_ShapeProperties.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "blipFill")
                {
                    ctShape.blipFill = CT_BlipFillProperties.Parse(childNode, namespaceManager);
                }
            }
            return ctShape;
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

        internal void Write(StreamWriter sw)
        {
            sw.Write("<xdr:pic");
            if (this.macroField != null)
            {
                sw.Write(string.Format(" macro=\"{0}\"", this.macroField));
            }
            if (this.fPublished)
            {
                sw.Write(" fPublished=\"1\"");
            }
            sw.Write(">");
            if (this.nvPicPr != null)
            {
                this.nvPicPr.Write(sw);
            }
            if (this.blipFill != null)
            {
                this.blipFill.Write(sw);
            }
            //if (this.spPr != null)
            //{
            //    this.spPr.Write(sw);
            //}
            sw.Write("</xdr:pic>");
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
            CT_PictureNonVisual ctNvPr = new CT_PictureNonVisual();
            XmlNode cnvprNode = node.SelectSingleNode("xdr:cNvPr", namespaceManager);
            if (cnvprNode != null)
            {
                CT_NonVisualDrawingProps ctProps = ctNvPr.AddNewCNvPr();

                ctProps.id = XmlHelper.ReadUInt(cnvprNode.Attributes["id"]);
                ctProps.name = cnvprNode.Attributes["name"].Value;
                ctProps.descr = cnvprNode.Attributes["descr"].Value;
                ctProps.hidden = XmlHelper.ReadBool(cnvprNode.Attributes["hidden"]);
                //TODO::hlinkClick, hlinkCover
            }
            XmlNode cNvPicPrNode = node.SelectSingleNode("xdr:cNvPicPr", namespaceManager);
            if (cNvPicPrNode != null)
            {
                CT_NonVisualPictureProperties ctNvPicPr = ctNvPr.AddNewCNvPicPr();
                ctNvPicPr.preferRelativeResize = XmlHelper.ReadBool(cNvPicPrNode.Attributes["preferRelativeResize"]);
                XmlNode picLocksNode = cNvPicPrNode.SelectSingleNode("a:picLocks", namespaceManager);
                if (picLocksNode != null)
                {
                    ctNvPicPr.picLocks = new CT_PictureLocking();
                    ctNvPicPr.picLocks.noChangeAspect = XmlHelper.ReadBool(picLocksNode.Attributes["noChangeAspect"]);
                    ctNvPicPr.picLocks.noAdjustHandles = XmlHelper.ReadBool(picLocksNode.Attributes["noAdjustHandles"]);
                    ctNvPicPr.picLocks.noChangeArrowheads = XmlHelper.ReadBool(picLocksNode.Attributes["noChangeArrowheads"]);
                    ctNvPicPr.picLocks.noChangeShapeType = XmlHelper.ReadBool(picLocksNode.Attributes["noChangeShapeType"]);
                    ctNvPicPr.picLocks.noCrop = XmlHelper.ReadBool(picLocksNode.Attributes["noCrop"]);
                    ctNvPicPr.picLocks.noEditPoints = XmlHelper.ReadBool(picLocksNode.Attributes["noEditPoints"]);
                    ctNvPicPr.picLocks.noGrp = XmlHelper.ReadBool(picLocksNode.Attributes["noGrp"]);
                    ctNvPicPr.picLocks.noMove = XmlHelper.ReadBool(picLocksNode.Attributes["noMove"]);
                    ctNvPicPr.picLocks.noResize = XmlHelper.ReadBool(picLocksNode.Attributes["noResize"]);
                    ctNvPicPr.picLocks.noRot = XmlHelper.ReadBool(picLocksNode.Attributes["noRot"]);
                    ctNvPicPr.picLocks.noSelect = XmlHelper.ReadBool(picLocksNode.Attributes["noSelect"]);
                }
            }
            return ctNvPr;
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

        internal void Write(StreamWriter sw)
        {
            sw.Write("<xdr:nvPicPr>");
            if (this.cNvPr != null)
            {
                sw.Write(string.Format("<xdr:cNvPr id=\"{0}\" name=\"{1}\"", this.cNvPr.id, this.cNvPr.name));
                XmlHelper.WriteAttribute(sw, "descr", this.cNvPr.descr);
                XmlHelper.WriteAttribute(sw, "hidden", this.cNvPr.hidden);
                sw.Write("/>");
            }
            if (this.cNvPicPr != null)
            {
                sw.Write("<xdr:cNvPicPr>");
                if (this.cNvPicPr.picLocks != null)
                {
                    sw.Write("<a:picLocks");
                    XmlHelper.WriteAttribute(sw, "noChangeAspect", this.cNvPicPr.picLocks.noChangeAspect);
                    XmlHelper.WriteAttribute(sw, "noAdjustHandles", this.cNvPicPr.picLocks.noAdjustHandles);
                    XmlHelper.WriteAttribute(sw, "noChangeArrowheads", this.cNvPicPr.picLocks.noChangeArrowheads);
                    XmlHelper.WriteAttribute(sw, "noChangeShapeType", this.cNvPicPr.picLocks.noChangeShapeType);
                    XmlHelper.WriteAttribute(sw, "noCrop", this.cNvPicPr.picLocks.noCrop);
                    XmlHelper.WriteAttribute(sw, "noEditPoints", this.cNvPicPr.picLocks.noEditPoints);
                    XmlHelper.WriteAttribute(sw, "noGrp", this.cNvPicPr.picLocks.noGrp);
                    XmlHelper.WriteAttribute(sw, "noMove", this.cNvPicPr.picLocks.noMove);
                    XmlHelper.WriteAttribute(sw, "noResize", this.cNvPicPr.picLocks.noResize);
                    XmlHelper.WriteAttribute(sw, "noRot", this.cNvPicPr.picLocks.noRot);
                    XmlHelper.WriteAttribute(sw, "noSelect", this.cNvPicPr.picLocks.noSelect);
                    sw.Write("/>");
                }
                sw.Write("</xdr:cNvPicPr>");
            }
            sw.Write("</xdr:nvPicPr>");
        }
    }

}
