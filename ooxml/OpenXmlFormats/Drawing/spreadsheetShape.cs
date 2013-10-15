using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Dml.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_Shape // empty interface: EG_ObjectChoices
    {
        private CT_ShapeNonVisual nvSpPrField;
        private CT_ShapeProperties spPrField;
        private CT_ShapeStyle styleField;
        private CT_TextBody txBodyField;

        private string macroField;
        private string textlinkField;
        private bool fLocksTextField;
        private bool fPublishedField;

        public void Set(CT_Shape obj)
        {
            this.macroField = obj.macro;
            this.textlinkField = obj.textlink;
            this.fLocksTextField = obj.fLocksText;
            this.fPublishedField = obj.fPublished;

            this.nvSpPrField = obj.nvSpPr;
            this.spPrField = obj.spPr;
            this.styleField = obj.style;
            this.txBodyField = obj.txBody;
        }

        public static CT_Shape Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Shape ctShape = new CT_Shape();
            if (node.Attributes["macro"] != null)
                ctShape.macroField = node.Attributes["macro"].Value;
            if (node.Attributes["textlink"] != null)
                ctShape.textlinkField = node.Attributes["textlink"].Value;
            if (node.Attributes["fLocksText"] != null && node.Attributes["fLocksText"].Value == "1")
                ctShape.fLocksTextField = true;
            if (node.Attributes["fPublished"] != null && node.Attributes["fPublished"].Value == "1")
                ctShape.fPublishedField = true;

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "nvPicPr")
                {
                    ctShape.nvSpPr = CT_ShapeNonVisual.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "spPr")
                {
                    ctShape.spPr = CT_ShapeProperties.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "style")
                {
                    throw new NotImplementedException();
                    //ctShape.blipFill = CT_BlipFillProperties.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "txBody")
                {
                    throw new NotImplementedException();
                }
            }
            return ctShape;
        }

        public CT_ShapeNonVisual AddNewNvSpPr()
        {
            this.nvSpPrField = new CT_ShapeNonVisual();
            return this.nvSpPrField;
        }

        public CT_ShapeProperties AddNewSpPr()
        {
            this.spPrField = new CT_ShapeProperties();
            return this.spPrField;
        }
        public CT_ShapeStyle AddNewStyle()
        {
            this.styleField = new CT_ShapeStyle();
            return this.styleField;
        }
        public CT_TextBody AddNewTxBody()
        {
            this.txBodyField = new CT_TextBody();
            return this.txBodyField;
        }

        public CT_ShapeNonVisual nvSpPr
        {
            get
            {
                return this.nvSpPrField;
            }
            set
            {
                this.nvSpPrField = value;
            }
        }
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
        public CT_TextBody txBody
        {
            get
            {
                return this.txBodyField;
            }
            set
            {
                this.txBodyField = value;
            }
        }
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

        [XmlAttribute]
        public string macro
        {
            get { return macroField; }
            set { macroField = value; }
        }

        [XmlAttribute]
        public string textlink
        {
            get { return textlinkField; }
            set { textlinkField = value; }
        }
        [XmlAttribute]
        public bool fLocksText
        {
            get { return fLocksTextField; }
            set { fLocksTextField = value; }
        }

        [XmlAttribute]
        public bool fPublished
        {
            get { return fPublishedField; }
            set { fPublishedField = value; }
        }

        internal void Write(StreamWriter sw)
        {
            sw.Write("<xdr:sp");
            if (this.macroField != null)
            {
                sw.Write(string.Format(" macro=\"{0}\"", this.macroField));
            }
            if (this.textlinkField != null)
            {
                sw.Write(string.Format(" textlink=\"{0}\"", this.textlinkField));
            }
            if (this.fLocksTextField)
            {
                sw.Write(" fLocksText=\"1\"");
            }
            if (this.fPublishedField)
            {
                sw.Write(" fPublished=\"1\"");
            }
            sw.Write(">");
            if (this.nvSpPr != null)
            {
                this.nvSpPr.Write(sw);
            }
            if (this.spPr != null)
            {
                this.spPr.Write(sw);
            }
            sw.Write("</xdr:sp>");
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_ShapeNonVisual
    {
        private CT_NonVisualDrawingProps cNvPrField;
        private CT_NonVisualDrawingShapeProps cNvSpPrField;

        public CT_NonVisualDrawingProps AddNewCNvPr()
        {
            this.cNvPrField = new CT_NonVisualDrawingProps();
            return this.cNvPrField;
        }
        public CT_NonVisualDrawingShapeProps AddNewCNvSpPr()
        {
            this.cNvSpPrField = new CT_NonVisualDrawingShapeProps();
            return this.cNvSpPrField;
        }
        public CT_NonVisualDrawingProps cNvPr
        {
            get
            {
                return this.cNvPrField;
            }
            set
            {
                this.cNvPrField = value;
            }
        }

        public CT_NonVisualDrawingShapeProps cNvSpPr
        {
            get
            {
                return this.cNvSpPrField;
            }
            set
            {
                this.cNvSpPrField = value;
            }
        }

        internal static CT_ShapeNonVisual Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_ShapeNonVisual ctNvSpPr = new CT_ShapeNonVisual();
            XmlNode cnvprNode = node.SelectSingleNode("xdr:cNvPr", namespaceManager);
            if (cnvprNode != null)
            {
                CT_NonVisualDrawingProps ctProps = ctNvSpPr.AddNewCNvPr();

                ctProps.id = uint.Parse(cnvprNode.Attributes["id"].Value);
                ctProps.name = cnvprNode.Attributes["name"].Value;
                ctProps.descr = cnvprNode.Attributes["descr"].Value;
                if (cnvprNode.Attributes["hidden"] != null && cnvprNode.Attributes["hidden"].Value == "1")
                    ctProps.hidden = true;
                //TODO::hlinkClick, hlinkCover
            }
            XmlNode cNvSpPrNode = node.SelectSingleNode("xdr:cNvSpPr", namespaceManager);
            if (cNvSpPrNode != null)
            {
                CT_NonVisualDrawingShapeProps ctcNvSpPr = ctNvSpPr.AddNewCNvSpPr();
                XmlNode spLocksNode = cNvSpPrNode.SelectSingleNode("a:spLocks", namespaceManager);
                if (spLocksNode != null)
                {
                    ctcNvSpPr.spLocks = new CT_ShapeLocking();

                    if (spLocksNode.Attributes["noChangeAspect"] != null && spLocksNode.Attributes["noChangeAspect"].Value == "1")
                        ctcNvSpPr.spLocks.noChangeAspect = true;

                    if (spLocksNode.Attributes["noAdjustHandles"] != null && spLocksNode.Attributes["noAdjustHandles"].Value == "1")
                        ctcNvSpPr.spLocks.noAdjustHandles = true;

                    if (spLocksNode.Attributes["noChangeArrowheads"] != null && spLocksNode.Attributes["noChangeArrowheads"].Value == "1")
                        ctcNvSpPr.spLocks.noChangeArrowheads = true;

                    if (spLocksNode.Attributes["noChangeShapeType"] != null && spLocksNode.Attributes["noChangeShapeType"].Value == "1")
                        ctcNvSpPr.spLocks.noChangeShapeType = true;

                    if (spLocksNode.Attributes["noEditPoints"] != null && spLocksNode.Attributes["noEditPoints"].Value == "1")
                        ctcNvSpPr.spLocks.noEditPoints = true;

                    if (spLocksNode.Attributes["noGrp"] != null && spLocksNode.Attributes["noGrp"].Value == "1")
                        ctcNvSpPr.spLocks.noGrp = true;

                    if (spLocksNode.Attributes["noMove"] != null && spLocksNode.Attributes["noMove"].Value == "1")
                        ctcNvSpPr.spLocks.noMove = true;

                    if (spLocksNode.Attributes["noResize"] != null && spLocksNode.Attributes["noResize"].Value == "1")
                        ctcNvSpPr.spLocks.noResize = true;

                    if (spLocksNode.Attributes["noRot"] != null && spLocksNode.Attributes["noRot"].Value == "1")
                        ctcNvSpPr.spLocks.noRot = true;

                    if (spLocksNode.Attributes["noSelect"] != null && spLocksNode.Attributes["noSelect"].Value == "1")
                        ctcNvSpPr.spLocks.noSelect = true;
                }
            }
            return ctNvSpPr;
        }

        internal void Write(StreamWriter sw)
        {
            sw.Write("<xdr:nvSpPr>");
            if (this.cNvPr != null)
            {
                sw.Write(string.Format("<xdr:cNvPr id=\"{0}\" name=\"{1}\" descr=\"{2}\"", this.cNvPr.id, this.cNvPr.name, this.cNvPr.descr));
                if (this.cNvPr.hidden)
                {
                    sw.Write(" hidden=\"1\"");
                }
                sw.Write("/>");
            }
            if (this.cNvSpPr != null)
            {
                sw.Write("<xdr:cNvSpPr>");
                if (this.cNvSpPr.spLocks != null)
                {
                    sw.Write("<a:spLocks");
                    if (this.cNvSpPr.spLocks.noChangeAspect)
                    {
                        sw.Write(" noChangeAspect=\"1\"");
                    }
                    if (this.cNvSpPr.spLocks.noAdjustHandles)
                    {
                        sw.Write(" noAdjustHandles=\"1\"");
                    }
                    if (this.cNvSpPr.spLocks.noChangeArrowheads)
                    {
                        sw.Write(" noChangeArrowheads=\"1\"");
                    }
                    if (this.cNvSpPr.spLocks.noChangeShapeType)
                    {
                        sw.Write(" noChangeShapeType=\"1\"");
                    }
                    if (this.cNvSpPr.spLocks.noEditPoints)
                    {
                        sw.Write(" noEditPoints=\"1\"");
                    }
                    if (this.cNvSpPr.spLocks.noGrp)
                    {
                        sw.Write(" noGrp=\"1\"");
                    }
                    if (this.cNvSpPr.spLocks.noMove)
                    {
                        sw.Write(" noMove=\"1\"");
                    }
                    if (this.cNvSpPr.spLocks.noResize)
                    {
                        sw.Write(" noResize=\"1\"");
                    }
                    if (this.cNvSpPr.spLocks.noRot)
                    {
                        sw.Write(" noRot=\"1\"");
                    }
                    if (this.cNvSpPr.spLocks.noSelect)
                    {
                        sw.Write(" noSelect=\"1\"");
                    }
                    sw.Write("/>");
                }
                sw.Write("</xdr:cNvSpPr>");
            }
            sw.Write("</xdr:nvSpPr>");
        }
    }
}
