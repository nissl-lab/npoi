using NPOI.OpenXml4Net.Util;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
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
            CT_Shape ctObj = new CT_Shape();
            ctObj.macro = XmlHelper.ReadString(node.Attributes["macro"]);
            ctObj.textlink = XmlHelper.ReadString(node.Attributes["textlink"]);
            ctObj.fLocksText = XmlHelper.ReadBool(node.Attributes["fLocksText"]);
            ctObj.fPublished = XmlHelper.ReadBool(node.Attributes["fPublished"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "nvSpPr")
                    ctObj.nvSpPr = CT_ShapeNonVisual.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "spPr")
                    ctObj.spPr = CT_ShapeProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "txBody")
                    ctObj.txBody = CT_TextBody.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "style")
                    ctObj.style = CT_ShapeStyle.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "macro", this.macro, true);
            XmlHelper.WriteAttribute(sw, "textlink", this.textlink,true);
            XmlHelper.WriteAttribute(sw, "fLocksText", this.fLocksText, false);
            XmlHelper.WriteAttribute(sw, "fPublished", this.fPublished, false);
            sw.Write(">");
            if (this.nvSpPr != null)
                this.nvSpPr.Write(sw, "nvSpPr");
            if (this.spPr != null)
                this.spPr.Write(sw, "spPr");
            if (this.style != null)
                this.style.Write(sw, "style");
            if (this.txBody != null)
                this.txBody.Write(sw, "txBody");
            sw.Write(string.Format("</xdr:{0}>", nodeName));
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

    }

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_TextBody
    {

        private CT_TextBodyProperties bodyPrField;

        private CT_TextListStyle lstStyleField;

        private List<CT_TextParagraph> pField;

        public static CT_TextBody Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TextBody ctObj = new CT_TextBody();
            ctObj.p = new List<CT_TextParagraph>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "bodyPr")
                    ctObj.bodyPr = CT_TextBodyProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lstStyle")
                    ctObj.lstStyle = CT_TextListStyle.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "p")
                    ctObj.p.Add(CT_TextParagraph.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}", nodeName));
            sw.Write(">");
            if (this.bodyPr != null)
                this.bodyPr.Write(sw, "bodyPr");
            if (this.lstStyle != null)
                this.lstStyle.Write(sw, "lstStyle");
            foreach (CT_TextParagraph x in this.p)
            {
                x.Write(sw, "p");
            }
            sw.Write(string.Format("</xdr:{0}>", nodeName));
        }

        public void SetPArray(CT_TextParagraph[] array)
        {
            if (array == null)
                pField.Clear();
            else
                pField = new List<CT_TextParagraph>(array);
        }
        public CT_TextParagraph AddNewP()
        {
            if (this.pField == null)
                pField = new List<CT_TextParagraph>();
            CT_TextParagraph tp = new CT_TextParagraph();
            pField.Add(tp);
            return tp;
        }
        public CT_TextBodyProperties AddNewBodyPr()
        {
            this.bodyPrField = new CT_TextBodyProperties();
            return this.bodyPrField;
        }
        public CT_TextListStyle AddNewLstStyle()
        {
            this.lstStyleField = new CT_TextListStyle();
            return this.lstStyleField;
        }

        public CT_TextBodyProperties bodyPr
        {
            get
            {
                return this.bodyPrField;
            }
            set
            {
                this.bodyPrField = value;
            }
        }


        public CT_TextListStyle lstStyle
        {
            get
            {
                return this.lstStyleField;
            }
            set
            {
                this.lstStyleField = value;
            }
        }
        public override string ToString()
        {
            if (p == null || p.Count == 0)
                return string.Empty;
            StringBuilder sb = new StringBuilder();
            foreach (CT_TextParagraph tp in p)
            {
                foreach (CT_RegularTextRun tr in tp.r)
                {
                    sb.Append(tr.t);
                }
            }
            return sb.ToString();
        }

        [XmlElement("p")]
        public List<CT_TextParagraph> p
        {
            get
            {
                return this.pField;
            }
            set
            {
                this.pField = value;
            }
        }

        public CT_TextParagraph GetPArray(int pos)
        {
            if (p == null)
                return null;
            return p[pos];
        }

        public int SizeOfPArray()
        {
            return p.Count;
        }
    }

    [Serializable]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_ShapeStyle
    {

        private CT_StyleMatrixReference lnRefField;

        private CT_StyleMatrixReference fillRefField;

        private CT_StyleMatrixReference effectRefField;

        private CT_FontReference fontRefField;
        public static CT_ShapeStyle Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ShapeStyle ctObj = new CT_ShapeStyle();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "lnRef")
                    ctObj.lnRef = CT_StyleMatrixReference.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "fillRef")
                    ctObj.fillRef = CT_StyleMatrixReference.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "effectRef")
                    ctObj.effectRef = CT_StyleMatrixReference.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "fontRef")
                    ctObj.fontRef = CT_FontReference.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}", nodeName));
            sw.Write(">");
            if (this.lnRef != null)
                this.lnRef.Write(sw, "lnRef");
            if (this.fillRef != null)
                this.fillRef.Write(sw, "fillRef");
            if (this.effectRef != null)
                this.effectRef.Write(sw, "effectRef");
            if (this.fontRef != null)
                this.fontRef.Write(sw, "fontRef");
            sw.Write(string.Format("</xdr:{0}>", nodeName));
        }

        public CT_StyleMatrixReference AddNewFillRef()
        {
            this.fillRefField = new CT_StyleMatrixReference();
            return this.fillRefField;
        }
        public CT_StyleMatrixReference AddNewLnRef()
        {
            this.lnRefField = new CT_StyleMatrixReference();
            return this.lnRefField;
        }
        public CT_FontReference AddNewFontRef()
        {
            this.fontRefField = new CT_FontReference();
            return this.fontRefField;
        }
        public CT_StyleMatrixReference AddNewEffectRef()
        {
            this.effectRefField = new CT_StyleMatrixReference();
            return this.effectRefField;
        }
        [XmlElement(Order = 0)]
        public CT_StyleMatrixReference lnRef
        {
            get
            {
                return this.lnRefField;
            }
            set
            {
                this.lnRefField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_StyleMatrixReference fillRef
        {
            get
            {
                return this.fillRefField;
            }
            set
            {
                this.fillRefField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_StyleMatrixReference effectRef
        {
            get
            {
                return this.effectRefField;
            }
            set
            {
                this.effectRefField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_FontReference fontRef
        {
            get
            {
                return this.fontRefField;
            }
            set
            {
                this.fontRefField = value;
            }
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
        public static CT_ShapeNonVisual Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ShapeNonVisual ctObj = new CT_ShapeNonVisual();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "cNvPr")
                    ctObj.cNvPr = CT_NonVisualDrawingProps.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cNvSpPr")
                    ctObj.cNvSpPr = CT_NonVisualDrawingShapeProps.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}", nodeName));
            sw.Write(">");
            if (this.cNvPr != null)
                this.cNvPr.Write(sw, "cNvPr");
            if (this.cNvSpPr != null)
                this.cNvSpPr.Write(sw, "cNvSpPr");
            sw.Write(string.Format("</xdr:{0}>", nodeName));
        }


    }
    [Serializable]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_NonVisualDrawingShapeProps
    {

        private CT_ShapeLocking spLocksField;

        private CT_OfficeArtExtensionList extLstField;

        private bool txBoxField;
        public static CT_NonVisualDrawingShapeProps Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_NonVisualDrawingShapeProps ctObj = new CT_NonVisualDrawingShapeProps();
            ctObj.txBox = XmlHelper.ReadBool(node.Attributes["txBox"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "spLocks")
                    ctObj.spLocks = CT_ShapeLocking.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_OfficeArtExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "txBox", this.txBox, false);
            sw.Write(">");
            if (this.spLocks != null)
                this.spLocks.Write(sw, "spLocks");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            sw.Write(string.Format("</xdr:{0}>", nodeName));
        }



        [XmlElement(Order = 0)]
        public CT_ShapeLocking spLocks
        {
            get
            {
                return this.spLocksField;
            }
            set
            {
                this.spLocksField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_OfficeArtExtensionList extLst
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

        [XmlAttribute]
        [DefaultValue(false)]
        public bool txBox
        {
            get
            {
                return this.txBoxField;
            }
            set
            {
                this.txBoxField = value;
            }
        }
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_GroupShape
    {
        CT_GroupShapeProperties grpSpPrField;
        CT_GroupShapeNonVisual nvGrpSpPrField;
        CT_Connector connectorField = null;
        CT_Picture pictureField = null;
        CT_Shape shapeField = null;

        public void Set(CT_GroupShape groupShape)
        {
            this.grpSpPrField = groupShape.grpSpPr;
            this.nvGrpSpPrField = groupShape.nvGrpSpPr;
        }

        public CT_GroupShapeProperties AddNewGrpSpPr()
        {
            this.grpSpPrField = new CT_GroupShapeProperties();
            return this.grpSpPrField;
        }
        public CT_GroupShapeNonVisual AddNewNvGrpSpPr()
        {
            this.nvGrpSpPrField = new CT_GroupShapeNonVisual();
            return this.nvGrpSpPrField;
        }
        public CT_Connector AddNewCxnSp()
        {
            connectorField = new CT_Connector();
            return connectorField;
        }
        public CT_Shape AddNewSp()
        {
            shapeField = new CT_Shape();
            return shapeField;
        }
        public CT_Picture AddNewPic()
        {
            pictureField = new CT_Picture();
            return pictureField;
        }

        public CT_GroupShapeNonVisual nvGrpSpPr
        {
            get { return nvGrpSpPrField; }
            set { nvGrpSpPrField = value; }
        }
        public CT_GroupShapeProperties grpSpPr
        {
            get { return grpSpPrField; }
            set { grpSpPrField = value; }

        }
        public static CT_GroupShape Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_GroupShape ctObj = new CT_GroupShape();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "nvGrpSpPr")
                    ctObj.nvGrpSpPr = CT_GroupShapeNonVisual.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "grpSpPr")
                    ctObj.grpSpPr = CT_GroupShapeProperties.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}", nodeName));
            sw.Write(">");
            if (this.nvGrpSpPr != null)
                this.nvGrpSpPr.Write(sw, "nvGrpSpPr");
            if (this.grpSpPr != null)
                this.grpSpPr.Write(sw, "grpSpPr");
            sw.Write(string.Format("</xdr:{0}>", nodeName));
        }

    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_GroupShapeNonVisual
    {
        CT_NonVisualDrawingProps cNvPrField;
        CT_NonVisualGroupDrawingShapeProps cNvGrpSpPrField;
        public static CT_GroupShapeNonVisual Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_GroupShapeNonVisual ctObj = new CT_GroupShapeNonVisual();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "cNvPr")
                    ctObj.cNvPr = CT_NonVisualDrawingProps.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cNvGrpSpPr")
                    ctObj.cNvGrpSpPr = CT_NonVisualGroupDrawingShapeProps.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}", nodeName));
            sw.Write(">");
            if (this.cNvPr != null)
                this.cNvPr.Write(sw, "cNvPr");
            if (this.cNvGrpSpPr != null)
                this.cNvGrpSpPr.Write(sw, "cNvGrpSpPr");
            sw.Write(string.Format("</xdr:{0}>", nodeName));
        }

        public CT_NonVisualGroupDrawingShapeProps AddNewCNvGrpSpPr()
        {
            this.cNvGrpSpPrField = new CT_NonVisualGroupDrawingShapeProps();
            return this.cNvGrpSpPrField;
        }
        public CT_NonVisualDrawingProps AddNewCNvPr()
        {
            this.cNvPrField = new CT_NonVisualDrawingProps();
            return this.cNvPrField;
        }

        public CT_NonVisualDrawingProps cNvPr
        {
            get { return cNvPrField; }
            set { cNvPrField = value; }
        }
        public CT_NonVisualGroupDrawingShapeProps cNvGrpSpPr
        {
            get { return cNvGrpSpPrField; }
            set { cNvGrpSpPrField = value; }
        }
    }

}
