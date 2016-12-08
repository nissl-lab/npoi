
using NPOI.OpenXml4Net.Util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Dml 
{
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_TextAnchoringType {
        
    
        t,
        
    
        ctr,
        
    
        b,
        
    
        just,
        
    
        dist,
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_TextVertOverflowType {
        
    
        overflow,
        
    
        ellipsis,
        
    
        clip,
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_TextHorzOverflowType {
        
    
        overflow,
        
    
        clip,
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_TextVerticalType {
        
    
        horz,
        
    
        vert,
        
    
        vert270,
        
    
        wordArtVert,
        
    
        eaVert,
        
    
        mongolianVert,
        
    
        wordArtVertRtl,
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_TextWrappingType {
        
    
        none,
        
    
        square,
    }
    

    [Serializable]
    
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public class CT_TextListStyle {
        
        private CT_TextParagraphProperties defPPrField;
        
        private CT_TextParagraphProperties lvl1pPrField;
        
        private CT_TextParagraphProperties lvl2pPrField;
        
        private CT_TextParagraphProperties lvl3pPrField;
        
        private CT_TextParagraphProperties lvl4pPrField;
        
        private CT_TextParagraphProperties lvl5pPrField;
        
        private CT_TextParagraphProperties lvl6pPrField;
        
        private CT_TextParagraphProperties lvl7pPrField;
        
        private CT_TextParagraphProperties lvl8pPrField;
        
        private CT_TextParagraphProperties lvl9pPrField;
        
        private CT_OfficeArtExtensionList extLstField;

        public static CT_TextListStyle Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TextListStyle ctObj = new CT_TextListStyle();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "defPPr")
                    ctObj.defPPr = CT_TextParagraphProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lvl1pPr")
                    ctObj.lvl1pPr = CT_TextParagraphProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lvl2pPr")
                    ctObj.lvl2pPr = CT_TextParagraphProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lvl3pPr")
                    ctObj.lvl3pPr = CT_TextParagraphProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lvl4pPr")
                    ctObj.lvl4pPr = CT_TextParagraphProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lvl5pPr")
                    ctObj.lvl5pPr = CT_TextParagraphProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lvl6pPr")
                    ctObj.lvl6pPr = CT_TextParagraphProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lvl7pPr")
                    ctObj.lvl7pPr = CT_TextParagraphProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lvl8pPr")
                    ctObj.lvl8pPr = CT_TextParagraphProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lvl9pPr")
                    ctObj.lvl9pPr = CT_TextParagraphProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_OfficeArtExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            sw.Write(">");
            if (this.defPPr != null)
                this.defPPr.Write(sw, "defPPr");
            if (this.lvl1pPr != null)
                this.lvl1pPr.Write(sw, "lvl1pPr");
            if (this.lvl2pPr != null)
                this.lvl2pPr.Write(sw, "lvl2pPr");
            if (this.lvl3pPr != null)
                this.lvl3pPr.Write(sw, "lvl3pPr");
            if (this.lvl4pPr != null)
                this.lvl4pPr.Write(sw, "lvl4pPr");
            if (this.lvl5pPr != null)
                this.lvl5pPr.Write(sw, "lvl5pPr");
            if (this.lvl6pPr != null)
                this.lvl6pPr.Write(sw, "lvl6pPr");
            if (this.lvl7pPr != null)
                this.lvl7pPr.Write(sw, "lvl7pPr");
            if (this.lvl8pPr != null)
                this.lvl8pPr.Write(sw, "lvl8pPr");
            if (this.lvl9pPr != null)
                this.lvl9pPr.Write(sw, "lvl9pPr");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            sw.Write(string.Format("</a:{0}>", nodeName));
        }

        public CT_TextParagraphProperties defPPr {
            get {
                return this.defPPrField;
            }
            set {
                this.defPPrField = value;
            }
        }
        
    
        public CT_TextParagraphProperties lvl1pPr {
            get {
                return this.lvl1pPrField;
            }
            set {
                this.lvl1pPrField = value;
            }
        }
        
    
        public CT_TextParagraphProperties lvl2pPr {
            get {
                return this.lvl2pPrField;
            }
            set {
                this.lvl2pPrField = value;
            }
        }
        
    
        public CT_TextParagraphProperties lvl3pPr {
            get {
                return this.lvl3pPrField;
            }
            set {
                this.lvl3pPrField = value;
            }
        }
        
    
        public CT_TextParagraphProperties lvl4pPr {
            get {
                return this.lvl4pPrField;
            }
            set {
                this.lvl4pPrField = value;
            }
        }
        
    
        public CT_TextParagraphProperties lvl5pPr {
            get {
                return this.lvl5pPrField;
            }
            set {
                this.lvl5pPrField = value;
            }
        }
        
    
        public CT_TextParagraphProperties lvl6pPr {
            get {
                return this.lvl6pPrField;
            }
            set {
                this.lvl6pPrField = value;
            }
        }
        
    
        public CT_TextParagraphProperties lvl7pPr {
            get {
                return this.lvl7pPrField;
            }
            set {
                this.lvl7pPrField = value;
            }
        }
        
    
        public CT_TextParagraphProperties lvl8pPr {
            get {
                return this.lvl8pPrField;
            }
            set {
                this.lvl8pPrField = value;
            }
        }
        
    
        public CT_TextParagraphProperties lvl9pPr {
            get {
                return this.lvl9pPrField;
            }
            set {
                this.lvl9pPrField = value;
            }
        }
        
    
        public CT_OfficeArtExtensionList extLst {
            get {
                return this.extLstField;
            }
            set {
                this.extLstField = value;
            }
        }
    }
    

    [Serializable]
    
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public class CT_TextNormalAutofit {
        
        private int fontScaleField;
        
        private int lnSpcReductionField;
        
        public CT_TextNormalAutofit() {
            this.fontScaleField = 100000;
            this.lnSpcReductionField = 0;
        }
        public static CT_TextNormalAutofit Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TextNormalAutofit ctObj = new CT_TextNormalAutofit();
            ctObj.fontScale = XmlHelper.ReadInt(node.Attributes["fontScale"]);
            ctObj.lnSpcReduction = XmlHelper.ReadInt(node.Attributes["lnSpcReduction"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "fontScale", this.fontScale);
            XmlHelper.WriteAttribute(sw, "lnSpcReduction", this.lnSpcReduction);
            sw.Write(">");
            sw.Write(string.Format("</a:{0}>", nodeName));
        }
    
        [XmlAttribute]
        [DefaultValue(100000)]
        public int fontScale {
            get {
                return this.fontScaleField;
            }
            set {
                this.fontScaleField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValue(0)]
        public int lnSpcReduction {
            get {
                return this.lnSpcReductionField;
            }
            set {
                this.lnSpcReductionField = value;
            }
        }
    }
    

    [Serializable]
    
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public class CT_TextShapeAutofit {
    }
    

    [Serializable]
    
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public class CT_TextNoAutofit {
    }


    [Serializable]
    
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_TextBodyProperties
    {

        private CT_PresetTextShape prstTxWarpField;

        private CT_TextNoAutofit noAutofitField;

        private CT_TextNormalAutofit normAutofitField;

        private CT_TextShapeAutofit spAutoFitField;

        private CT_Scene3D scene3dField;

        private CT_Shape3D sp3dField;

        private CT_FlatText flatTxField;

        private CT_OfficeArtExtensionList extLstField;

        private int rotField;

        private bool rotFieldSpecified;

        private bool spcFirstLastParaField;

        private bool spcFirstLastParaFieldSpecified;

        private ST_TextVertOverflowType vertOverflowField;

        private bool vertOverflowFieldSpecified;

        private ST_TextHorzOverflowType horzOverflowField;

        private bool horzOverflowFieldSpecified;

        private ST_TextVerticalType vertField;

        private bool vertFieldSpecified;

        private ST_TextWrappingType wrapField;

        private bool wrapFieldSpecified;

        private int lInsField;

        private bool lInsFieldSpecified;

        private int tInsField;

        private bool tInsFieldSpecified;

        private int rInsField;

        private bool rInsFieldSpecified;

        private int bInsField;

        private bool bInsFieldSpecified;

        private int numColField;

        private bool numColFieldSpecified;

        private int spcColField;

        private bool spcColFieldSpecified;

        private bool rtlColField;

        private bool rtlColFieldSpecified;

        private bool fromWordArtField;

        private bool fromWordArtFieldSpecified;

        private ST_TextAnchoringType anchorField;

        private bool anchorFieldSpecified;

        private bool anchorCtrField;

        private bool anchorCtrFieldSpecified;

        private bool forceAAField;

        private bool forceAAFieldSpecified;

        private bool uprightField;

        private bool compatLnSpcField;

        private bool compatLnSpcFieldSpecified;

        public CT_TextBodyProperties()
        {
            this.uprightField = false;
            this.vertField = ST_TextVerticalType.horz;
            this.wrapField = ST_TextWrappingType.none;
            this.spcFirstLastParaField = false;
        }
        public static CT_TextBodyProperties Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TextBodyProperties ctObj = new CT_TextBodyProperties();
            ctObj.rot = XmlHelper.ReadInt(node.Attributes["rot"]);
            ctObj.spcFirstLastPara = XmlHelper.ReadBool(node.Attributes["spcFirstLastPara"]);
            if (node.Attributes["vertOverflow"] != null)
                ctObj.vertOverflow = (ST_TextVertOverflowType)Enum.Parse(typeof(ST_TextVertOverflowType), node.Attributes["vertOverflow"].Value);
            if (node.Attributes["horzOverflow"] != null)
                ctObj.horzOverflow = (ST_TextHorzOverflowType)Enum.Parse(typeof(ST_TextHorzOverflowType), node.Attributes["horzOverflow"].Value);
            if (node.Attributes["vert"] != null)
                ctObj.vert = (ST_TextVerticalType)Enum.Parse(typeof(ST_TextVerticalType), node.Attributes["vert"].Value);
            if (node.Attributes["wrap"] != null)
                ctObj.wrap = (ST_TextWrappingType)Enum.Parse(typeof(ST_TextWrappingType), node.Attributes["wrap"].Value);
            ctObj.lIns = XmlHelper.ReadInt(node.Attributes["lIns"]);
            ctObj.tIns = XmlHelper.ReadInt(node.Attributes["tIns"]);
            ctObj.rIns = XmlHelper.ReadInt(node.Attributes["rIns"]);
            ctObj.bIns = XmlHelper.ReadInt(node.Attributes["bIns"]);
            ctObj.numCol = XmlHelper.ReadInt(node.Attributes["numCol"]);
            ctObj.spcCol = XmlHelper.ReadInt(node.Attributes["spcCol"]);
            ctObj.rtlCol = XmlHelper.ReadBool(node.Attributes["rtlCol"]);
            ctObj.fromWordArt = XmlHelper.ReadBool(node.Attributes["fromWordArt"]);
            if (node.Attributes["anchor"] != null)
                ctObj.anchor = (ST_TextAnchoringType)Enum.Parse(typeof(ST_TextAnchoringType), node.Attributes["anchor"].Value);
            ctObj.anchorCtr = XmlHelper.ReadBool(node.Attributes["anchorCtr"]);
            ctObj.forceAA = XmlHelper.ReadBool(node.Attributes["forceAA"]);
            ctObj.upright = XmlHelper.ReadBool(node.Attributes["upright"]);
            ctObj.compatLnSpc = XmlHelper.ReadBool(node.Attributes["compatLnSpc"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "prstTxWarp")
                    ctObj.prstTxWarp = CT_PresetTextShape.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "noAutofit")
                    ctObj.noAutofit = new CT_TextNoAutofit();
                else if (childNode.LocalName == "normAutofit")
                    ctObj.normAutofit = CT_TextNormalAutofit.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "spAutoFit")
                    ctObj.spAutoFit = new CT_TextShapeAutofit();
                else if (childNode.LocalName == "scene3d")
                    ctObj.scene3d = CT_Scene3D.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sp3d")
                    ctObj.sp3d = CT_Shape3D.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "flatTx")
                    ctObj.flatTx = CT_FlatText.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_OfficeArtExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "rot", this.rot,true);
            if(spcFirstLastPara)
                XmlHelper.WriteAttribute(sw, "spcFirstLastPara", this.spcFirstLastPara);
            XmlHelper.WriteAttribute(sw, "vertOverflow", this.vertOverflow.ToString());
            if(this.horzOverflow!= ST_TextHorzOverflowType.overflow)
                XmlHelper.WriteAttribute(sw, "horzOverflow", this.horzOverflow.ToString());
            if(this.vert!= ST_TextVerticalType.vert)
                XmlHelper.WriteAttribute(sw, "vert", this.vert.ToString());
            if(this.wrap!= ST_TextWrappingType.none)
                XmlHelper.WriteAttribute(sw, "wrap", this.wrap.ToString());
            XmlHelper.WriteAttribute(sw, "lIns", this.lIns);
            XmlHelper.WriteAttribute(sw, "tIns", this.tIns);
            XmlHelper.WriteAttribute(sw, "rIns", this.rIns);
            XmlHelper.WriteAttribute(sw, "bIns", this.bIns);
            XmlHelper.WriteAttribute(sw, "numCol", this.numCol);
            XmlHelper.WriteAttribute(sw, "spcCol", this.spcCol);
            XmlHelper.WriteAttribute(sw, "rtlCol", this.rtlCol);
            XmlHelper.WriteAttribute(sw, "fromWordArt", this.fromWordArt, false);
            XmlHelper.WriteAttribute(sw, "anchor", this.anchor.ToString());
            XmlHelper.WriteAttribute(sw, "anchorCtr", this.anchorCtr, false);
            XmlHelper.WriteAttribute(sw, "forceAA", this.forceAA, false);
            if(upright)
                XmlHelper.WriteAttribute(sw, "upright", this.upright);
            if (compatLnSpc)
                XmlHelper.WriteAttribute(sw, "compatLnSpc", this.compatLnSpc);
            sw.Write(">");
            if (this.prstTxWarp != null)
                this.prstTxWarp.Write(sw, "prstTxWarp");
            if (this.noAutofit != null)
                sw.Write("<a:noAutofit/>");
            if (this.normAutofit != null)
                this.normAutofit.Write(sw, "normAutofit");
            if (this.spAutoFit != null)
                sw.Write("<a:spAutoFit/>");
            if (this.scene3d != null)
                this.scene3d.Write(sw, "scene3d");
            if (this.sp3d != null)
                this.sp3d.Write(sw, "sp3d");
            if (this.flatTx != null)
                this.flatTx.Write(sw, "flatTx");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            sw.Write(string.Format("</a:{0}>", nodeName));
        }

        public CT_PresetTextShape prstTxWarp
        {
            get
            {
                return this.prstTxWarpField;
            }
            set
            {
                this.prstTxWarpField = value;
            }
        }

        public CT_TextNoAutofit noAutofit
        {
            get
            {
                return this.noAutofitField;
            }
            set
            {
                this.noAutofitField = value;
            }
        }

        public CT_TextNormalAutofit normAutofit
        {
            get
            {
                return this.normAutofitField;
            }
            set
            {
                this.normAutofitField = value;
            }
        }
        public CT_TextShapeAutofit spAutoFit
        {
            get
            {
                return this.spAutoFitField;
            }
            set
            {
                this.spAutoFitField = value;
            }
        }


        public CT_Scene3D scene3d
        {
            get
            {
                return this.scene3dField;
            }
            set
            {
                this.scene3dField = value;
            }
        }


        public CT_Shape3D sp3d
        {
            get
            {
                return this.sp3dField;
            }
            set
            {
                this.sp3dField = value;
            }
        }


        public CT_FlatText flatTx
        {
            get
            {
                return this.flatTxField;
            }
            set
            {
                this.flatTxField = value;
            }
        }


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
        public int rot
        {
            get
            {
                return this.rotField;
            }
            set
            {
                this.rotField = value;
                this.rotFieldSpecified = true;
            }
        }


        [XmlIgnore]
        public bool rotSpecified
        {
            get
            {
                return this.rotFieldSpecified;
            }
            set
            {
                this.rotFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public bool spcFirstLastPara
        {
            get
            {
                return this.spcFirstLastParaField;
            }
            set
            {
                this.spcFirstLastParaField = value;
                this.spcFirstLastParaFieldSpecified = value;
            }
        }


        [XmlIgnore]
        public bool spcFirstLastParaSpecified
        {
            get
            {
                return this.spcFirstLastParaFieldSpecified;
            }
            set
            {
                this.spcFirstLastParaFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public ST_TextVertOverflowType vertOverflow
        {
            get
            {
                return this.vertOverflowField;
            }
            set
            {
                this.vertOverflowField = value;
                this.vertOverflowFieldSpecified = true;
            }
        }


        [XmlIgnore]
        public bool vertOverflowSpecified
        {
            get
            {
                return this.vertOverflowFieldSpecified;
            }
            set
            {
                this.vertOverflowFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public ST_TextHorzOverflowType horzOverflow
        {
            get
            {
                return this.horzOverflowField;
            }
            set
            {
                this.horzOverflowField = value;
                this.horzOverflowFieldSpecified = true;
            }
        }


        [XmlIgnore]
        public bool horzOverflowSpecified
        {
            get
            {
                return this.horzOverflowFieldSpecified;
            }
            set
            {
                this.horzOverflowFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public ST_TextVerticalType vert
        {
            get
            {
                return this.vertField;
            }
            set
            {
                this.vertField = value;
                this.vertFieldSpecified = true;
            }
        }


        [XmlIgnore]
        public bool vertSpecified
        {
            get
            {
                return this.vertFieldSpecified;
            }
            set
            {
                this.vertFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public ST_TextWrappingType wrap
        {
            get
            {
                return this.wrapField;
            }
            set
            {
                this.wrapField = value;
                this.wrapFieldSpecified = true;
            }
        }


        [XmlIgnore]
        public bool wrapSpecified
        {
            get
            {
                return this.wrapFieldSpecified;
            }
            set
            {
                this.wrapFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public int lIns
        {
            get
            {
                return this.lInsField;
            }
            set
            {
                this.lInsField = value;
                this.lInsFieldSpecified = true;
            }
        }


        [XmlIgnore]
        public bool lInsSpecified
        {
            get
            {
                return this.lInsFieldSpecified;
            }
            set
            {
                this.lInsFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public int tIns
        {
            get
            {
                return this.tInsField;
            }
            set
            {
                this.tInsField = value;
                this.tInsFieldSpecified = true;
            }
        }


        [XmlIgnore]
        public bool tInsSpecified
        {
            get
            {
                return this.tInsFieldSpecified;
            }
            set
            {
                this.tInsFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public int rIns
        {
            get
            {
                return this.rInsField;
            }
            set
            {
                this.rInsField = value;
                this.rInsFieldSpecified = true;
            }
        }


        [XmlIgnore]
        public bool rInsSpecified
        {
            get
            {
                return this.rInsFieldSpecified;
            }
            set
            {
                this.rInsFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public int bIns
        {
            get
            {
                return this.bInsField;
            }
            set
            {
                this.bInsField = value;
                this.bInsFieldSpecified = true;
            }
        }


        [XmlIgnore]
        public bool bInsSpecified
        {
            get
            {
                return this.bInsFieldSpecified;
            }
            set
            {
                this.bInsFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public int numCol
        {
            get
            {
                return this.numColField;
            }
            set
            {
                this.numColField = value;
                this.numColFieldSpecified = true;
            }
        }


        [XmlIgnore]
        public bool numColSpecified
        {
            get
            {
                return this.numColFieldSpecified;
            }
            set
            {
                this.numColFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public int spcCol
        {
            get
            {
                return this.spcColField;
            }
            set
            {
                this.spcColField = value;
                this.spcColFieldSpecified = true;
            }
        }


        [XmlIgnore]
        public bool spcColSpecified
        {
            get
            {
                return this.spcColFieldSpecified;
            }
            set
            {
                this.spcColFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public bool rtlCol
        {
            get
            {
                return this.rtlColField;
            }
            set
            {
                this.rtlColField = value;
                this.rtlColFieldSpecified = value;
            }
        }


        [XmlIgnore]
        public bool rtlColSpecified
        {
            get
            {
                return this.rtlColFieldSpecified;
            }
            set
            {
                this.rtlColFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public bool fromWordArt
        {
            get
            {
                return this.fromWordArtField;
            }
            set
            {
                this.fromWordArtField = value;
                this.fromWordArtFieldSpecified = value;
            }
        }


        [XmlIgnore]
        public bool fromWordArtSpecified
        {
            get
            {
                return this.fromWordArtFieldSpecified;
            }
            set
            {
                this.fromWordArtFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public ST_TextAnchoringType anchor
        {
            get
            {
                return this.anchorField;
            }
            set
            {
                this.anchorField = value;
                this.anchorFieldSpecified = true;
            }
        }


        [XmlIgnore]
        public bool anchorSpecified
        {
            get
            {
                return this.anchorFieldSpecified;
            }
            set
            {
                this.anchorFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public bool anchorCtr
        {
            get
            {
                return this.anchorCtrField;
            }
            set
            {
                this.anchorCtrField = value;
                this.anchorCtrFieldSpecified = value;
            }
        }


        [XmlIgnore]
        public bool anchorCtrSpecified
        {
            get
            {
                return this.anchorCtrFieldSpecified;
            }
            set
            {
                this.anchorCtrFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public bool forceAA
        {
            get
            {
                return this.forceAAField;
            }
            set
            {
                this.forceAAField = value;
                this.forceAAFieldSpecified = value;
            }
        }


        [XmlIgnore]
        public bool forceAASpecified
        {
            get
            {
                return this.forceAAFieldSpecified;
            }
            set
            {
                this.forceAAFieldSpecified = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(false)]
        public bool upright
        {
            get
            {
                return this.uprightField;
            }
            set
            {
                this.uprightField = value;
            }
        }


        [XmlAttribute]
        public bool compatLnSpc
        {
            get
            {
                return this.compatLnSpcField;
            }
            set
            {
                this.compatLnSpcField = value;
                this.compatLnSpcFieldSpecified = value;
            }
        }


        [XmlIgnore]
        public bool compatLnSpcSpecified
        {
            get
            {
                return this.compatLnSpcFieldSpecified;
            }
            set
            {
                this.compatLnSpcFieldSpecified = value;
            }
        }

        public void UnsetTIns()
        {
            this.tInsFieldSpecified = false;
        }

        public void UnsetVertOverflow()
        {
            this.vertOverflowFieldSpecified = false;
        }

        public void UnsetVert()
        {
            this.vertFieldSpecified = false;
        }

        public bool IsSetVert()
        {
            return this.vertFieldSpecified;
        }

        public bool IsSetBIns()
        {
            return this.bInsFieldSpecified;
        }

        public bool IsSetLIns()
        {
            return this.lInsFieldSpecified;
        }

        public bool IsSetRIns()
        {
            return this.rInsFieldSpecified;
        }

        public bool IsSetTIns()
        {
            return this.tInsFieldSpecified;
        }

        public void UnsetBIns()
        {
            this.bInsFieldSpecified = false;
        }

        public void UnsetLIns()
        {
            this.lInsFieldSpecified = false;
        }

        public void UnsetRIns()
        {
            this.rInsFieldSpecified = false;
        }

        public bool IsSetSpAutoFit()
        {
            return this.spAutoFitField != null;
        }

        public bool IsSetNoAutofit()
        {
            return this.noAutofitField != null;
        }

        public bool IsSetNormAutofit()
        {
            return this.normAutofitField != null;
        }

        public void UnsetSpAutoFit()
        {
            this.spAutoFitField = null;
        }

        public void UnsetNoAutofit()
        {
            this.noAutofitField = null;
        }

        public void UnsetNormAutofit()
        {
            this.normAutofitField = null;
        }

        public CT_TextNoAutofit AddNewNoAutofit()
        {
            this.noAutofitField = new CT_TextNoAutofit();
            return this.noAutofitField;
        }

        public CT_TextNormalAutofit AddNewNormAutofit()
        {
            this.normAutofitField = new CT_TextNormalAutofit();
            return this.normAutofitField;
        }

        public CT_TextShapeAutofit AddNewSpAutoFit()
        {
            this.spAutoFitField = new CT_TextShapeAutofit();
            return this.spAutoFitField;
        }

        public void UnsetHorzOverflow()
        {
            this.horzOverflowFieldSpecified = false;
        }

        public bool IsSetHorzOverflow()
        {
            return this.horzOverflowFieldSpecified;
        }


        public bool IsSetVertOverflow()
        {
            return this.vertOverflowFieldSpecified;
        }

        public bool IsSetAnchor()
        {
            return this.anchorFieldSpecified;
        }

        public void UnsetAnchor()
        {
            this.anchorFieldSpecified = false;
        }

        public bool IsSetWrap()
        {
            return this.wrapFieldSpecified;
        }
        public void UnsetWrap()
        {
            this.wrapFieldSpecified = false;
        }
    }
    

    [Serializable]
    
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public class CT_TextBody {
        
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
            sw.Write(string.Format("<a:{0}", nodeName));
            sw.Write(">");
            if (this.bodyPr != null)
                this.bodyPr.Write(sw, "bodyPr");
            if (this.lstStyle != null)
                this.lstStyle.Write(sw, "lstStyle");
            foreach (CT_TextParagraph x in this.p)
            {
                x.Write(sw, "p");
            }
            sw.Write(string.Format("</a:{0}>", nodeName));
        }

        public void SetPArray(CT_TextParagraph[] array)
        {
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
                this.lstStyleField=new CT_TextListStyle();
            return this.lstStyleField;   
        }
    
        public CT_TextBodyProperties bodyPr {
            get {
                return this.bodyPrField;
            }
            set {
                this.bodyPrField = value;
            }
        }
        
    
        public CT_TextListStyle lstStyle {
            get {
                return this.lstStyleField;
            }
            set {
                this.lstStyleField = value;
            }
        }
        public override string ToString()
        {
            if (p == null||p.Count==0)
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
        public List<CT_TextParagraph> p {
            get {
                return this.pField;
            }
            set {
                this.pField = value;
            }
        }
    }
}
