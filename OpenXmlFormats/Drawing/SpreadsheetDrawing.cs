using System;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using NPOI.OpenXml4Net.Util;

namespace NPOI.OpenXmlFormats.Dml.Spreadsheet
{
    [Serializable]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_NonVisualDrawingProps
    {
        public static CT_NonVisualDrawingProps Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_NonVisualDrawingProps ctObj = new CT_NonVisualDrawingProps();
            ctObj.id = XmlHelper.ReadUInt(node.Attributes["id"]);
            ctObj.name = XmlHelper.ReadString(node.Attributes["name"]);
            ctObj.descr = XmlHelper.ReadString(node.Attributes["descr"]);
            ctObj.hidden = XmlHelper.ReadBool(node.Attributes["hidden"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "hlinkClick")
                    ctObj.hlinkClick = CT_Hyperlink.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "hlinkHover")
                    ctObj.hlinkHover = CT_Hyperlink.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_OfficeArtExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "id", this.id, true);
            XmlHelper.WriteAttribute(sw, "name", this.name);
            XmlHelper.WriteAttribute(sw, "descr", this.descr);
            XmlHelper.WriteAttribute(sw, "hidden", this.hidden, false);
            sw.Write(">");
            if (this.hlinkClick != null)
                this.hlinkClick.Write(sw, "hlinkClick");
            if (this.hlinkHover != null)
                this.hlinkHover.Write(sw, "hlinkHover");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            sw.Write(string.Format("</xdr:{0}>", nodeName));
        }

        private CT_Hyperlink hlinkClickField = null;

        private CT_Hyperlink hlinkHoverField = null;

        private CT_OfficeArtExtensionList extLstField = null;

        private uint idField;

        private string nameField = null;

        private string descrField;

        private bool? hiddenField = null;

        [XmlElement(Order = 0)]
        public CT_Hyperlink hlinkClick
        {
            get
            {
                return this.hlinkClickField;
            }
            set
            {
                this.hlinkClickField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_Hyperlink hlinkHover
        {
            get
            {
                return this.hlinkHoverField;
            }
            set
            {
                this.hlinkHoverField = value;
            }
        }

        [XmlElement(Order = 2)]
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
        public uint id
        {
            get
            {
                return this.idField;
            }
            set
            {
                this.idField = value;
            }
        }

        [XmlAttribute]
        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue("")]
        public string descr
        {
            get
            {
                return null == this.descrField ? "" : descrField;
            }
            set
            {
                this.descrField = value;
            }
        }
        [XmlIgnore]
        public bool descrSpecified
        {
            get { return (null != descrField); }
        }
        [XmlAttribute]
        [DefaultValue(false)]
        public bool hidden
        {
            get
            {
                return null == this.hiddenField ? false : (bool)hiddenField;
            }
            set
            {
                this.hiddenField = value;
            }
        }

        [XmlIgnore]
        public bool hiddenSpecified
        {
            get { return (null != hiddenField); }
        }
    }
    [Serializable]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_NonVisualGraphicFrameProperties
    {

        private CT_GraphicalObjectFrameLocking graphicFrameLocksField;

        private CT_OfficeArtExtensionList extLstField;

        public CT_NonVisualGraphicFrameProperties()
        {
            //this.extLstField = new CT_OfficeArtExtensionList();
            //this.graphicFrameLocksField = new CT_GraphicalObjectFrameLocking();
        }
        public static CT_NonVisualGraphicFrameProperties Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_NonVisualGraphicFrameProperties ctObj = new CT_NonVisualGraphicFrameProperties();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "graphicFrameLocks")
                    ctObj.graphicFrameLocks = CT_GraphicalObjectFrameLocking.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_OfficeArtExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}", nodeName));
            sw.Write(">");
            if (this.graphicFrameLocks != null)
                this.graphicFrameLocks.Write(sw, "graphicFrameLocks");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            sw.Write(string.Format("</xdr:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_GraphicalObjectFrameLocking graphicFrameLocks
        {
            get
            {
                return this.graphicFrameLocksField;
            }
            set
            {
                this.graphicFrameLocksField = value;
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
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_GraphicalObjectFrameNonVisual
    {

        CT_NonVisualDrawingProps cNvPrField;
        CT_NonVisualGraphicFrameProperties cNvGraphicFramePrField;
        public static CT_GraphicalObjectFrameNonVisual Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_GraphicalObjectFrameNonVisual ctObj = new CT_GraphicalObjectFrameNonVisual();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "cNvPr")
                    ctObj.cNvPr = CT_NonVisualDrawingProps.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cNvGraphicFramePr")
                    ctObj.cNvGraphicFramePr = CT_NonVisualGraphicFrameProperties.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}", nodeName));
            sw.Write(">");
            if (this.cNvPr != null)
                this.cNvPr.Write(sw, "cNvPr");
            if (this.cNvGraphicFramePr != null)
                this.cNvGraphicFramePr.Write(sw, "cNvGraphicFramePr");
            sw.Write(string.Format("</xdr:{0}>", nodeName));
        }

        public CT_NonVisualDrawingProps AddNewCNvPr()
        {
            this.cNvPrField = new CT_NonVisualDrawingProps();
            return this.cNvPrField;
        }
        public CT_NonVisualGraphicFrameProperties AddNewCNvGraphicFramePr()
        {
            this.cNvGraphicFramePrField = new CT_NonVisualGraphicFrameProperties();
            return this.cNvGraphicFramePrField;
        }

        public CT_NonVisualDrawingProps cNvPr
        {
            get { return cNvPrField; }
            set { cNvPrField = value; }
        }
        public CT_NonVisualGraphicFrameProperties cNvGraphicFramePr
        {
            get { return cNvGraphicFramePrField; }
            set { cNvGraphicFramePrField = value; }
        }
    }
    [Serializable]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_NonVisualPictureProperties
    {
        public static CT_NonVisualPictureProperties Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_NonVisualPictureProperties ctObj = new CT_NonVisualPictureProperties();
            ctObj.preferRelativeResize = XmlHelper.ReadBool(node.Attributes["preferRelativeResize"], true);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "picLocks")
                    ctObj.picLocks = CT_PictureLocking.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_OfficeArtExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}", nodeName));
            if(!preferRelativeResize)
                XmlHelper.WriteAttribute(sw, "preferRelativeResize", this.preferRelativeResize);
            sw.Write(">");
            if (this.picLocks != null)
                this.picLocks.Write(sw, "picLocks");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            sw.Write(string.Format("</xdr:{0}>", nodeName));
        }

        private CT_PictureLocking picLocksField = null;

        private CT_OfficeArtExtensionList extLstField = null;

        private bool preferRelativeResizeField = true;

        public CT_NonVisualPictureProperties()
        {
        }

        public CT_PictureLocking AddNewPicLocks()
        {
            this.picLocksField = new CT_PictureLocking();
            return picLocksField;
        }

        [XmlElement(Order = 0)]
        public CT_PictureLocking picLocks
        {
            get
            {
                return this.picLocksField;
            }
            set
            {
                this.picLocksField = value;
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
        public bool preferRelativeResize
        {
            get
            {
                return this.preferRelativeResizeField;
            }
            set
            {
                this.preferRelativeResizeField = value;
            }
        }
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", IsNullable = true)]
    public class CT_BlipFillProperties
    {

        private CT_Blip blipField = null;

        private CT_RelativeRect srcRectField = null;

        private CT_TileInfoProperties tileField = null;

        private CT_StretchInfoProperties stretchField = null;

        private uint dpiField;
        private bool dpiFieldSpecified;

        private bool rotWithShapeField;

        private bool rotWithShapeFieldSpecified;

        public CT_Blip AddNewBlip()
        {
            this.blipField = new CT_Blip();
            return blipField;
        }

        public CT_StretchInfoProperties AddNewStretch()
        {
            this.stretchField = new CT_StretchInfoProperties();
            return stretchField;
        }

        public static CT_BlipFillProperties Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_BlipFillProperties ctObj = new CT_BlipFillProperties();
            ctObj.dpi = XmlHelper.ReadUInt(node.Attributes["dpi"]);
            ctObj.rotWithShape = XmlHelper.ReadBool(node.Attributes["rotWithShape"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "blip")
                    ctObj.blip = CT_Blip.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "srcRect")
                    ctObj.srcRect = CT_RelativeRect.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tile")
                    ctObj.tile = CT_TileInfoProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "stretch")
                    ctObj.stretch = CT_StretchInfoProperties.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "dpi", this.dpi);
            if(rotWithShape)
                XmlHelper.WriteAttribute(sw, "rotWithShape", this.rotWithShape);
            sw.Write(">");
            if (this.blip != null)
                this.blip.Write(sw, "blip");
            if (this.srcRect != null)
                this.srcRect.Write(sw, "srcRect");
            if (this.tile != null)
                this.tile.Write(sw, "tile");
            if (this.stretch != null)
                this.stretch.Write(sw, "stretch");
            sw.Write(string.Format("</xdr:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_Blip blip
        {
            get
            {
                return this.blipField;
            }
            set
            {
                this.blipField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_RelativeRect srcRect
        {
            get
            {
                return this.srcRectField;
            }
            set
            {
                this.srcRectField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_TileInfoProperties tile
        {
            get
            {
                return this.tileField;
            }
            set
            {
                this.tileField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_StretchInfoProperties stretch
        {
            get
            {
                return this.stretchField;
            }
            set
            {
                this.stretchField = value;
            }
        }

        [XmlAttribute]
        public uint dpi
        {
            get
            {
                return (uint)this.dpiField;
            }
            set
            {
                this.dpiField = value;
            }
        }
        [XmlIgnore]
        public bool dpiSpecified
        {
            get
            {
                return dpiFieldSpecified;
            }
            set
            {
                this.dpiFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public bool rotWithShape
        {
            get
            {
                return (bool)this.rotWithShapeField;
            }
            set
            {
                this.rotWithShapeField = value;
            }
        }
        [XmlIgnore]
        public bool rotWithShapeSpecified
        {
            get
            {
                return rotWithShapeFieldSpecified;
            }
            set
            {
                this.rotWithShapeFieldSpecified = value;
            }
        }

        public bool IsSetBlip()
        {
            return this.blipField != null;
        }

    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_ShapeProperties
    {

        private CT_Transform2D xfrmField = null;

        private CT_CustomGeometry2D custGeomField = null;

        private CT_PresetGeometry2D prstGeomField = null;

        private CT_NoFillProperties noFillField = null;

        private CT_SolidColorFillProperties solidFillField = null;

        private CT_GradientFillProperties gradFillField = null;

        private CT_BlipFillProperties blipFillField = null;

        private CT_PatternFillProperties pattFillField = null;

        private CT_GroupFillProperties grpFillField = null;

        private CT_LineProperties lnField = null;

        private CT_EffectList effectLstField = null;

        private CT_EffectContainer effectDagField = null;

        private CT_Scene3D scene3dField = null;

        private CT_Shape3D sp3dField = null;

        private CT_OfficeArtExtensionList extLstField = null;

        private ST_BlackWhiteMode bwModeField = ST_BlackWhiteMode.none;


        public CT_PresetGeometry2D AddNewPrstGeom()
        {
            this.prstGeomField = new CT_PresetGeometry2D();
            return this.prstGeomField;
        }
        public CT_Transform2D AddNewXfrm()
        {
            this.xfrmField = new CT_Transform2D();
            return this.xfrmField;
        }
        public CT_SolidColorFillProperties AddNewSolidFill()
        {
            this.solidFillField = new CT_SolidColorFillProperties();
            return this.solidFillField;
        }
        public bool IsSetPattFill()
        {
            return this.pattFillField != null;
        }
        public bool IsSetSolidFill()
        {
            return this.solidFillField != null;
        }
        public bool IsSetLn()
        {
            return this.lnField != null;
        }
        public CT_LineProperties AddNewLn()
        {
            this.lnField = new CT_LineProperties();
            return lnField;
        }
        public void unsetPattFill()
        {
            this.pattFill = null;
        }
        public void unsetSolidFill()
        {
            this.solidFill = null;
        }

        [XmlElement(Order = 0)]
        public CT_Transform2D xfrm
        {
            get
            {
                return this.xfrmField;
            }
            set
            {
                this.xfrmField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_CustomGeometry2D custGeom
        {
            get
            {
                return this.custGeomField;
            }
            set
            {
                this.custGeomField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_PresetGeometry2D prstGeom
        {
            get
            {
                return this.prstGeomField;
            }
            set
            {
                this.prstGeomField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_NoFillProperties noFill
        {
            get
            {
                return this.noFillField;
            }
            set
            {
                this.noFillField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_SolidColorFillProperties solidFill
        {
            get
            {
                return this.solidFillField;
            }
            set
            {
                this.solidFillField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_GradientFillProperties gradFill
        {
            get
            {
                return this.gradFillField;
            }
            set
            {
                this.gradFillField = value;
            }
        }

        [XmlElement(Order = 6)]
        public CT_BlipFillProperties blipFill
        {
            get
            {
                return this.blipFillField;
            }
            set
            {
                this.blipFillField = value;
            }
        }

        [XmlElement(Order = 7)]
        public CT_PatternFillProperties pattFill
        {
            get
            {
                return this.pattFillField;
            }
            set
            {
                this.pattFillField = value;
            }
        }

        [XmlElement(Order = 8)]
        public CT_GroupFillProperties grpFill
        {
            get
            {
                return this.grpFillField;
            }
            set
            {
                this.grpFillField = value;
            }
        }

        [XmlElement(Order = 9)]
        public CT_LineProperties ln
        {
            get
            {
                return this.lnField;
            }
            set
            {
                this.lnField = value;
            }
        }

        [XmlElement(Order = 10)]
        public CT_EffectList effectLst
        {
            get
            {
                return this.effectLstField;
            }
            set
            {
                this.effectLstField = value;
            }
        }

        [XmlElement(Order = 11)]
        public CT_EffectContainer effectDag
        {
            get
            {
                return this.effectDagField;
            }
            set
            {
                this.effectDagField = value;
            }
        }

        [XmlElement(Order = 12)]
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

        [XmlElement(Order = 13)]
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

        [XmlElement(Order = 14)]
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
        public ST_BlackWhiteMode bwMode
        {
            get
            {
                return this.bwModeField;
            }
            set
            {
                this.bwModeField = value;
            }
        }
        [XmlIgnore]
        public bool bwModeSpecified
        {
            get { return ST_BlackWhiteMode.none != this.bwModeField; }
        }

        public static CT_ShapeProperties Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ShapeProperties ctObj = new CT_ShapeProperties();
            if (node.Attributes["bwMode"] != null)
                ctObj.bwMode = (ST_BlackWhiteMode)Enum.Parse(typeof(ST_BlackWhiteMode), node.Attributes["bwMode"].Value);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "xfrm")
                    ctObj.xfrm = CT_Transform2D.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "custGeom")
                    ctObj.custGeom = CT_CustomGeometry2D.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "prstGeom")
                    ctObj.prstGeom = CT_PresetGeometry2D.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "noFill")
                    ctObj.noFill = new CT_NoFillProperties();
                else if (childNode.LocalName == "solidFill")
                    ctObj.solidFill = CT_SolidColorFillProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "gradFill")
                    ctObj.gradFill = CT_GradientFillProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "blipFill")
                    ctObj.blipFill = CT_BlipFillProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pattFill")
                    ctObj.pattFill = CT_PatternFillProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "grpFill")
                    ctObj.grpFill = new CT_GroupFillProperties();
                else if (childNode.LocalName == "ln")
                    ctObj.ln = CT_LineProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "effectLst")
                    ctObj.effectLst = CT_EffectList.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "effectDag")
                    ctObj.effectDag = CT_EffectContainer.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "scene3d")
                    ctObj.scene3d = CT_Scene3D.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sp3d")
                    ctObj.sp3d = CT_Shape3D.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_OfficeArtExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}", nodeName));
            if(bwMode!= ST_BlackWhiteMode.none)
                XmlHelper.WriteAttribute(sw, "bwMode", this.bwMode.ToString());
            sw.Write(">");
            if (this.xfrm != null)
                this.xfrm.Write(sw, "a:xfrm");
            if (this.custGeom != null)
                this.custGeom.Write(sw, "custGeom");
            if (this.prstGeom != null)
                this.prstGeom.Write(sw, "prstGeom");
            if (this.noFill != null)
                sw.Write("<a:noFill/>");
            if (this.solidFill != null)
                this.solidFill.Write(sw, "solidFill");
            if (this.gradFill != null)
                this.gradFill.Write(sw, "gradFill");
            if (this.blipFill != null)
                this.blipFill.Write(sw, "blipFill");
            if (this.pattFill != null)
                this.pattFill.Write(sw, "pattFill");
            if (this.grpFill != null)
                sw.Write("<a:grpFill/>");
            if (this.ln != null)
                this.ln.Write(sw, "ln");
            if (this.effectLst != null)
                this.effectLst.Write(sw, "effectLst");
            if (this.effectDag != null)
                this.effectDag.Write(sw, "effectDag");
            if (this.scene3d != null)
                this.scene3d.Write(sw, "scene3d");
            if (this.sp3d != null)
                this.sp3d.Write(sw, "sp3d");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            sw.Write(string.Format("</xdr:{0}>", nodeName));
        }

    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_GraphicalObjectFrame
    {
        CT_GraphicalObjectFrameNonVisual nvGraphicFramePrField;
        CT_Transform2D xfrmField;
        CT_GraphicalObject graphicField;
        private string macroField;
        private bool fPublishedField;

        public void Set(CT_GraphicalObjectFrame obj)
        {
            this.xfrmField = obj.xfrmField;
            this.graphicField = obj.graphicField;
            this.nvGraphicFramePrField = obj.nvGraphicFramePrField;
            this.macroField = obj.macroField;
            this.fPublishedField = obj.fPublishedField;
        }

        public CT_Transform2D AddNewXfrm()
        {
            this.xfrmField = new CT_Transform2D();
            return this.xfrmField;
        }
        public CT_GraphicalObject AddNewGraphic()
        {
            this.graphicField = new CT_GraphicalObject();
            return this.graphicField;
        }

        public CT_GraphicalObjectFrameNonVisual AddNewNvGraphicFramePr()
        {
            this.nvGraphicFramePr = new CT_GraphicalObjectFrameNonVisual();
            return this.nvGraphicFramePr;
        }
        [XmlElement]
        public CT_GraphicalObjectFrameNonVisual nvGraphicFramePr
        {
            get { return nvGraphicFramePrField; }
            set { nvGraphicFramePrField = value; }
        }
        [XmlElement]
        public CT_Transform2D xfrm
        {
            get { return xfrmField; }
            set { xfrmField = value; }
        }
        [XmlAttribute]
        public string macro
        {
            get { return macroField; }
            set { macroField = value; }
        }
        [XmlAttribute]
        [DefaultValue(false)]
        public bool fPublished
        {
            get { return fPublishedField; }
            set { fPublishedField = value; }
        }
        [XmlElement]
        public CT_GraphicalObject graphic
        {
            get { return graphicField; }
            set { graphicField = value; }
        }
        public static CT_GraphicalObjectFrame Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_GraphicalObjectFrame ctObj = new CT_GraphicalObjectFrame();
            ctObj.macro = XmlHelper.ReadString(node.Attributes["macro"]);
            ctObj.fPublished = XmlHelper.ReadBool(node.Attributes["fPublished"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "nvGraphicFramePr")
                    ctObj.nvGraphicFramePr = CT_GraphicalObjectFrameNonVisual.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "xfrm")
                    ctObj.xfrm = CT_Transform2D.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "graphic")
                    ctObj.graphic = CT_GraphicalObject.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "macro", this.macro, true);
            XmlHelper.WriteAttribute(sw, "fPublished", this.fPublished, false);
            sw.Write(">");
            if (this.nvGraphicFramePr != null)
                this.nvGraphicFramePr.Write(sw, "nvGraphicFramePr");
            if (this.xfrm != null)
                this.xfrm.Write(sw, "xdr:xfrm");
            if (this.graphic != null)
                this.graphic.Write(sw, "graphic");
            sw.Write(string.Format("</xdr:{0}>", nodeName));
        }

    }



    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_ConnectorNonVisual
    {
        private CT_NonVisualDrawingProps cNvPrField;
        private CT_NonVisualConnectorProperties cNvCxnSpPrField;

        public CT_NonVisualConnectorProperties cNvCxnSpPr
        {
            get
            {
                return this.cNvCxnSpPrField;
            }
            set
            {
                this.cNvCxnSpPrField = value;
            }
        }
        public CT_NonVisualDrawingProps AddNewCNvPr()
        {
            this.cNvPr = new CT_NonVisualDrawingProps();
            return this.cNvPr;
        }
        public CT_NonVisualConnectorProperties AddNewCNvCxnSpPr()
        {
            this.cNvCxnSpPr = new CT_NonVisualConnectorProperties();
            return this.cNvCxnSpPr;
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
        public static CT_ConnectorNonVisual Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ConnectorNonVisual ctObj = new CT_ConnectorNonVisual();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "cNvCxnSpPr")
                    ctObj.cNvCxnSpPr = CT_NonVisualConnectorProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cNvPr")
                    ctObj.cNvPr = CT_NonVisualDrawingProps.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}", nodeName));
            sw.Write(">");
            if (this.cNvPr != null)
                this.cNvPr.Write(sw, "cNvPr");
            if (this.cNvCxnSpPr != null)
                this.cNvCxnSpPr.Write(sw, "cNvCxnSpPr");
            sw.Write(string.Format("</xdr:{0}>", nodeName));
        }


    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_Marker
    {
        private int _col;
        private long _colOff;
        private int _row;
        private long _rowOff;

        public int col
        {
            get 
            {
                return _col;
            }
            set 
            {
                _col = value;
            }
        }
        public long colOff
        {
            get 
            {
                return _colOff;
            }
            set
            {
                _colOff = value;
            }
        }
        public int row
        {
            get { return _row; }
            set { _row = value; }
        }
        public long rowOff
        {
            get
            {
                return _rowOff;
            }
            set
            {
                _rowOff = value;
            }
        }

        public static CT_Marker Parse(XmlNode node, XmlNamespaceManager nameSpaceManager)
        {
            CT_Marker ctMarker = new CT_Marker();
            foreach (XmlNode subnode in node.ChildNodes)
            {
                if (subnode.LocalName == "col")
                {
                    ctMarker.col = Int32.Parse(subnode.InnerText);
                }
                else if (subnode.LocalName == "colOff")
                {
                    ctMarker.colOff = Int64.Parse(subnode.InnerText);
                }
                else if (subnode.LocalName == "row")
                {
                    ctMarker.row = Int32.Parse(subnode.InnerText);
                }
                else if (subnode.LocalName == "rowOff")
                {
                    ctMarker.rowOff = Int64.Parse(subnode.InnerText);
                }
            }
            return ctMarker;
        }

        public override string ToString()
        {
            StringBuilder sb=new StringBuilder();
            using(StringWriter sw =new StringWriter(sb))
            {
                sw.Write("<xdr:col>");
                sw.Write(this.col.ToString());
                sw.Write("</xdr:col>");
                sw.Write("<xdr:colOff>");
                sw.Write(this.colOff.ToString());
                sw.Write("</xdr:colOff>");
                sw.Write("<xdr:row>");
                sw.Write(this.row.ToString());
                sw.Write("</xdr:row>");
                sw.Write("<xdr:rowOff>");
                sw.Write(this.rowOff.ToString());
                sw.Write("</xdr:rowOff>");
            }
            return sb.ToString();
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}>", nodeName));
            sw.Write("<xdr:col>");
            sw.Write(this.col.ToString());
            sw.Write("</xdr:col>");
            sw.Write("<xdr:colOff>");
            sw.Write(this.colOff.ToString());
            sw.Write("</xdr:colOff>");
            sw.Write("<xdr:row>");
            sw.Write(this.row.ToString());
            sw.Write("</xdr:row>");
            sw.Write("<xdr:rowOff>");
            sw.Write(this.rowOff.ToString());
            sw.Write("</xdr:rowOff>");
            sw.Write(string.Format("</{0}>", nodeName));
        }
    }
    public enum ST_EditAs
    {
        NONE,
        twoCell,
        oneCell,
        absolute
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_AnchorClientData
    {
        
        public static CT_AnchorClientData Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_AnchorClientData ctObj = new CT_AnchorClientData();
            ctObj.fLocksWithSheet = XmlHelper.ReadBool(node.Attributes["fLocksWithSheet"]);
            ctObj.fPrintsWithSheet = XmlHelper.ReadBool(node.Attributes["fPrintsWithSheet"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "fLocksWithSheet", this.fLocksWithSheet, false);
            XmlHelper.WriteAttribute(sw, "fPrintsWithSheet", this.fPrintsWithSheet,false);
            sw.Write(">");
            sw.Write(string.Format("</xdr:{0}>", nodeName));
        }
        bool _fLocksWithSheet;
        bool _fPrintsWithSheet;

        [XmlAttribute]
        public bool fLocksWithSheet
        {
            get
            {
                return _fLocksWithSheet;
            }
            set
            {
                _fLocksWithSheet = value;
            }
        }
        [XmlAttribute]
        public bool fPrintsWithSheet
        {
            get
            {
                return _fPrintsWithSheet;
            }
            set
            {
                _fPrintsWithSheet = value;
            }
        }
    }



    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    [XmlRoot("wsDr", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", IsNullable = true)]
    public class CT_Drawing
    {
        private List<IEG_Anchor> cellAnchors = new List<IEG_Anchor>();
        //private List<CT_AbsoulteCellAnchor> absoluteCellAnchors = new List<CT_AbsoulteCellAnchor>();

        public CT_TwoCellAnchor AddNewTwoCellAnchor()
        {
            CT_TwoCellAnchor anchor = new CT_TwoCellAnchor();
            cellAnchors.Add(anchor);
            return anchor;
        }
        public int SizeOfTwoCellAnchorArray()
        {
            int count = 0;
            foreach (IEG_Anchor anchor in cellAnchors)
            {
                if (anchor is CT_TwoCellAnchor)
                {
                    count++;
                }
            }
            return count;
        }

        public void Save(Stream stream)
        {
            using (StreamWriter sw = new StreamWriter(stream))
            {
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                sw.Write("<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">");
                foreach (IEG_Anchor anchor in this.cellAnchors)
                {
                    anchor.Write(sw);
                }
                sw.Write("</xdr:wsDr>");
            }
        }

        [XmlIgnore]
        public List<IEG_Anchor> CellAnchors
        {
            get { return cellAnchors; }
            set { cellAnchors = value; }
        }

        //[XmlElement("absoluteAnchor")]
        //public List<CT_TwoCellAnchor> AbsoluteAnchors
        //{
        //    get { return absoluteAnchors; }
        //    set { absoluteAnchors = value; }
        //}

        public void Set(CT_Drawing ctDrawing)
        {
            this.cellAnchors.Clear();
            foreach (IEG_Anchor anchor in ctDrawing.cellAnchors)
            {
                this.cellAnchors.Add(anchor);
            }
        }

        public int SizeOfAbsoluteAnchorArray()
        {
            return 0;
        }

        public int SizeOfOneCellAnchorArray()
        {
            int count = 0;
            foreach (IEG_Anchor anchor in cellAnchors)
            {
                if (anchor is CT_OneCellAnchor)
                {
                    count++;
                }
            }
            return count;
        }

        public static CT_Drawing Parse(XmlDocument xmldoc, XmlNamespaceManager namespaceManager)
        {
            XmlNodeList cellanchorNodes = xmldoc.SelectNodes("/xdr:wsDr/*", namespaceManager);
            CT_Drawing ctDrawing = new CT_Drawing();
            foreach (XmlNode node in cellanchorNodes)
            {
                if (node.LocalName == "twoCellAnchor")
                {
                    CT_TwoCellAnchor twoCellAnchor = CT_TwoCellAnchor.Parse(node, namespaceManager);
                    ctDrawing.cellAnchors.Add(twoCellAnchor);
                }
                else if (node.LocalName == "oneCellAnchor")
                {
                    CT_OneCellAnchor oneCellAnchor = CT_OneCellAnchor.Parse(node, namespaceManager);
                    ctDrawing.cellAnchors.Add(oneCellAnchor);
                }
                else if (node.LocalName == "absCellAnchor")
                {
                    CT_AbsoluteCellAnchor absCellAnchor = CT_AbsoluteCellAnchor.Parse(node, namespaceManager);
                    ctDrawing.cellAnchors.Add(absCellAnchor);
                }
            }
            return ctDrawing;
        }
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_LegacyDrawing
    {

        private string idField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id
        {
            get
            {
                return this.idField;
            }
            set
            {
                this.idField = value;
            }
        }
    }
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_OneCellAnchor: IEG_Anchor
    {
        private CT_Marker fromField = new CT_Marker();
        private CT_PositiveSize2D extField; //= new CT_PositiveSize2D();
        private CT_AnchorClientData clientDataField = new CT_AnchorClientData(); // 1..1 element
        private CT_Shape shapeField = null;
        private CT_GroupShape groupShapeField = null;
        private CT_GraphicalObjectFrame graphicalObjectField = null;
        private CT_Connector connectorField = null;
        private CT_Picture pictureField = null;

        [XmlElement]
        public CT_Marker from
        {
            get { return fromField; }
            set { fromField = value; }
        }

        [XmlElement]
        public CT_PositiveSize2D ext
        {
            get { return this.extField; }
            set { this.extField = value; }
        }
        public CT_AnchorClientData AddNewClientData()
        {
            this.clientDataField = new CT_AnchorClientData();
            return this.clientDataField;
        }
        [XmlElement]
        public CT_AnchorClientData clientData
        {
            get { return clientDataField; }
            set { clientDataField = value; }
        }
        public CT_Shape sp
        {
            get { return shapeField; }
            set { shapeField = value; }
        }
        public CT_GroupShape groupShape
        {
            get { return groupShapeField; }
            set { groupShapeField = value; }
        }
        public CT_GraphicalObjectFrame graphicFrame
        {
            get { return graphicalObjectField; }
            set { graphicalObjectField = value; }
        }
        public CT_Connector connector
        {
            get { return connectorField; }
            set { connectorField = value; }
        }
        public CT_Picture picture
        {
            get { return pictureField; }
            set { pictureField = value; }
        }
         
        public void Write(StreamWriter sw)
        {
            sw.Write("<xdr:oneCellAnchor>");
            this.from.Write(sw, "xdr:from");
            this.ext.Write(sw, "xdr:ext");
            if (this.sp != null)
                sp.Write(sw, "sp");
            else if (this.connector != null)
                this.connector.Write(sw, "cxnSp");
            else if (this.groupShape != null)
                this.groupShape.Write(sw, "grpSp");
            else if (this.graphicalObjectField != null)
                this.graphicalObjectField.Write(sw, "graphicFrame");
            else if (this.pictureField != null)
                this.picture.Write(sw, "pic");
            if (this.clientData != null)
                this.clientData.Write(sw, "clientData");
            sw.Write("</xdr:oneCellAnchor>");
        }

        internal static CT_OneCellAnchor Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_OneCellAnchor oneCellAnchor = new CT_OneCellAnchor();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "from")
                {
                    oneCellAnchor.from = CT_Marker.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "ext")
                {
                    oneCellAnchor.ext = CT_PositiveSize2D.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "sp")
                {
                    oneCellAnchor.sp = CT_Shape.Parse(childNode, namespaceManager); ;
                }
                else if (childNode.LocalName == "pic")
                {
                    oneCellAnchor.picture = CT_Picture.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "cxnSp")
                {
                    oneCellAnchor.connector = CT_Connector.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "grpSp")
                {
                    oneCellAnchor.groupShape = CT_GroupShape.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "graphicFrame")
                {
                    oneCellAnchor.graphicFrame = CT_GraphicalObjectFrame.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "clientData")
                {
                    oneCellAnchor.clientData = CT_AnchorClientData.Parse(childNode, namespaceManager);
                }
            }
            return oneCellAnchor;
        }
    }

    public interface IEG_Anchor
    {
        CT_Shape sp { get; set; }
        CT_Connector connector { get; set; }
        CT_GraphicalObjectFrame graphicFrame { get; set; }
        CT_Picture picture { get; set; }
        CT_GroupShape groupShape { get; set; }
        CT_AnchorClientData clientData { get; set; }
        void Write(StreamWriter sw);
    }
    public class CT_AbsoluteCellAnchor : IEG_Anchor
    {
        CT_Point2D posField;
        CT_PositiveSize2D extField;
        CT_AnchorClientData clientDataField = new CT_AnchorClientData();
        private CT_Shape shapeField = null;
        private CT_GroupShape groupShapeField = null;
        private CT_GraphicalObjectFrame graphicalObjectField = null;
        private CT_Connector connectorField = null;
        private CT_Picture pictureField = null;
        public static CT_AbsoluteCellAnchor Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_AbsoluteCellAnchor absCellAnchor = new CT_AbsoluteCellAnchor();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "pos")
                {
                    absCellAnchor.pos = CT_Point2D.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "ext")
                {
                    absCellAnchor.ext = CT_PositiveSize2D.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "sp")
                {
                    absCellAnchor.sp = CT_Shape.Parse(childNode, namespaceManager); ;
                }
                else if (childNode.LocalName == "pic")
                {
                    absCellAnchor.picture = CT_Picture.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "cxnSp")
                {
                    absCellAnchor.connector = CT_Connector.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "grpSp")
                {
                    absCellAnchor.groupShape = CT_GroupShape.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "graphicFrame")
                {
                    absCellAnchor.graphicFrame = CT_GraphicalObjectFrame.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "clientData")
                {
                    absCellAnchor.clientData = CT_AnchorClientData.Parse(childNode, namespaceManager);
                }
            }
            return absCellAnchor;
        }

        public CT_AnchorClientData clientData
        {
            get { return clientDataField; }
            set { clientDataField = value; }
        }

        public CT_Point2D AddNewOff()
        {
            this.posField = new CT_Point2D();
            return this.posField;
        }

        public CT_Point2D pos
        {
            get
            {
                return this.posField;
            }
            set
            {
                this.posField = value;
            }
        }
        public CT_PositiveSize2D ext
        {
            get { return this.extField; }
            set { this.extField = value; }
        }
        public CT_AnchorClientData AddNewClientData()
        {
            this.clientDataField = new CT_AnchorClientData();
            return this.clientDataField;
        }
        public CT_Shape sp
        {
            get { return shapeField; }
            set { shapeField = value; }
        }
        public CT_GroupShape groupShape
        {
            get { return groupShapeField; }
            set { groupShapeField = value; }
        }
        public CT_GraphicalObjectFrame graphicFrame
        {
            get { return graphicalObjectField; }
            set { graphicalObjectField = value; }
        }
        public CT_Connector connector
        {
            get { return connectorField; }
            set { connectorField = value; }
        }
        public CT_Picture picture
        {
            get { return pictureField; }
            set { pictureField = value; }
        }
        public void Write(StreamWriter sw)
        {
            sw.Write("<xdr:absCellAnchor>");
            if (this.pos!=null)
                this.pos.Write(sw, "pos");
            if (this.sp != null)
                sp.Write(sw, "sp");
            else if (this.connector != null)
                this.connector.Write(sw, "cxnSp");
            else if (this.groupShape != null)
                this.groupShape.Write(sw, "grpSp");
            else if (this.graphicalObjectField != null)
                this.graphicalObjectField.Write(sw, "graphicFrame");
            else if (this.pictureField != null)
                this.picture.Write(sw, "pic");
            sw.Write("</xdr:oneCellAnchor>");

            if (this.clientData != null)
            {
                this.clientData.Write(sw, "clientData");
            }
            sw.Write("</xdr:absCellAnchor");
        }
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_TwoCellAnchor : IEG_Anchor //- was empty interface
    {
        private CT_Marker fromField = new CT_Marker(); // 1..1 element
        private CT_Marker toField = new CT_Marker(); // 1..1 element
        // 1..1 element choice - one of CT_Shape, CT_GroupShape, CT_GraphicalObjectFrame, CT_Connector or CT_Picture
        private CT_Shape shapeField = null;
        private CT_GroupShape groupShapeField = null;
        private CT_GraphicalObjectFrame graphicalObjectField = null;
        private CT_Connector connectorField = null;
        private CT_Picture pictureField = null;

        private CT_AnchorClientData clientDataField = null; // 1..1 element

        private ST_EditAs editAsField = ST_EditAs.NONE; // 0..1 attribute

        public CT_TwoCellAnchor()
        {
            this.clientDataField = new CT_AnchorClientData();
        }

        public CT_Shape AddNewSp()
        {
            shapeField = new CT_Shape();
            return shapeField;
        }

        public CT_GroupShape AddNewGrpSp()
        {
            groupShapeField = new CT_GroupShape();
            return groupShapeField;
        }

        public CT_GraphicalObjectFrame AddNewGraphicFrame()
        {
            graphicalObjectField = new CT_GraphicalObjectFrame();
            return graphicalObjectField;
        }

        public CT_Connector AddNewCxnSp()
        {
            connectorField = new CT_Connector();
            return connectorField;
        }

        public CT_Picture AddNewPic()
        {
            pictureField = new CT_Picture();
            return pictureField;
        }

        public CT_AnchorClientData AddNewClientData()
        {
            this.clientDataField = new CT_AnchorClientData();
            return this.clientDataField;
        }
        [XmlElement]
        public CT_AnchorClientData clientData
        {
            get { return clientDataField; }
            set { clientDataField = value; }
        }
        [XmlAttribute]
        public ST_EditAs editAs
        {
            get { return editAsField; }
            set { editAsField = value; }
        }
        bool editAsSpecifiedField = false;
        [XmlIgnore]
        public bool editAsSpecified
        {
            get { return editAsSpecifiedField; }
            set { editAsSpecifiedField = value; }
        }

        [XmlElement]
        public CT_Marker from
        {
            get { return fromField; }
            set { fromField = value; }
        }

        [XmlElement]
        public CT_Marker to
        {
            get { return toField; }
            set { toField = value; }
        }

        #region Choice - one of CT_Shape, CT_GroupShape, CT_GraphicalObjectFrame, CT_Connector or CT_Picture

        [XmlElement]
        public CT_Shape sp
        {
            get { return shapeField; }
            set { shapeField = value; }
        }
        [XmlElement]
        public CT_GroupShape groupShape
        {
            get { return groupShapeField; }
            set { groupShapeField = value; }
        }

        [XmlElement]
        public CT_GraphicalObjectFrame graphicFrame
        {
            get { return graphicalObjectField; }
            set { graphicalObjectField = value; }
        }

        [XmlElement]
        public CT_Connector connector
        {
            get { return connectorField; }
            set { connectorField = value; }
        }

        [XmlElement("pic")]
        public CT_Picture picture
        {
            get { return pictureField; }
            set { pictureField = value; }
        }

        #endregion Choice - one of CT_Shape, CT_GroupShape, CT_GraphicalObjectFrame, CT_Connector or CT_Picture

        public void Write(StreamWriter sw)
        {
            sw.Write("<xdr:twoCellAnchor");
            if(this.editAsField!= ST_EditAs.NONE)
                sw.Write(string.Format(" editAs=\"{0}\"",this.editAsField.ToString()));
            sw.Write(">");
            this.from.Write(sw, "xdr:from");
            this.to.Write(sw, "xdr:to");
            if (this.sp != null)
                sp.Write(sw, "sp");
            else if (this.connector != null)
                this.connector.Write(sw, "cxnSp");
            else if (this.groupShape != null)
                this.groupShape.Write(sw, "grpSp");
            else if (this.graphicalObjectField != null)
                this.graphicalObjectField.Write(sw, "graphicFrame");
            else if (this.pictureField != null)
                this.picture.Write(sw, "pic");
            if (this.clientData != null)
            {
                this.clientData.Write(sw, "clientData");
            }
            sw.Write("</xdr:twoCellAnchor>");
        }

        internal static CT_TwoCellAnchor Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_TwoCellAnchor twoCellAnchor = new CT_TwoCellAnchor();
            if (node.Attributes["editAs"] != null)
                twoCellAnchor.editAs = (ST_EditAs)Enum.Parse(typeof(ST_EditAs), node.Attributes["editAs"].Value);

            foreach(XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "from")
                {
                    twoCellAnchor.from = CT_Marker.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "to")
                {
                    twoCellAnchor.to = CT_Marker.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "sp")
                {
                    twoCellAnchor.sp = CT_Shape.Parse(childNode, namespaceManager); ;
                }
                else if (childNode.LocalName == "pic")
                {
                    twoCellAnchor.picture = CT_Picture.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "cxnSp")
                {
                    twoCellAnchor.connector = CT_Connector.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "grpSp")
                {
                    twoCellAnchor.groupShape = CT_GroupShape.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "graphicFrame")
                {
                    twoCellAnchor.graphicFrame = CT_GraphicalObjectFrame.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "clientData")
                {
                    twoCellAnchor.clientData = CT_AnchorClientData.Parse(childNode, namespaceManager);
                }
            }
            return twoCellAnchor;
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_Connector // empty interface: EG_ObjectChoices
    {
        string macroField;
        bool fPublishedField;
        private CT_ShapeProperties spPrField;
        private CT_ShapeStyle styleField;
        private CT_ConnectorNonVisual nvCxnSpPrField;
        public CT_ConnectorNonVisual nvCxnSpPr
        {
            get { return nvCxnSpPrField; }
            set { nvCxnSpPrField = value; }
        }
        public static CT_Connector Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Connector ctObj = new CT_Connector();
            ctObj.macro = XmlHelper.ReadString(node.Attributes["macro"]);
            ctObj.fPublished = XmlHelper.ReadBool(node.Attributes["fPublished"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "nvCxnSpPr")
                    ctObj.nvCxnSpPr = CT_ConnectorNonVisual.Parse(childNode, namespaceManager);
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
            XmlHelper.WriteAttribute(sw, "macro", this.macro, true);
            XmlHelper.WriteAttribute(sw, "fPublished", this.fPublished,false);
            sw.Write(">");
            if (this.nvCxnSpPr != null)
                this.nvCxnSpPr.Write(sw, "nvCxnSpPr");
            if (this.spPr != null)
                this.spPr.Write(sw, "spPr");
            if (this.style != null)
                this.style.Write(sw, "style");
            sw.Write(string.Format("</xdr:{0}>", nodeName));
        }

        public void Set(CT_Connector obj)
        {
            this.macroField = obj.macro;
            this.fPublishedField = obj.fPublished;
            this.spPrField = obj.spPr;
            this.styleField = obj.style;
            this.nvCxnSpPrField = obj.nvCxnSpPr;
        }
        public CT_ConnectorNonVisual AddNewNvCxnSpPr()
        {
            this.nvCxnSpPr = new CT_ConnectorNonVisual();
            return nvCxnSpPr;
        }
        public CT_ShapeProperties AddNewSpPr()
        {
            this.spPrField = new CT_ShapeProperties();
            return spPrField;
        }
        public CT_ShapeStyle AddNewStyle()
        {
            this.styleField = new CT_ShapeStyle();
            return this.styleField;
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
            get { return this.macroField; }
            set { this.macroField = value; }
        }
        [XmlAttribute]
        public bool fPublished
        {
            get { return this.fPublishedField; }
            set { this.fPublishedField = value; }
        }
    }

    

}
