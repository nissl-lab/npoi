using NPOI.OpenXml4Net.Util;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Dml
{
    [Serializable]
    
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_ConnectorLocking
    {

        private CT_OfficeArtExtensionList extLstField;

        private bool noGrpField;

        private bool noSelectField;

        private bool noRotField;

        private bool noChangeAspectField;

        private bool noMoveField;

        private bool noResizeField;

        private bool noEditPointsField;

        private bool noAdjustHandlesField;

        private bool noChangeArrowheadsField;

        private bool noChangeShapeTypeField;
        public static CT_ConnectorLocking Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ConnectorLocking ctObj = new CT_ConnectorLocking();
            ctObj.noGrp = XmlHelper.ReadBool(node.Attributes["noGrp"]);
            ctObj.noSelect = XmlHelper.ReadBool(node.Attributes["noSelect"]);
            ctObj.noRot = XmlHelper.ReadBool(node.Attributes["noRot"]);
            ctObj.noChangeAspect = XmlHelper.ReadBool(node.Attributes["noChangeAspect"]);
            ctObj.noMove = XmlHelper.ReadBool(node.Attributes["noMove"]);
            ctObj.noResize = XmlHelper.ReadBool(node.Attributes["noResize"]);
            ctObj.noEditPoints = XmlHelper.ReadBool(node.Attributes["noEditPoints"]);
            ctObj.noAdjustHandles = XmlHelper.ReadBool(node.Attributes["noAdjustHandles"]);
            ctObj.noChangeArrowheads = XmlHelper.ReadBool(node.Attributes["noChangeArrowheads"]);
            ctObj.noChangeShapeType = XmlHelper.ReadBool(node.Attributes["noChangeShapeType"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_OfficeArtExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "noGrp", this.noGrp);
            XmlHelper.WriteAttribute(sw, "noSelect", this.noSelect);
            XmlHelper.WriteAttribute(sw, "noRot", this.noRot);
            XmlHelper.WriteAttribute(sw, "noChangeAspect", this.noChangeAspect);
            XmlHelper.WriteAttribute(sw, "noMove", this.noMove);
            XmlHelper.WriteAttribute(sw, "noResize", this.noResize);
            XmlHelper.WriteAttribute(sw, "noEditPoints", this.noEditPoints);
            XmlHelper.WriteAttribute(sw, "noAdjustHandles", this.noAdjustHandles);
            XmlHelper.WriteAttribute(sw, "noChangeArrowheads", this.noChangeArrowheads);
            XmlHelper.WriteAttribute(sw, "noChangeShapeType", this.noChangeShapeType);
            sw.Write(">");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            sw.Write(string.Format("</a:{0}>", nodeName));
        }

        public CT_ConnectorLocking()
        {
            //this.extLstField = new CT_OfficeArtExtensionList();
            this.noGrpField = false;
            this.noSelectField = false;
            this.noRotField = false;
            this.noChangeAspectField = false;
            this.noMoveField = false;
            this.noResizeField = false;
            this.noEditPointsField = false;
            this.noAdjustHandlesField = false;
            this.noChangeArrowheadsField = false;
            this.noChangeShapeTypeField = false;
        }

        [XmlElement(Order = 0)]
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
        public bool noGrp
        {
            get
            {
                return this.noGrpField;
            }
            set
            {
                this.noGrpField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noSelect
        {
            get
            {
                return this.noSelectField;
            }
            set
            {
                this.noSelectField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noRot
        {
            get
            {
                return this.noRotField;
            }
            set
            {
                this.noRotField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noChangeAspect
        {
            get
            {
                return this.noChangeAspectField;
            }
            set
            {
                this.noChangeAspectField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noMove
        {
            get
            {
                return this.noMoveField;
            }
            set
            {
                this.noMoveField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noResize
        {
            get
            {
                return this.noResizeField;
            }
            set
            {
                this.noResizeField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noEditPoints
        {
            get
            {
                return this.noEditPointsField;
            }
            set
            {
                this.noEditPointsField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noAdjustHandles
        {
            get
            {
                return this.noAdjustHandlesField;
            }
            set
            {
                this.noAdjustHandlesField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noChangeArrowheads
        {
            get
            {
                return this.noChangeArrowheadsField;
            }
            set
            {
                this.noChangeArrowheadsField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noChangeShapeType
        {
            get
            {
                return this.noChangeShapeTypeField;
            }
            set
            {
                this.noChangeShapeTypeField = value;
            }
        }
    }



    [Serializable]
    
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_PictureLocking
    {

        private CT_OfficeArtExtensionList extLstField;

        private bool noGrpField;

        private bool noSelectField;

        private bool noRotField;

        private bool noChangeAspectField;

        private bool noMoveField;

        private bool noResizeField;

        private bool noEditPointsField;

        private bool noAdjustHandlesField;

        private bool noChangeArrowheadsField;

        private bool noChangeShapeTypeField;

        private bool noCropField;

        public CT_PictureLocking()
        {
            //this.extLstField = new CT_OfficeArtExtensionList();
            this.noGrpField = false;
            this.noSelectField = false;
            this.noRotField = false;
            this.noChangeAspectField = false;
            this.noMoveField = false;
            this.noResizeField = false;
            this.noEditPointsField = false;
            this.noAdjustHandlesField = false;
            this.noChangeArrowheadsField = false;
            this.noChangeShapeTypeField = false;
            this.noCropField = false;
        }
        public static CT_PictureLocking Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PictureLocking ctObj = new CT_PictureLocking();
            ctObj.noGrp = XmlHelper.ReadBool(node.Attributes["noGrp"]);
            ctObj.noSelect = XmlHelper.ReadBool(node.Attributes["noSelect"]);
            ctObj.noRot = XmlHelper.ReadBool(node.Attributes["noRot"]);
            ctObj.noChangeAspect = XmlHelper.ReadBool(node.Attributes["noChangeAspect"]);
            ctObj.noMove = XmlHelper.ReadBool(node.Attributes["noMove"]);
            ctObj.noResize = XmlHelper.ReadBool(node.Attributes["noResize"]);
            ctObj.noEditPoints = XmlHelper.ReadBool(node.Attributes["noEditPoints"]);
            ctObj.noAdjustHandles = XmlHelper.ReadBool(node.Attributes["noAdjustHandles"]);
            ctObj.noChangeArrowheads = XmlHelper.ReadBool(node.Attributes["noChangeArrowheads"]);
            ctObj.noChangeShapeType = XmlHelper.ReadBool(node.Attributes["noChangeShapeType"]);
            ctObj.noCrop = XmlHelper.ReadBool(node.Attributes["noCrop"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_OfficeArtExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            if(noGrp)
                XmlHelper.WriteAttribute(sw, "noGrp", this.noGrp);
            if (noSelect)
                XmlHelper.WriteAttribute(sw, "noSelect", this.noSelect);
            if (noRot)
                XmlHelper.WriteAttribute(sw, "noRot", this.noRot);
            if (noChangeAspect)
                XmlHelper.WriteAttribute(sw, "noChangeAspect", this.noChangeAspect);
            if (noMove)
                XmlHelper.WriteAttribute(sw, "noMove", this.noMove);
            if (noResize)
                XmlHelper.WriteAttribute(sw, "noResize", this.noResize);
            if (noEditPoints)
                XmlHelper.WriteAttribute(sw, "noEditPoints", this.noEditPoints);
            if (noAdjustHandles)
                XmlHelper.WriteAttribute(sw, "noAdjustHandles", this.noAdjustHandles);
            if (noChangeArrowheads)
                XmlHelper.WriteAttribute(sw, "noChangeArrowheads", this.noChangeArrowheads);
            if (noChangeShapeType)
                XmlHelper.WriteAttribute(sw, "noChangeShapeType", this.noChangeShapeType);
            if (noCrop)
                XmlHelper.WriteAttribute(sw, "noCrop", this.noCrop);
            sw.Write(">");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            sw.Write(string.Format("</a:{0}>", nodeName));
        }
        [XmlElement(Order = 0)]
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
        public bool noGrp
        {
            get
            {
                return this.noGrpField;
            }
            set
            {
                this.noGrpField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noSelect
        {
            get
            {
                return this.noSelectField;
            }
            set
            {
                this.noSelectField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noRot
        {
            get
            {
                return this.noRotField;
            }
            set
            {
                this.noRotField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noChangeAspect
        {
            get
            {
                return this.noChangeAspectField;
            }
            set
            {
                this.noChangeAspectField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noMove
        {
            get
            {
                return this.noMoveField;
            }
            set
            {
                this.noMoveField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noResize
        {
            get
            {
                return this.noResizeField;
            }
            set
            {
                this.noResizeField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noEditPoints
        {
            get
            {
                return this.noEditPointsField;
            }
            set
            {
                this.noEditPointsField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noAdjustHandles
        {
            get
            {
                return this.noAdjustHandlesField;
            }
            set
            {
                this.noAdjustHandlesField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noChangeArrowheads
        {
            get
            {
                return this.noChangeArrowheadsField;
            }
            set
            {
                this.noChangeArrowheadsField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noChangeShapeType
        {
            get
            {
                return this.noChangeShapeTypeField;
            }
            set
            {
                this.noChangeShapeTypeField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noCrop
        {
            get
            {
                return this.noCropField;
            }
            set
            {
                this.noCropField = value;
            }
        }
    }





    [Serializable]
    
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_GraphicalObjectFrameLocking
    {

        private CT_OfficeArtExtensionList extLstField;

        private bool noGrpField;

        private bool noDrilldownField;

        private bool noSelectField;

        private bool noChangeAspectField;

        private bool noMoveField;

        private bool noResizeField;

        public CT_GraphicalObjectFrameLocking()
        {
            //this.extLstField = new CT_OfficeArtExtensionList();
            this.noGrpField = false;
            this.noDrilldownField = false;
            this.noSelectField = false;
            this.noChangeAspectField = false;
            this.noMoveField = false;
            this.noResizeField = false;
        }
        public static CT_GraphicalObjectFrameLocking Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_GraphicalObjectFrameLocking ctObj = new CT_GraphicalObjectFrameLocking();
            ctObj.noGrp = XmlHelper.ReadBool(node.Attributes["noGrp"]);
            ctObj.noDrilldown = XmlHelper.ReadBool(node.Attributes["noDrilldown"]);
            ctObj.noSelect = XmlHelper.ReadBool(node.Attributes["noSelect"]);
            ctObj.noChangeAspect = XmlHelper.ReadBool(node.Attributes["noChangeAspect"]);
            ctObj.noMove = XmlHelper.ReadBool(node.Attributes["noMove"]);
            ctObj.noResize = XmlHelper.ReadBool(node.Attributes["noResize"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_OfficeArtExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "noGrp", this.noGrp, false);
            XmlHelper.WriteAttribute(sw, "noDrilldown", this.noDrilldown, false);
            XmlHelper.WriteAttribute(sw, "noSelect", this.noSelect, false);
            XmlHelper.WriteAttribute(sw, "noChangeAspect", this.noChangeAspect, false);
            XmlHelper.WriteAttribute(sw, "noMove", this.noMove, false);
            XmlHelper.WriteAttribute(sw, "noResize", this.noResize, false);
            sw.Write(">");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            sw.Write(string.Format("</a:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
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
        public bool noGrp
        {
            get
            {
                return this.noGrpField;
            }
            set
            {
                this.noGrpField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noDrilldown
        {
            get
            {
                return this.noDrilldownField;
            }
            set
            {
                this.noDrilldownField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noSelect
        {
            get
            {
                return this.noSelectField;
            }
            set
            {
                this.noSelectField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noChangeAspect
        {
            get
            {
                return this.noChangeAspectField;
            }
            set
            {
                this.noChangeAspectField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noMove
        {
            get
            {
                return this.noMoveField;
            }
            set
            {
                this.noMoveField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noResize
        {
            get
            {
                return this.noResizeField;
            }
            set
            {
                this.noResizeField = value;
            }
        }
    }


    [Serializable]
    
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
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
            sw.Write(string.Format("<pic:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "id", this.id, true);
            XmlHelper.WriteAttribute(sw, "name", this.name);
            XmlHelper.WriteAttribute(sw, "descr", this.descr);
            if(this.hidden)
                XmlHelper.WriteAttribute(sw, "hidden", this.hidden);
            sw.Write(">");
            if (this.hlinkClick != null)
                this.hlinkClick.Write(sw, "hlinkClick");
            if (this.hlinkHover != null)
                this.hlinkHover.Write(sw, "hlinkHover");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            sw.Write(string.Format("</pic:{0}>", nodeName));
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
            sw.Write(string.Format("<a:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "txBox", this.txBox, false);
            sw.Write(">");
            if (this.spLocks != null)
                this.spLocks.Write(sw, "spLocks");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            sw.Write(string.Format("</a:{0}>", nodeName));
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
    
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_NonVisualConnectorProperties
    {
        public static CT_NonVisualConnectorProperties Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_NonVisualConnectorProperties ctObj = new CT_NonVisualConnectorProperties();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "cxnSpLocks")
                    ctObj.cxnSpLocks = CT_ConnectorLocking.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "stCxn")
                    ctObj.stCxn = CT_Connection.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "endCxn")
                    ctObj.endCxn = CT_Connection.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_OfficeArtExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}", nodeName));
            sw.Write(">");
            if (this.cxnSpLocks != null)
                this.cxnSpLocks.Write(sw, "cxnSpLocks");
            if (this.stCxn != null)
                this.stCxn.Write(sw, "stCxn");
            if (this.endCxn != null)
                this.endCxn.Write(sw, "endCxn");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            sw.Write(string.Format("</xdr:{0}>", nodeName));
        }


        private CT_ConnectorLocking cxnSpLocksField;

        private CT_Connection stCxnField;

        private CT_Connection endCxnField;

        private CT_OfficeArtExtensionList extLstField;

        public CT_NonVisualConnectorProperties()
        {
        }

        [XmlElement(Order = 0)]
        public CT_ConnectorLocking cxnSpLocks
        {
            get
            {
                return this.cxnSpLocksField;
            }
            set
            {
                this.cxnSpLocksField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_Connection stCxn
        {
            get
            {
                return this.stCxnField;
            }
            set
            {
                this.stCxnField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_Connection endCxn
        {
            get
            {
                return this.endCxnField;
            }
            set
            {
                this.endCxnField = value;
            }
        }

        [XmlElement(Order = 3)]
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
    
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
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
            sw.Write(string.Format("<pic:{0}", nodeName));
            if (!preferRelativeResize)
                XmlHelper.WriteAttribute(sw, "preferRelativeResize", this.preferRelativeResize);
            sw.Write(">");
            if (this.picLocks != null)  
                this.picLocks.Write(sw, "picLocks");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            sw.Write(string.Format("</pic:{0}>", nodeName));
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
                return preferRelativeResizeField;
            }
            set
            {
                this.preferRelativeResizeField = value;
            }
        }
    }


    [Serializable]
    
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_NonVisualGroupDrawingShapeProps
    {

        private CT_GroupLocking grpSpLocksField;

        private CT_OfficeArtExtensionList extLstField;

        public static CT_NonVisualGroupDrawingShapeProps Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_NonVisualGroupDrawingShapeProps ctObj = new CT_NonVisualGroupDrawingShapeProps();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "grpSpLocks")
                    ctObj.grpSpLocks = CT_GroupLocking.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_OfficeArtExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            sw.Write(">");
            if (this.grpSpLocks != null)
                this.grpSpLocks.Write(sw, "grpSpLocks");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            sw.Write(string.Format("</a:{0}>", nodeName));
        }

        public CT_NonVisualGroupDrawingShapeProps()
        {
            //this.extLstField = new CT_OfficeArtExtensionList();
           //this.grpSpLocksField = new CT_GroupLocking();
        }

        [XmlElement(Order = 0)]
        public CT_GroupLocking grpSpLocks
        {
            get
            {
                return this.grpSpLocksField;
            }
            set
            {
                this.grpSpLocksField = value;
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
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
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
            sw.Write(string.Format("<a:{0}", nodeName));
            sw.Write(">");
            if (this.graphicFrameLocks != null)
                this.graphicFrameLocks.Write(sw, "graphicFrameLocks");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            sw.Write(string.Format("</a:{0}>", nodeName));
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
    
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_ShapeLocking
    {

        private CT_OfficeArtExtensionList extLstField;

        private bool noGrpField;

        private bool noSelectField;

        private bool noRotField;

        private bool noChangeAspectField;

        private bool noMoveField;

        private bool noResizeField;

        private bool noEditPointsField;

        private bool noAdjustHandlesField;

        private bool noChangeArrowheadsField;

        private bool noChangeShapeTypeField;

        private bool noTextEditField;
        public static CT_ShapeLocking Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ShapeLocking ctObj = new CT_ShapeLocking();
            ctObj.noGrp = XmlHelper.ReadBool(node.Attributes["noGrp"]);
            ctObj.noSelect = XmlHelper.ReadBool(node.Attributes["noSelect"]);
            ctObj.noRot = XmlHelper.ReadBool(node.Attributes["noRot"]);
            ctObj.noChangeAspect = XmlHelper.ReadBool(node.Attributes["noChangeAspect"]);
            ctObj.noMove = XmlHelper.ReadBool(node.Attributes["noMove"]);
            ctObj.noResize = XmlHelper.ReadBool(node.Attributes["noResize"]);
            ctObj.noEditPoints = XmlHelper.ReadBool(node.Attributes["noEditPoints"]);
            ctObj.noAdjustHandles = XmlHelper.ReadBool(node.Attributes["noAdjustHandles"]);
            ctObj.noChangeArrowheads = XmlHelper.ReadBool(node.Attributes["noChangeArrowheads"]);
            ctObj.noChangeShapeType = XmlHelper.ReadBool(node.Attributes["noChangeShapeType"]);
            ctObj.noTextEdit = XmlHelper.ReadBool(node.Attributes["noTextEdit"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_OfficeArtExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "noGrp", this.noGrp, false);
            XmlHelper.WriteAttribute(sw, "noSelect", this.noSelect, false);
            XmlHelper.WriteAttribute(sw, "noRot", this.noRot, false);
            XmlHelper.WriteAttribute(sw, "noChangeAspect", this.noChangeAspect, false);
            XmlHelper.WriteAttribute(sw, "noMove", this.noMove, false);
            XmlHelper.WriteAttribute(sw, "noResize", this.noResize, false);
            XmlHelper.WriteAttribute(sw, "noEditPoints", this.noEditPoints, false);
            XmlHelper.WriteAttribute(sw, "noAdjustHandles", this.noAdjustHandles, false);
            XmlHelper.WriteAttribute(sw, "noChangeArrowheads", this.noChangeArrowheads, false);
            XmlHelper.WriteAttribute(sw, "noChangeShapeType", this.noChangeShapeType, false);
            XmlHelper.WriteAttribute(sw, "noTextEdit", this.noTextEdit, false);
            sw.Write(">");
            if(this.extLst!=null)
                this.extLst.Write(sw, "extLst");
            sw.Write(string.Format("</a:{0}>", nodeName));
        }
        public CT_ShapeLocking()
        {
            //this.extLstField = new CT_OfficeArtExtensionList();
            this.noGrpField = false;
            this.noSelectField = false;
            this.noRotField = false;
            this.noChangeAspectField = false;
            this.noMoveField = false;
            this.noResizeField = false;
            this.noEditPointsField = false;
            this.noAdjustHandlesField = false;
            this.noChangeArrowheadsField = false;
            this.noChangeShapeTypeField = false;
            this.noTextEditField = false;
        }

        [XmlElement(Order = 0)]
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
        public bool noGrp
        {
            get
            {
                return this.noGrpField;
            }
            set
            {
                this.noGrpField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noSelect
        {
            get
            {
                return this.noSelectField;
            }
            set
            {
                this.noSelectField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noRot
        {
            get
            {
                return this.noRotField;
            }
            set
            {
                this.noRotField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noChangeAspect
        {
            get
            {
                return this.noChangeAspectField;
            }
            set
            {
                this.noChangeAspectField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noMove
        {
            get
            {
                return this.noMoveField;
            }
            set
            {
                this.noMoveField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noResize
        {
            get
            {
                return this.noResizeField;
            }
            set
            {
                this.noResizeField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noEditPoints
        {
            get
            {
                return this.noEditPointsField;
            }
            set
            {
                this.noEditPointsField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noAdjustHandles
        {
            get
            {
                return this.noAdjustHandlesField;
            }
            set
            {
                this.noAdjustHandlesField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noChangeArrowheads
        {
            get
            {
                return this.noChangeArrowheadsField;
            }
            set
            {
                this.noChangeArrowheadsField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noChangeShapeType
        {
            get
            {
                return this.noChangeShapeTypeField;
            }
            set
            {
                this.noChangeShapeTypeField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noTextEdit
        {
            get
            {
                return this.noTextEditField;
            }
            set
            {
                this.noTextEditField = value;
            }
        }
    }

    [Serializable]
    
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_GroupLocking
    {

        private CT_OfficeArtExtensionList extLstField;

        private bool noGrpField;

        private bool noUngrpField;

        private bool noSelectField;

        private bool noRotField;

        private bool noChangeAspectField;

        private bool noMoveField;

        private bool noResizeField;

        public CT_GroupLocking()
        {
            //this.extLstField = new CT_OfficeArtExtensionList();
            this.noGrpField = false;
            this.noUngrpField = false;
            this.noSelectField = false;
            this.noRotField = false;
            this.noChangeAspectField = false;
            this.noMoveField = false;
            this.noResizeField = false;
        }
        public static CT_GroupLocking Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_GroupLocking ctObj = new CT_GroupLocking();
            ctObj.noGrp = XmlHelper.ReadBool(node.Attributes["noGrp"]);
            ctObj.noUngrp = XmlHelper.ReadBool(node.Attributes["noUngrp"]);
            ctObj.noSelect = XmlHelper.ReadBool(node.Attributes["noSelect"]);
            ctObj.noRot = XmlHelper.ReadBool(node.Attributes["noRot"]);
            ctObj.noChangeAspect = XmlHelper.ReadBool(node.Attributes["noChangeAspect"]);
            ctObj.noMove = XmlHelper.ReadBool(node.Attributes["noMove"]);
            ctObj.noResize = XmlHelper.ReadBool(node.Attributes["noResize"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_OfficeArtExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "noGrp", this.noGrp);
            XmlHelper.WriteAttribute(sw, "noUngrp", this.noUngrp);
            XmlHelper.WriteAttribute(sw, "noSelect", this.noSelect);
            XmlHelper.WriteAttribute(sw, "noRot", this.noRot);
            XmlHelper.WriteAttribute(sw, "noChangeAspect", this.noChangeAspect);
            XmlHelper.WriteAttribute(sw, "noMove", this.noMove);
            XmlHelper.WriteAttribute(sw, "noResize", this.noResize);
            sw.Write(">");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            sw.Write(string.Format("</a:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
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
        public bool noGrp
        {
            get
            {
                return this.noGrpField;
            }
            set
            {
                this.noGrpField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noUngrp
        {
            get
            {
                return this.noUngrpField;
            }
            set
            {
                this.noUngrpField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noSelect
        {
            get
            {
                return this.noSelectField;
            }
            set
            {
                this.noSelectField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noRot
        {
            get
            {
                return this.noRotField;
            }
            set
            {
                this.noRotField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noChangeAspect
        {
            get
            {
                return this.noChangeAspectField;
            }
            set
            {
                this.noChangeAspectField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noMove
        {
            get
            {
                return this.noMoveField;
            }
            set
            {
                this.noMoveField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool noResize
        {
            get
            {
                return this.noResizeField;
            }
            set
            {
                this.noResizeField = value;
            }
        }
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
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
            sw.Write(string.Format("<{0}", nodeName));
            if(this.bwMode!= ST_BlackWhiteMode.none)
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
                this.blipFill.Write(sw, "a:blipFill");
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
            sw.Write(string.Format("</{0}>", nodeName));
        }

    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_GroupShapeProperties
    {

        private CT_GroupTransform2D xfrmField;

        private CT_NoFillProperties noFillField;

        private CT_SolidColorFillProperties solidFillField;

        private CT_GradientFillProperties gradFillField;

        private CT_BlipFillProperties blipFillField;

        private CT_PatternFillProperties pattFillField;

        private CT_GroupFillProperties grpFillField;

        private CT_EffectList effectLstField;

        private CT_EffectContainer effectDagField;

        private CT_Scene3D scene3dField;

        private CT_OfficeArtExtensionList extLstField;

        private ST_BlackWhiteMode bwModeField;

        private bool bwModeFieldSpecified;
        public static CT_GroupShapeProperties Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_GroupShapeProperties ctObj = new CT_GroupShapeProperties();
            if (node.Attributes["bwMode"] != null)
                ctObj.bwMode = (ST_BlackWhiteMode)Enum.Parse(typeof(ST_BlackWhiteMode), node.Attributes["bwMode"].Value);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "xfrm")
                    ctObj.xfrm = CT_GroupTransform2D.Parse(childNode, namespaceManager);
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
                else if (childNode.LocalName == "effectLst")
                    ctObj.effectLst = CT_EffectList.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "effectDag")
                    ctObj.effectDag = CT_EffectContainer.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "scene3d")
                    ctObj.scene3d = CT_Scene3D.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_OfficeArtExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            if(this.bwMode!= ST_BlackWhiteMode.none)
                XmlHelper.WriteAttribute(sw, "bwMode", this.bwMode.ToString());
            sw.Write(">");
            if (this.xfrm != null)
                this.xfrm.Write(sw, "xfrm");
            if (this.noFill != null)
                sw.Write("<a:noFill/>");
            if (this.solidFill != null)
                this.solidFill.Write(sw, "solidFill");
            if (this.gradFill != null)
                this.gradFill.Write(sw, "gradFill");
            if (this.blipFill != null)
                this.blipFill.Write(sw, "a:blipFill");
            if (this.pattFill != null)
                this.pattFill.Write(sw, "pattFill");
            if (this.grpFill != null)
                sw.Write("<a:grpFill/>");
            if (this.effectLst != null)
                this.effectLst.Write(sw, "effectLst");
            if (this.effectDag != null)
                this.effectDag.Write(sw, "effectDag");
            if (this.scene3d != null)
                this.scene3d.Write(sw, "scene3d");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            sw.Write(string.Format("</a:{0}>", nodeName));
        }

        public CT_GroupShapeProperties()
        {
            //this.extlstfield = new ct_officeartextensionlist();
            //this.scene3dfield = new ct_scene3d();
            //this.effectdagfield = new ct_effectcontainer();
            //this.effectlstfield = new ct_effectlist();
            //this.grpfillfield = new ct_groupfillproperties();
            //this.pattfillfield = new ct_patternfillproperties();
            //this.blipfillfield = new ct_blipfillproperties();
            //this.gradfillfield = new ct_gradientfillproperties();
            //this.solidfillfield = new ct_solidcolorfillproperties();
            //this.nofillfield = new ct_nofillproperties();
            //this.xfrmField = new CT_GroupTransform2D();
        }

        public CT_GroupTransform2D AddNewXfrm()
        {
            this.xfrmField = new CT_GroupTransform2D();
            return this.xfrmField;
        }
        [XmlElement(Order = 0)]
        public CT_GroupTransform2D xfrm
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

        [XmlElement(Order = 2)]
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

        [XmlElement(Order = 3)]
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

        [XmlElement(Order = 4)]
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

        [XmlElement(Order = 5)]
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

        [XmlElement(Order = 6)]
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

        [XmlElement(Order = 7)]
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

        [XmlElement(Order = 8)]
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

        [XmlElement(Order = 9)]
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

        [XmlElement(Order = 10)]
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
            get
            {
                return this.bwModeFieldSpecified;
            }
            set
            {
                this.bwModeFieldSpecified = value;
            }
        }
    }
}
