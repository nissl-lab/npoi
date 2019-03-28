
using NPOI.OpenXml4Net.Util;
using System;
using System.IO;
using System.Xml;
using System.Xml.Serialization;


namespace NPOI.OpenXmlFormats.Dml
{



    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_StyleMatrixReference
    {

        private CT_ScRgbColor scrgbClrField;

        private CT_SRgbColor srgbClrField;

        private CT_HslColor hslClrField;

        private CT_SystemColor sysClrField;

        private CT_SchemeColor schemeClrField;

        private CT_PresetColor prstClrField;

        private uint idxField;
        public static CT_StyleMatrixReference Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_StyleMatrixReference ctObj = new CT_StyleMatrixReference();
            ctObj.idx = XmlHelper.ReadUInt(node.Attributes["idx"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "scrgbClr")
                    ctObj.scrgbClr = CT_ScRgbColor.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "srgbClr")
                    ctObj.srgbClr = CT_SRgbColor.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "hslClr")
                    ctObj.hslClr = CT_HslColor.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sysClr")
                    ctObj.sysClr = CT_SystemColor.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "schemeClr")
                    ctObj.schemeClr = CT_SchemeColor.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "prstClr")
                    ctObj.prstClr = CT_PresetColor.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "idx", this.idx, true);
            sw.Write(">");
            if (this.scrgbClr != null)
                this.scrgbClr.Write(sw, "scrgbClr");
            if (this.srgbClr != null)
                this.srgbClr.Write(sw, "srgbClr");
            if (this.hslClr != null)
                this.hslClr.Write(sw, "hslClr");
            if (this.sysClr != null)
                this.sysClr.Write(sw, "sysClr");
            if (this.schemeClr != null)
                this.schemeClr.Write(sw, "schemeClr");
            if (this.prstClr != null)
                this.prstClr.Write(sw, "prstClr");
            sw.Write(string.Format("</a:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_ScRgbColor scrgbClr
        {
            get
            {
                return this.scrgbClrField;
            }
            set
            {
                this.scrgbClrField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_SRgbColor srgbClr
        {
            get
            {
                return this.srgbClrField;
            }
            set
            {
                this.srgbClrField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_HslColor hslClr
        {
            get
            {
                return this.hslClrField;
            }
            set
            {
                this.hslClrField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_SystemColor sysClr
        {
            get
            {
                return this.sysClrField;
            }
            set
            {
                this.sysClrField = value;
            }
        }
        public CT_SchemeColor AddNewSchemeClr()
        {
            this.schemeClrField = new CT_SchemeColor();
            return this.schemeClrField;
        }
        [XmlElement(Order = 4)]
        public CT_SchemeColor schemeClr
        {
            get
            {
                return this.schemeClrField;
            }
            set
            {
                this.schemeClrField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_PresetColor prstClr
        {
            get
            {
                return this.prstClrField;
            }
            set
            {
                this.prstClrField = value;
            }
        }


        [XmlAttribute]
        public uint idx
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
    }




    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_FontReference
    {

        private CT_ScRgbColor scrgbClrField;

        private CT_SRgbColor srgbClrField;

        private CT_HslColor hslClrField;

        private CT_SystemColor sysClrField;

        private CT_SchemeColor schemeClrField;

        private CT_PresetColor prstClrField;

        private ST_FontCollectionIndex idxField;
        public static CT_FontReference Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_FontReference ctObj = new CT_FontReference();
            if (node.Attributes["idx"] != null)
                ctObj.idx = (ST_FontCollectionIndex)Enum.Parse(typeof(ST_FontCollectionIndex), node.Attributes["idx"].Value);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "scrgbClr")
                    ctObj.scrgbClr = CT_ScRgbColor.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "srgbClr")
                    ctObj.srgbClr = CT_SRgbColor.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "hslClr")
                    ctObj.hslClr = CT_HslColor.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sysClr")
                    ctObj.sysClr = CT_SystemColor.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "schemeClr")
                    ctObj.schemeClr = CT_SchemeColor.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "prstClr")
                    ctObj.prstClr = CT_PresetColor.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "idx", this.idx.ToString());
            sw.Write(">");
            if (this.scrgbClr != null)
                this.scrgbClr.Write(sw, "scrgbClr");
            if (this.srgbClr != null)
                this.srgbClr.Write(sw, "srgbClr");
            if (this.hslClr != null)
                this.hslClr.Write(sw, "hslClr");
            if (this.sysClr != null)
                this.sysClr.Write(sw, "sysClr");
            if (this.schemeClr != null)
                this.schemeClr.Write(sw, "schemeClr");
            if (this.prstClr != null)
                this.prstClr.Write(sw, "prstClr");
            sw.Write(string.Format("</a:{0}>", nodeName));
        }
        [XmlElement(Order = 0)]
        public CT_ScRgbColor scrgbClr
        {
            get
            {
                return this.scrgbClrField;
            }
            set
            {
                this.scrgbClrField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_SRgbColor srgbClr
        {
            get
            {
                return this.srgbClrField;
            }
            set
            {
                this.srgbClrField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_HslColor hslClr
        {
            get
            {
                return this.hslClrField;
            }
            set
            {
                this.hslClrField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_SystemColor sysClr
        {
            get
            {
                return this.sysClrField;
            }
            set
            {
                this.sysClrField = value;
            }
        }
        public CT_SchemeColor AddNewSchemeClr()
        {
            this.schemeClrField = new CT_SchemeColor();
            return this.schemeClrField;
        }
        [XmlElement(Order = 4)]
        public CT_SchemeColor schemeClr
        {
            get
            {
                return this.schemeClrField;
            }
            set
            {
                this.schemeClrField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_PresetColor prstClr
        {
            get
            {
                return this.prstClrField;
            }
            set
            {
                this.prstClrField = value;
            }
        }


        [XmlAttribute]
        public ST_FontCollectionIndex idx
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
    }




    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
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
            sw.Write(string.Format("<a:{0}", nodeName));
            sw.Write(">");
            if (this.lnRef != null)
                this.lnRef.Write(sw, "lnRef");
            if (this.fillRef != null)
                this.fillRef.Write(sw, "fillRef");
            if (this.effectRef != null)
                this.effectRef.Write(sw, "effectRef");
            if (this.fontRef != null)
                this.fontRef.Write(sw, "fontRef");
            sw.Write(string.Format("</a:{0}>", nodeName));
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
}
