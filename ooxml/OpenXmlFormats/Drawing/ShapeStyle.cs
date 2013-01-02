
using System;
using System.Xml.Serialization;


namespace NPOI.OpenXmlFormats.Dml
{



    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_StyleMatrixReference
    {

        private CT_ScRgbColor scrgbClrField;

        private CT_SRgbColor srgbClrField;

        private CT_HslColor hslClrField;

        private CT_SystemColor sysClrField;

        private CT_SchemeColor schemeClrField;

        private CT_PresetColor prstClrField;

        private uint idxField;

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
    public partial class CT_FontReference
    {

        private CT_ScRgbColor scrgbClrField;

        private CT_SRgbColor srgbClrField;

        private CT_HslColor hslClrField;

        private CT_SystemColor sysClrField;

        private CT_SchemeColor schemeClrField;

        private CT_PresetColor prstClrField;

        private ST_FontCollectionIndex idxField;

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
    public partial class CT_ShapeStyle
    {

        private CT_StyleMatrixReference lnRefField;

        private CT_StyleMatrixReference fillRefField;

        private CT_StyleMatrixReference effectRefField;

        private CT_FontReference fontRefField;

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
