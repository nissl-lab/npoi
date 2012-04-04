namespace NPOI.OpenXmlFormats.Dml {
    
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_StyleMatrixReference {
        
        private CT_ScRgbColor scrgbClrField;
        
        private CT_SRgbColor srgbClrField;
        
        private CT_HslColor hslClrField;
        
        private CT_SystemColor sysClrField;
        
        private CT_SchemeColor schemeClrField;
        
        private CT_PresetColor prstClrField;
        
        private uint idxField;
        
        /// <remarks/>
        public CT_ScRgbColor scrgbClr {
            get {
                return this.scrgbClrField;
            }
            set {
                this.scrgbClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SRgbColor srgbClr {
            get {
                return this.srgbClrField;
            }
            set {
                this.srgbClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_HslColor hslClr {
            get {
                return this.hslClrField;
            }
            set {
                this.hslClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SystemColor sysClr {
            get {
                return this.sysClrField;
            }
            set {
                this.sysClrField = value;
            }
        }
        public CT_SchemeColor AddNewSchemeClr()
        {
            this.schemeClrField = new CT_SchemeColor();
            return this.schemeClrField;
        }
        /// <remarks/>
        public CT_SchemeColor schemeClr {
            get {
                return this.schemeClrField;
            }
            set {
                this.schemeClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_PresetColor prstClr {
            get {
                return this.prstClrField;
            }
            set {
                this.prstClrField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public uint idx {
            get {
                return this.idxField;
            }
            set {
                this.idxField = value;
            }
        }
    }
    
    
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_FontReference {
        
        private CT_ScRgbColor scrgbClrField;
        
        private CT_SRgbColor srgbClrField;
        
        private CT_HslColor hslClrField;
        
        private CT_SystemColor sysClrField;
        
        private CT_SchemeColor schemeClrField;
        
        private CT_PresetColor prstClrField;
        
        private ST_FontCollectionIndex idxField;
        
        /// <remarks/>
        public CT_ScRgbColor scrgbClr {
            get {
                return this.scrgbClrField;
            }
            set {
                this.scrgbClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SRgbColor srgbClr {
            get {
                return this.srgbClrField;
            }
            set {
                this.srgbClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_HslColor hslClr {
            get {
                return this.hslClrField;
            }
            set {
                this.hslClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SystemColor sysClr {
            get {
                return this.sysClrField;
            }
            set {
                this.sysClrField = value;
            }
        }
        public CT_SchemeColor AddNewSchemeClr()
        {
            this.schemeClrField=new CT_SchemeColor();
            return this.schemeClrField;
        }
        /// <remarks/>
        public CT_SchemeColor schemeClr {
            get {
                return this.schemeClrField;
            }
            set {
                this.schemeClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_PresetColor prstClr {
            get {
                return this.prstClrField;
            }
            set {
                this.prstClrField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_FontCollectionIndex idx {
            get {
                return this.idxField;
            }
            set {
                this.idxField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_FontCollectionIndex {
        
        /// <remarks/>
        major,
        
        /// <remarks/>
        minor,
        
        /// <remarks/>
        none,
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_ShapeStyle {
        
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
            this.fontRefField=new CT_FontReference();
            return this.fontRefField;
        }
        public CT_StyleMatrixReference AddNewEffectRef()
        {
            this.effectRefField = new CT_StyleMatrixReference();
            return this.effectRefField;
        }
        /// <remarks/>
        public CT_StyleMatrixReference lnRef {
            get {
                return this.lnRefField;
            }
            set {
                this.lnRefField = value;
            }
        }
        
        /// <remarks/>
        public CT_StyleMatrixReference fillRef {
            get {
                return this.fillRefField;
            }
            set {
                this.fillRefField = value;
            }
        }
        
        /// <remarks/>
        public CT_StyleMatrixReference effectRef {
            get {
                return this.effectRefField;
            }
            set {
                this.effectRefField = value;
            }
        }
        
        /// <remarks/>
        public CT_FontReference fontRef {
            get {
                return this.fontRefField;
            }
            set {
                this.fontRefField = value;
            }
        }
    }
}
