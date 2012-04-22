using System;
using System.ComponentModel;
namespace NPOI.OpenXmlFormats.Dml {
    
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/picture")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/picture", IsNullable=true)]
    public partial class CT_PictureNonVisual {
        
        private CT_NonVisualDrawingProps cNvPrField;
        
        private CT_NonVisualPictureProperties cNvPicPrField;
        
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


        /// <remarks/>
        public CT_NonVisualDrawingProps cNvPr {
            get {
                return this.cNvPrField;
            }
            set {
                this.cNvPrField = value;
            }
        }
        
        /// <remarks/>
        public CT_NonVisualPictureProperties cNvPicPr {
            get {
                return this.cNvPicPrField;
            }
            set {
                this.cNvPicPrField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_NonVisualDrawingProps {
        
        private CT_Hyperlink hlinkClickField;
        
        private CT_Hyperlink hlinkHoverField;
        
        private CT_OfficeArtExtensionList extLstField;
        
        private uint idField;
        
        private string nameField;
        
        private string descrField;
        
        private bool hiddenField;
        
        public CT_NonVisualDrawingProps() {
            this.descrField = "";
            this.hiddenField = false;
        }
        
        /// <remarks/>
        public CT_Hyperlink hlinkClick {
            get {
                return this.hlinkClickField;
            }
            set {
                this.hlinkClickField = value;
            }
        }
        
        /// <remarks/>
        public CT_Hyperlink hlinkHover {
            get {
                return this.hlinkHoverField;
            }
            set {
                this.hlinkHoverField = value;
            }
        }
        
        /// <remarks/>
        public CT_OfficeArtExtensionList extLst {
            get {
                return this.extLstField;
            }
            set {
                this.extLstField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public uint id {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string name {
            get {
                return this.nameField;
            }
            set {
                this.nameField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue("")]
        public string descr {
            get {
                return this.descrField;
            }
            set {
                this.descrField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(false)]
        public bool hidden {
            get {
                return this.hiddenField;
            }
            set {
                this.hiddenField = value;
            }
        }
    }
    
    
       
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_PictureLocking {
        
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
        
        public CT_PictureLocking() {
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
        
        /// <remarks/>
        public CT_OfficeArtExtensionList extLst {
            get {
                return this.extLstField;
            }
            set {
                this.extLstField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(false)]
        public bool noGrp {
            get {
                return this.noGrpField;
            }
            set {
                this.noGrpField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(false)]
        public bool noSelect {
            get {
                return this.noSelectField;
            }
            set {
                this.noSelectField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(false)]
        public bool noRot {
            get {
                return this.noRotField;
            }
            set {
                this.noRotField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(false)]
        public bool noChangeAspect {
            get {
                return this.noChangeAspectField;
            }
            set {
                this.noChangeAspectField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(false)]
        public bool noMove {
            get {
                return this.noMoveField;
            }
            set {
                this.noMoveField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(false)]
        public bool noResize {
            get {
                return this.noResizeField;
            }
            set {
                this.noResizeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(false)]
        public bool noEditPoints {
            get {
                return this.noEditPointsField;
            }
            set {
                this.noEditPointsField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(false)]
        public bool noAdjustHandles {
            get {
                return this.noAdjustHandlesField;
            }
            set {
                this.noAdjustHandlesField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(false)]
        public bool noChangeArrowheads {
            get {
                return this.noChangeArrowheadsField;
            }
            set {
                this.noChangeArrowheadsField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(false)]
        public bool noChangeShapeType {
            get {
                return this.noChangeShapeTypeField;
            }
            set {
                this.noChangeShapeTypeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(false)]
        public bool noCrop {
            get {
                return this.noCropField;
            }
            set {
                this.noCropField = value;
            }
        }

        public void SetNoChangeAspect(bool p)
        {
            throw new NotImplementedException();
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_NonVisualPictureProperties {
        
        private CT_PictureLocking picLocksField;
        
        private CT_OfficeArtExtensionList extLstField;
        
        private bool preferRelativeResizeField;
        
        public CT_NonVisualPictureProperties() {
            this.preferRelativeResizeField = true;
        }

        public CT_PictureLocking AddNewPicLocks()
        {
            this.picLocksField = new CT_PictureLocking();
            return picLocksField;
        }

        /// <remarks/>
        public CT_PictureLocking picLocks {
            get {
                return this.picLocksField;
            }
            set {
                this.picLocksField = value;
            }
        }
        
        /// <remarks/>
        public CT_OfficeArtExtensionList extLst {
            get {
                return this.extLstField;
            }
            set {
                this.extLstField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(true)]
        public bool preferRelativeResize {
            get {
                return this.preferRelativeResizeField;
            }
            set {
                this.preferRelativeResizeField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/picture")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/picture", IsNullable=true)]
    public partial class CT_Picture {
        
        private CT_PictureNonVisual nvPicPrField;
        
        private CT_BlipFillProperties blipFillField;
        
        private CT_ShapeProperties spPrField;
        
        /// <remarks/>
        public CT_PictureNonVisual nvPicPr {
            get {
                return this.nvPicPrField;
            }
            set {
                this.nvPicPrField = value;
            }
        }
        
        /// <remarks/>
        public CT_BlipFillProperties blipFill {
            get {
                return this.blipFillField;
            }
            set {
                this.blipFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_ShapeProperties spPr {
            get {
                return this.spPrField;
            }
            set {
                this.spPrField = value;
            }
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
        public void Set(CT_Picture obj)
        {
            throw new NotImplementedException();
        }

    }
}
