namespace NPOI.OpenXmlFormats.Dml {
    
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GvmlUseShapeRectangle {
    }
    
    
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_GroupLocking {
        
        private CT_OfficeArtExtensionList extLstField;
        
        private bool noGrpField;
        
        private bool noUngrpField;
        
        private bool noSelectField;
        
        private bool noRotField;
        
        private bool noChangeAspectField;
        
        private bool noMoveField;
        
        private bool noResizeField;
        
        public CT_GroupLocking() {
            this.noGrpField = false;
            this.noUngrpField = false;
            this.noSelectField = false;
            this.noRotField = false;
            this.noChangeAspectField = false;
            this.noMoveField = false;
            this.noResizeField = false;
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
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noGrp {
            get {
                return this.noGrpField;
            }
            set {
                this.noGrpField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noUngrp {
            get {
                return this.noUngrpField;
            }
            set {
                this.noUngrpField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noSelect {
            get {
                return this.noSelectField;
            }
            set {
                this.noSelectField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noRot {
            get {
                return this.noRotField;
            }
            set {
                this.noRotField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noChangeAspect {
            get {
                return this.noChangeAspectField;
            }
            set {
                this.noChangeAspectField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noMove {
            get {
                return this.noMoveField;
            }
            set {
                this.noMoveField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noResize {
            get {
                return this.noResizeField;
            }
            set {
                this.noResizeField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_NonVisualGroupDrawingShapeProps {
        
        private CT_GroupLocking grpSpLocksField;
        
        private CT_OfficeArtExtensionList extLstField;
        
        /// <remarks/>
        public CT_GroupLocking grpSpLocks {
            get {
                return this.grpSpLocksField;
            }
            set {
                this.grpSpLocksField = value;
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
    }
   
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_GraphicalObjectFrameLocking {
        
        private CT_OfficeArtExtensionList extLstField;
        
        private bool noGrpField;
        
        private bool noDrilldownField;
        
        private bool noSelectField;
        
        private bool noChangeAspectField;
        
        private bool noMoveField;
        
        private bool noResizeField;
        
        public CT_GraphicalObjectFrameLocking() {
            this.noGrpField = false;
            this.noDrilldownField = false;
            this.noSelectField = false;
            this.noChangeAspectField = false;
            this.noMoveField = false;
            this.noResizeField = false;
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
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noGrp {
            get {
                return this.noGrpField;
            }
            set {
                this.noGrpField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noDrilldown {
            get {
                return this.noDrilldownField;
            }
            set {
                this.noDrilldownField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noSelect {
            get {
                return this.noSelectField;
            }
            set {
                this.noSelectField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noChangeAspect {
            get {
                return this.noChangeAspectField;
            }
            set {
                this.noChangeAspectField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noMove {
            get {
                return this.noMoveField;
            }
            set {
                this.noMoveField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noResize {
            get {
                return this.noResizeField;
            }
            set {
                this.noResizeField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_NonVisualGraphicFrameProperties {
        
        private CT_GraphicalObjectFrameLocking graphicFrameLocksField;
        
        private CT_OfficeArtExtensionList extLstField;
        
        /// <remarks/>
        public CT_GraphicalObjectFrameLocking graphicFrameLocks {
            get {
                return this.graphicFrameLocksField;
            }
            set {
                this.graphicFrameLocksField = value;
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
    }
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_ConnectorLocking {
        
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
        
        public CT_ConnectorLocking() {
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
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noGrp {
            get {
                return this.noGrpField;
            }
            set {
                this.noGrpField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noSelect {
            get {
                return this.noSelectField;
            }
            set {
                this.noSelectField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noRot {
            get {
                return this.noRotField;
            }
            set {
                this.noRotField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noChangeAspect {
            get {
                return this.noChangeAspectField;
            }
            set {
                this.noChangeAspectField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noMove {
            get {
                return this.noMoveField;
            }
            set {
                this.noMoveField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noResize {
            get {
                return this.noResizeField;
            }
            set {
                this.noResizeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noEditPoints {
            get {
                return this.noEditPointsField;
            }
            set {
                this.noEditPointsField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noAdjustHandles {
            get {
                return this.noAdjustHandlesField;
            }
            set {
                this.noAdjustHandlesField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noChangeArrowheads {
            get {
                return this.noChangeArrowheadsField;
            }
            set {
                this.noChangeArrowheadsField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noChangeShapeType {
            get {
                return this.noChangeShapeTypeField;
            }
            set {
                this.noChangeShapeTypeField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_NonVisualConnectorProperties {
        
        private CT_ConnectorLocking cxnSpLocksField;
        
        private CT_Connection stCxnField;
        
        private CT_Connection endCxnField;
        
        private CT_OfficeArtExtensionList extLstField;
        
        /// <remarks/>
        public CT_ConnectorLocking cxnSpLocks {
            get {
                return this.cxnSpLocksField;
            }
            set {
                this.cxnSpLocksField = value;
            }
        }
        
        /// <remarks/>
        public CT_Connection stCxn {
            get {
                return this.stCxnField;
            }
            set {
                this.stCxnField = value;
            }
        }
        
        /// <remarks/>
        public CT_Connection endCxn {
            get {
                return this.endCxnField;
            }
            set {
                this.endCxnField = value;
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
    }
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IncludeInSchema=false)]
    public enum ItemsChoiceType2 {
        
        /// <remarks/>
        arcTo,
        
        /// <remarks/>
        close,
        
        /// <remarks/>
        cubicBezTo,
        
        /// <remarks/>
        lnTo,
        
        /// <remarks/>
        moveTo,
        
        /// <remarks/>
        quadBezTo,
    }
    
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_ShapeLocking {
        
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
        
        public CT_ShapeLocking() {
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
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noGrp {
            get {
                return this.noGrpField;
            }
            set {
                this.noGrpField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noSelect {
            get {
                return this.noSelectField;
            }
            set {
                this.noSelectField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noRot {
            get {
                return this.noRotField;
            }
            set {
                this.noRotField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noChangeAspect {
            get {
                return this.noChangeAspectField;
            }
            set {
                this.noChangeAspectField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noMove {
            get {
                return this.noMoveField;
            }
            set {
                this.noMoveField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noResize {
            get {
                return this.noResizeField;
            }
            set {
                this.noResizeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noEditPoints {
            get {
                return this.noEditPointsField;
            }
            set {
                this.noEditPointsField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noAdjustHandles {
            get {
                return this.noAdjustHandlesField;
            }
            set {
                this.noAdjustHandlesField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noChangeArrowheads {
            get {
                return this.noChangeArrowheadsField;
            }
            set {
                this.noChangeArrowheadsField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noChangeShapeType {
            get {
                return this.noChangeShapeTypeField;
            }
            set {
                this.noChangeShapeTypeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool noTextEdit {
            get {
                return this.noTextEditField;
            }
            set {
                this.noTextEditField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_NonVisualDrawingShapeProps {
        
        private CT_ShapeLocking spLocksField;
        
        private CT_OfficeArtExtensionList extLstField;
        
        private bool txBoxField;
        
        public CT_NonVisualDrawingShapeProps() {
            this.txBoxField = false;
        }
        
        /// <remarks/>
        public CT_ShapeLocking spLocks {
            get {
                return this.spLocksField;
            }
            set {
                this.spLocksField = value;
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
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool txBox {
            get {
                return this.txBoxField;
            }
            set {
                this.txBoxField = value;
            }
        }
    }
    
      
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GvmlTextShape {
        
        private CT_TextBody txBodyField;
        
        private object itemField;
        
        private CT_OfficeArtExtensionList extLstField;
        
        /// <remarks/>
        public CT_TextBody txBody {
            get {
                return this.txBodyField;
            }
            set {
                this.txBodyField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("useSpRect", typeof(CT_GvmlUseShapeRectangle))]
        [System.Xml.Serialization.XmlElementAttribute("xfrm", typeof(CT_Transform2D))]
        public object Item {
            get {
                return this.itemField;
            }
            set {
                this.itemField = value;
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
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GvmlShapeNonVisual {
        
        private CT_NonVisualDrawingProps cNvPrField;
        
        private CT_NonVisualDrawingShapeProps cNvSpPrField;
        
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
        public CT_NonVisualDrawingShapeProps cNvSpPr {
            get {
                return this.cNvSpPrField;
            }
            set {
                this.cNvSpPrField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GvmlShape {
        
        private CT_GvmlShapeNonVisual nvSpPrField;
        
        private CT_ShapeProperties spPrField;
        
        private CT_GvmlTextShape txSpField;
        
        private CT_ShapeStyle styleField;
        
        private CT_OfficeArtExtensionList extLstField;
        
        /// <remarks/>
        public CT_GvmlShapeNonVisual nvSpPr {
            get {
                return this.nvSpPrField;
            }
            set {
                this.nvSpPrField = value;
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
        
        /// <remarks/>
        public CT_GvmlTextShape txSp {
            get {
                return this.txSpField;
            }
            set {
                this.txSpField = value;
            }
        }
        
        /// <remarks/>
        public CT_ShapeStyle style {
            get {
                return this.styleField;
            }
            set {
                this.styleField = value;
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
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GvmlConnectorNonVisual {
        
        private CT_NonVisualDrawingProps cNvPrField;
        
        private CT_NonVisualConnectorProperties cNvCxnSpPrField;
        
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
        public CT_NonVisualConnectorProperties cNvCxnSpPr {
            get {
                return this.cNvCxnSpPrField;
            }
            set {
                this.cNvCxnSpPrField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GvmlConnector {
        
        private CT_GvmlConnectorNonVisual nvCxnSpPrField;
        
        private CT_ShapeProperties spPrField;
        
        private CT_ShapeStyle styleField;
        
        private CT_OfficeArtExtensionList extLstField;
        
        /// <remarks/>
        public CT_GvmlConnectorNonVisual nvCxnSpPr {
            get {
                return this.nvCxnSpPrField;
            }
            set {
                this.nvCxnSpPrField = value;
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
        
        /// <remarks/>
        public CT_ShapeStyle style {
            get {
                return this.styleField;
            }
            set {
                this.styleField = value;
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
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GvmlPictureNonVisual {
        
        private CT_NonVisualDrawingProps cNvPrField;
        
        private CT_NonVisualPictureProperties cNvPicPrField;
        
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
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GvmlPicture {
        
        private CT_GvmlPictureNonVisual nvPicPrField;
        
        private CT_BlipFillProperties blipFillField;
        
        private CT_ShapeProperties spPrField;
        
        private CT_ShapeStyle styleField;
        
        private CT_OfficeArtExtensionList extLstField;
        
        /// <remarks/>
        public CT_GvmlPictureNonVisual nvPicPr {
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
        
        /// <remarks/>
        public CT_ShapeStyle style {
            get {
                return this.styleField;
            }
            set {
                this.styleField = value;
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
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GvmlGraphicFrameNonVisual {
        
        private CT_NonVisualDrawingProps cNvPrField;
        
        private CT_NonVisualGraphicFrameProperties cNvGraphicFramePrField;
        
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
        public CT_NonVisualGraphicFrameProperties cNvGraphicFramePr {
            get {
                return this.cNvGraphicFramePrField;
            }
            set {
                this.cNvGraphicFramePrField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GvmlGraphicalObjectFrame {
        
        private CT_GvmlGraphicFrameNonVisual nvGraphicFramePrField;
        
        private CT_GraphicalObject graphicField;
        
        private CT_Transform2D xfrmField;
        
        private CT_OfficeArtExtensionList extLstField;
        
        /// <remarks/>
        public CT_GvmlGraphicFrameNonVisual nvGraphicFramePr {
            get {
                return this.nvGraphicFramePrField;
            }
            set {
                this.nvGraphicFramePrField = value;
            }
        }
        
        /// <remarks/>
        public CT_GraphicalObject graphic {
            get {
                return this.graphicField;
            }
            set {
                this.graphicField = value;
            }
        }
        
        /// <remarks/>
        public CT_Transform2D xfrm {
            get {
                return this.xfrmField;
            }
            set {
                this.xfrmField = value;
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
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GvmlGroupShapeNonVisual {
        
        private CT_NonVisualDrawingProps cNvPrField;
        
        private CT_NonVisualGroupDrawingShapeProps cNvGrpSpPrField;
        
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
        public CT_NonVisualGroupDrawingShapeProps cNvGrpSpPr {
            get {
                return this.cNvGrpSpPrField;
            }
            set {
                this.cNvGrpSpPrField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GvmlGroupShape {
        
        private CT_GvmlGroupShapeNonVisual nvGrpSpPrField;
        
        private CT_GroupShapeProperties grpSpPrField;
        
        private object[] itemsField;
        
        private CT_OfficeArtExtensionList extLstField;
        
        /// <remarks/>
        public CT_GvmlGroupShapeNonVisual nvGrpSpPr {
            get {
                return this.nvGrpSpPrField;
            }
            set {
                this.nvGrpSpPrField = value;
            }
        }
        
        /// <remarks/>
        public CT_GroupShapeProperties grpSpPr {
            get {
                return this.grpSpPrField;
            }
            set {
                this.grpSpPrField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("cxnSp", typeof(CT_GvmlConnector))]
        [System.Xml.Serialization.XmlElementAttribute("graphicFrame", typeof(CT_GvmlGraphicalObjectFrame))]
        [System.Xml.Serialization.XmlElementAttribute("grpSp", typeof(CT_GvmlGroupShape))]
        [System.Xml.Serialization.XmlElementAttribute("pic", typeof(CT_GvmlPicture))]
        [System.Xml.Serialization.XmlElementAttribute("sp", typeof(CT_GvmlShape))]
        [System.Xml.Serialization.XmlElementAttribute("txSp", typeof(CT_GvmlTextShape))]
        public object[] Items {
            get {
                return this.itemsField;
            }
            set {
                this.itemsField = value;
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
    }
}
