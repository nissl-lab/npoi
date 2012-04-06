using System.Xml.Serialization;
using System.ComponentModel;
namespace NPOI.OpenXmlFormats.Vml
{
    
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_ShapeDefaults {
        
        private CT_Fill fillField;
        
        private CT_Stroke strokeField;
        
        private CT_Textbox textboxField;
        
        private CT_Shadow shadowField;
        
        private CT_Skew skewField;
        
        private CT_Extrusion extrusionField;
        
        private CT_Callout calloutField;
        
        private CT_Lock lockField;
        
        private CT_ColorMru colormruField;
        
        private CT_ColorMenu colormenuField;
        
        private ST_Ext extField;
        
        private bool extFieldSpecified;
        
        private string spidmaxField;
        
        private string styleField;
        
        private ST_TrueFalse1 fill1Field;
        
        private bool fill1FieldSpecified;
        
        private string fillcolorField;
        
        private ST_TrueFalse1 stroke1Field;
        
        private bool stroke1FieldSpecified;
        
        private string strokecolorField;
        
        private ST_TrueFalse1 allowincellField;
        
        private bool allowincellFieldSpecified;
        
        /// <remarks/>
        [XmlElement(Namespace="urn:schemas-microsoft-com:vml")]
        public CT_Fill fill {
            get {
                return this.fillField;
            }
            set {
                this.fillField = value;
            }
        }
        
        /// <remarks/>
        [XmlElement(Namespace="urn:schemas-microsoft-com:vml")]
        public CT_Stroke stroke {
            get {
                return this.strokeField;
            }
            set {
                this.strokeField = value;
            }
        }
        
        /// <remarks/>
        [XmlElement(Namespace="urn:schemas-microsoft-com:vml")]
        public CT_Textbox textbox {
            get {
                return this.textboxField;
            }
            set {
                this.textboxField = value;
            }
        }
        
        /// <remarks/>
        [XmlElement(Namespace="urn:schemas-microsoft-com:vml")]
        public CT_Shadow shadow {
            get {
                return this.shadowField;
            }
            set {
                this.shadowField = value;
            }
        }
        
        /// <remarks/>
        public CT_Skew skew {
            get {
                return this.skewField;
            }
            set {
                this.skewField = value;
            }
        }
        
        /// <remarks/>
        public CT_Extrusion extrusion {
            get {
                return this.extrusionField;
            }
            set {
                this.extrusionField = value;
            }
        }
        
        /// <remarks/>
        public CT_Callout callout {
            get {
                return this.calloutField;
            }
            set {
                this.calloutField = value;
            }
        }
        
        /// <remarks/>
        public CT_Lock @lock {
            get {
                return this.lockField;
            }
            set {
                this.lockField = value;
            }
        }
        
        /// <remarks/>
        public CT_ColorMru colormru {
            get {
                return this.colormruField;
            }
            set {
                this.colormruField = value;
            }
        }
        
        /// <remarks/>
        public CT_ColorMenu colormenu {
            get {
                return this.colormenuField;
            }
            set {
                this.colormenuField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="urn:schemas-microsoft-com:vml")]
        public ST_Ext ext {
            get {
                return this.extField;
            }
            set {
                this.extField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool extSpecified {
            get {
                return this.extFieldSpecified;
            }
            set {
                this.extFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="integer")]
        public string spidmax {
            get {
                return this.spidmaxField;
            }
            set {
                this.spidmaxField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string style {
            get {
                return this.styleField;
            }
            set {
                this.styleField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute("fill")]
        public ST_TrueFalse1 fill1 {
            get {
                return this.fill1Field;
            }
            set {
                this.fill1Field = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool fill1Specified {
            get {
                return this.fill1FieldSpecified;
            }
            set {
                this.fill1FieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string fillcolor {
            get {
                return this.fillcolorField;
            }
            set {
                this.fillcolorField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute("stroke")]
        public ST_TrueFalse1 stroke1 {
            get {
                return this.stroke1Field;
            }
            set {
                this.stroke1Field = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool stroke1Specified {
            get {
                return this.stroke1FieldSpecified;
            }
            set {
                this.stroke1FieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string strokecolor {
            get {
                return this.strokecolorField;
            }
            set {
                this.strokecolorField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_TrueFalse1 allowincell {
            get {
                return this.allowincellField;
            }
            set {
                this.allowincellField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool allowincellSpecified {
            get {
                return this.allowincellFieldSpecified;
            }
            set {
                this.allowincellFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    public partial class CT_Fill {
        
        private string idField;
        
        private ST_FillType typeField;
        
        private bool typeFieldSpecified;
        
        private ST_TrueFalse onField;
        
        private bool onFieldSpecified;
        
        private string colorField;
        
        private string opacityField;
        
        private string color2Field;
        
        private string srcField;
        
        private string sizeField;
        
        private string originField;
        
        private string positionField;
        
        private ST_ImageAspect aspectField;
        
        private bool aspectFieldSpecified;
        
        private string colorsField;
        
        private decimal angleField;
        
        private bool angleFieldSpecified;
        
        private ST_TrueFalse alignshapeField;
        
        private bool alignshapeFieldSpecified;
        
        private string focusField;
        
        private string focussizeField;
        
        private string focuspositionField;
        
        private ST_FillMethod methodField;
        
        private bool methodFieldSpecified;
        
        private ST_TrueFalse recolorField;
        
        private bool recolorFieldSpecified;
        
        private ST_TrueFalse rotateField;
        
        private bool rotateFieldSpecified;
        
        private string id1Field;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string id {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_FillType type {
            get {
                return this.typeField;
            }
            set {
                this.typeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool typeSpecified {
            get {
                return this.typeFieldSpecified;
            }
            set {
                this.typeFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse on {
            get {
                return this.onField;
            }
            set {
                this.onField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool onSpecified {
            get {
                return this.onFieldSpecified;
            }
            set {
                this.onFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string color {
            get {
                return this.colorField;
            }
            set {
                this.colorField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string opacity {
            get {
                return this.opacityField;
            }
            set {
                this.opacityField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string color2 {
            get {
                return this.color2Field;
            }
            set {
                this.color2Field = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string src {
            get {
                return this.srcField;
            }
            set {
                this.srcField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string size {
            get {
                return this.sizeField;
            }
            set {
                this.sizeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string origin {
            get {
                return this.originField;
            }
            set {
                this.originField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string position {
            get {
                return this.positionField;
            }
            set {
                this.positionField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_ImageAspect aspect {
            get {
                return this.aspectField;
            }
            set {
                this.aspectField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool aspectSpecified {
            get {
                return this.aspectFieldSpecified;
            }
            set {
                this.aspectFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string colors {
            get {
                return this.colorsField;
            }
            set {
                this.colorsField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public decimal angle {
            get {
                return this.angleField;
            }
            set {
                this.angleField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool angleSpecified {
            get {
                return this.angleFieldSpecified;
            }
            set {
                this.angleFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse alignshape {
            get {
                return this.alignshapeField;
            }
            set {
                this.alignshapeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool alignshapeSpecified {
            get {
                return this.alignshapeFieldSpecified;
            }
            set {
                this.alignshapeFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string focus {
            get {
                return this.focusField;
            }
            set {
                this.focusField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string focussize {
            get {
                return this.focussizeField;
            }
            set {
                this.focussizeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string focusposition {
            get {
                return this.focuspositionField;
            }
            set {
                this.focuspositionField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_FillMethod method {
            get {
                return this.methodField;
            }
            set {
                this.methodField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool methodSpecified {
            get {
                return this.methodFieldSpecified;
            }
            set {
                this.methodFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse recolor {
            get {
                return this.recolorField;
            }
            set {
                this.recolorField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool recolorSpecified {
            get {
                return this.recolorFieldSpecified;
            }
            set {
                this.recolorFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse rotate {
            get {
                return this.rotateField;
            }
            set {
                this.rotateField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool rotateSpecified {
            get {
                return this.rotateFieldSpecified;
            }
            set {
                this.rotateFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute("id", Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id1 {
            get {
                return this.id1Field;
            }
            set {
                this.id1Field = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    public enum ST_FillType {
        
        /// <remarks/>
        solid,
        
        /// <remarks/>
        gradient,
        
        /// <remarks/>
        gradientRadial,
        
        /// <remarks/>
        tile,
        
        /// <remarks/>
        pattern,
        
        /// <remarks/>
        frame,
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    public enum ST_TrueFalse {
        
        /// <remarks/>
        t,
        
        /// <remarks/>
        f,
        
        /// <remarks/>
        @true,
        
        /// <remarks/>
        @false,
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    public enum ST_ImageAspect {
        
        /// <remarks/>
        ignore,
        
        /// <remarks/>
        atMost,
        
        /// <remarks/>
        atLeast,
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    public enum ST_FillMethod {
        
        /// <remarks/>
        none,
        
        /// <remarks/>
        linear,
        
        /// <remarks/>
        sigma,
        
        /// <remarks/>
        any,
        
        /// <remarks/>
        [XmlEnum("linear sigma")]
        linearsigma,
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_Skew {
        
        private ST_Ext extField;
        
        private bool extFieldSpecified;
        
        private string idField;
        
        private ST_TrueFalse1 onField;
        
        private bool onFieldSpecified;
        
        private string offsetField;
        
        private string originField;
        
        private string matrixField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="urn:schemas-microsoft-com:vml")]
        public ST_Ext ext {
            get {
                return this.extField;
            }
            set {
                this.extField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool extSpecified {
            get {
                return this.extFieldSpecified;
            }
            set {
                this.extFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string id {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 on {
            get {
                return this.onField;
            }
            set {
                this.onField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool onSpecified {
            get {
                return this.onFieldSpecified;
            }
            set {
                this.onFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string offset {
            get {
                return this.offsetField;
            }
            set {
                this.offsetField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string origin {
            get {
                return this.originField;
            }
            set {
                this.originField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string matrix {
            get {
                return this.matrixField;
            }
            set {
                this.matrixField = value;
            }
        }
    }
        
    /// <remarks/>
    [System.Serializable]
    [XmlType(TypeName="ST_TrueFalse", Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot("ST_TrueFalse", Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_TrueFalse1 {
        
        /// <remarks/>
        t,
        
        /// <remarks/>
        f,
        
        /// <remarks/>
        @true,
        
        /// <remarks/>
        @false,
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_Extrusion {
        
        private ST_Ext extField;
        
        private bool extFieldSpecified;
        
        private ST_TrueFalse1 onField;
        
        private bool onFieldSpecified;
        
        private ST_ExtrusionType typeField;
        
        private ST_ExtrusionRender renderField;
        
        private string viewpointoriginField;
        
        private string viewpointField;
        
        private ST_ExtrusionPlane planeField;
        
        private float skewangleField;
        
        private bool skewangleFieldSpecified;
        
        private string skewamtField;
        
        private string foredepthField;
        
        private string backdepthField;
        
        private string orientationField;
        
        private float orientationangleField;
        
        private bool orientationangleFieldSpecified;
        
        private ST_TrueFalse1 lockrotationcenterField;
        
        private bool lockrotationcenterFieldSpecified;
        
        private ST_TrueFalse1 autorotationcenterField;
        
        private bool autorotationcenterFieldSpecified;
        
        private string rotationcenterField;
        
        private string rotationangleField;
        
        private ST_ColorMode colormodeField;
        
        private bool colormodeFieldSpecified;
        
        private string colorField;
        
        private float shininessField;
        
        private bool shininessFieldSpecified;
        
        private string specularityField;
        
        private string diffusityField;
        
        private ST_TrueFalse1 metalField;
        
        private bool metalFieldSpecified;
        
        private string edgeField;
        
        private string facetField;
        
        private ST_TrueFalse1 lightfaceField;
        
        private bool lightfaceFieldSpecified;
        
        private string brightnessField;
        
        private string lightpositionField;
        
        private string lightlevelField;
        
        private ST_TrueFalse1 lightharshField;
        
        private bool lightharshFieldSpecified;
        
        private string lightposition2Field;
        
        private string lightlevel2Field;
        
        private ST_TrueFalse1 lightharsh2Field;
        
        private bool lightharsh2FieldSpecified;
        
        public CT_Extrusion() {
            this.typeField = ST_ExtrusionType.parallel;
            this.renderField = ST_ExtrusionRender.solid;
            this.planeField = ST_ExtrusionPlane.XY;
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="urn:schemas-microsoft-com:vml")]
        public ST_Ext ext {
            get {
                return this.extField;
            }
            set {
                this.extField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool extSpecified {
            get {
                return this.extFieldSpecified;
            }
            set {
                this.extFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 on {
            get {
                return this.onField;
            }
            set {
                this.onField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool onSpecified {
            get {
                return this.onFieldSpecified;
            }
            set {
                this.onFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(ST_ExtrusionType.parallel)]
        public ST_ExtrusionType type {
            get {
                return this.typeField;
            }
            set {
                this.typeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(ST_ExtrusionRender.solid)]
        public ST_ExtrusionRender render {
            get {
                return this.renderField;
            }
            set {
                this.renderField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string viewpointorigin {
            get {
                return this.viewpointoriginField;
            }
            set {
                this.viewpointoriginField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string viewpoint {
            get {
                return this.viewpointField;
            }
            set {
                this.viewpointField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(ST_ExtrusionPlane.XY)]
        public ST_ExtrusionPlane plane {
            get {
                return this.planeField;
            }
            set {
                this.planeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public float skewangle {
            get {
                return this.skewangleField;
            }
            set {
                this.skewangleField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool skewangleSpecified {
            get {
                return this.skewangleFieldSpecified;
            }
            set {
                this.skewangleFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string skewamt {
            get {
                return this.skewamtField;
            }
            set {
                this.skewamtField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string foredepth {
            get {
                return this.foredepthField;
            }
            set {
                this.foredepthField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string backdepth {
            get {
                return this.backdepthField;
            }
            set {
                this.backdepthField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string orientation {
            get {
                return this.orientationField;
            }
            set {
                this.orientationField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public float orientationangle {
            get {
                return this.orientationangleField;
            }
            set {
                this.orientationangleField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool orientationangleSpecified {
            get {
                return this.orientationangleFieldSpecified;
            }
            set {
                this.orientationangleFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 lockrotationcenter {
            get {
                return this.lockrotationcenterField;
            }
            set {
                this.lockrotationcenterField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool lockrotationcenterSpecified {
            get {
                return this.lockrotationcenterFieldSpecified;
            }
            set {
                this.lockrotationcenterFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 autorotationcenter {
            get {
                return this.autorotationcenterField;
            }
            set {
                this.autorotationcenterField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool autorotationcenterSpecified {
            get {
                return this.autorotationcenterFieldSpecified;
            }
            set {
                this.autorotationcenterFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string rotationcenter {
            get {
                return this.rotationcenterField;
            }
            set {
                this.rotationcenterField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string rotationangle {
            get {
                return this.rotationangleField;
            }
            set {
                this.rotationangleField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_ColorMode colormode {
            get {
                return this.colormodeField;
            }
            set {
                this.colormodeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool colormodeSpecified {
            get {
                return this.colormodeFieldSpecified;
            }
            set {
                this.colormodeFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string color {
            get {
                return this.colorField;
            }
            set {
                this.colorField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public float shininess {
            get {
                return this.shininessField;
            }
            set {
                this.shininessField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool shininessSpecified {
            get {
                return this.shininessFieldSpecified;
            }
            set {
                this.shininessFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string specularity {
            get {
                return this.specularityField;
            }
            set {
                this.specularityField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string diffusity {
            get {
                return this.diffusityField;
            }
            set {
                this.diffusityField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 metal {
            get {
                return this.metalField;
            }
            set {
                this.metalField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool metalSpecified {
            get {
                return this.metalFieldSpecified;
            }
            set {
                this.metalFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string edge {
            get {
                return this.edgeField;
            }
            set {
                this.edgeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string facet {
            get {
                return this.facetField;
            }
            set {
                this.facetField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 lightface {
            get {
                return this.lightfaceField;
            }
            set {
                this.lightfaceField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool lightfaceSpecified {
            get {
                return this.lightfaceFieldSpecified;
            }
            set {
                this.lightfaceFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string brightness {
            get {
                return this.brightnessField;
            }
            set {
                this.brightnessField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string lightposition {
            get {
                return this.lightpositionField;
            }
            set {
                this.lightpositionField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string lightlevel {
            get {
                return this.lightlevelField;
            }
            set {
                this.lightlevelField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 lightharsh {
            get {
                return this.lightharshField;
            }
            set {
                this.lightharshField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool lightharshSpecified {
            get {
                return this.lightharshFieldSpecified;
            }
            set {
                this.lightharshFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string lightposition2 {
            get {
                return this.lightposition2Field;
            }
            set {
                this.lightposition2Field = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string lightlevel2 {
            get {
                return this.lightlevel2Field;
            }
            set {
                this.lightlevel2Field = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 lightharsh2 {
            get {
                return this.lightharsh2Field;
            }
            set {
                this.lightharsh2Field = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool lightharsh2Specified {
            get {
                return this.lightharsh2FieldSpecified;
            }
            set {
                this.lightharsh2FieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_ExtrusionType {
        
        /// <remarks/>
        perspective,
        
        /// <remarks/>
        parallel,
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_ExtrusionRender {
        
        /// <remarks/>
        solid,
        
        /// <remarks/>
        wireFrame,
        
        /// <remarks/>
        boundingCube,
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_ExtrusionPlane {
        
        /// <remarks/>
        XY,
        
        /// <remarks/>
        ZX,
        
        /// <remarks/>
        YZ,
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_ColorMode {
        
        /// <remarks/>
        auto,
        
        /// <remarks/>
        custom,
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_Callout {
        
        private ST_Ext extField;
        
        private bool extFieldSpecified;
        
        private ST_TrueFalse1 onField;
        
        private bool onFieldSpecified;
        
        private string typeField;
        
        private string gapField;
        
        private ST_Angle angleField;
        
        private bool angleFieldSpecified;
        
        private ST_TrueFalse1 dropautoField;
        
        private bool dropautoFieldSpecified;
        
        private string dropField;
        
        private string distanceField;
        
        private ST_TrueFalse1 lengthspecifiedField;
        
        private string lengthField;
        
        private ST_TrueFalse1 accentbarField;
        
        private bool accentbarFieldSpecified;
        
        private ST_TrueFalse1 textborderField;
        
        private bool textborderFieldSpecified;
        
        private ST_TrueFalse1 minusxField;
        
        private bool minusxFieldSpecified;
        
        private ST_TrueFalse1 minusyField;
        
        private bool minusyFieldSpecified;
        
        public CT_Callout() {
            this.lengthspecifiedField = ST_TrueFalse1.f;
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="urn:schemas-microsoft-com:vml")]
        public ST_Ext ext {
            get {
                return this.extField;
            }
            set {
                this.extField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool extSpecified {
            get {
                return this.extFieldSpecified;
            }
            set {
                this.extFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 on {
            get {
                return this.onField;
            }
            set {
                this.onField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool onSpecified {
            get {
                return this.onFieldSpecified;
            }
            set {
                this.onFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string type {
            get {
                return this.typeField;
            }
            set {
                this.typeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string gap {
            get {
                return this.gapField;
            }
            set {
                this.gapField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_Angle angle {
            get {
                return this.angleField;
            }
            set {
                this.angleField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool angleSpecified {
            get {
                return this.angleFieldSpecified;
            }
            set {
                this.angleFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 dropauto {
            get {
                return this.dropautoField;
            }
            set {
                this.dropautoField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool dropautoSpecified {
            get {
                return this.dropautoFieldSpecified;
            }
            set {
                this.dropautoFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string drop {
            get {
                return this.dropField;
            }
            set {
                this.dropField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string distance {
            get {
                return this.distanceField;
            }
            set {
                this.distanceField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(ST_TrueFalse1.f)]
        public ST_TrueFalse1 lengthspecified {
            get {
                return this.lengthspecifiedField;
            }
            set {
                this.lengthspecifiedField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string length {
            get {
                return this.lengthField;
            }
            set {
                this.lengthField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 accentbar {
            get {
                return this.accentbarField;
            }
            set {
                this.accentbarField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool accentbarSpecified {
            get {
                return this.accentbarFieldSpecified;
            }
            set {
                this.accentbarFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 textborder {
            get {
                return this.textborderField;
            }
            set {
                this.textborderField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool textborderSpecified {
            get {
                return this.textborderFieldSpecified;
            }
            set {
                this.textborderFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 minusx {
            get {
                return this.minusxField;
            }
            set {
                this.minusxField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool minusxSpecified {
            get {
                return this.minusxFieldSpecified;
            }
            set {
                this.minusxFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 minusy {
            get {
                return this.minusyField;
            }
            set {
                this.minusyField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool minusySpecified {
            get {
                return this.minusyFieldSpecified;
            }
            set {
                this.minusyFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_Angle {
        
        /// <remarks/>
        any,
        
        /// <remarks/>
        [XmlEnum("30")]
        Item30,
        
        /// <remarks/>
        [XmlEnum("45")]
        Item45,
        
        /// <remarks/>
        [XmlEnum("60")]
        Item60,
        
        /// <remarks/>
        [XmlEnum("90")]
        Item90,
        
        /// <remarks/>
        auto,
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_Lock {
        
        private ST_Ext extField;
        
        private bool extFieldSpecified;
        
        private ST_TrueFalse1 positionField;
        
        private bool positionFieldSpecified;
        
        private ST_TrueFalse1 selectionField;
        
        private bool selectionFieldSpecified;
        
        private ST_TrueFalse1 groupingField;
        
        private bool groupingFieldSpecified;
        
        private ST_TrueFalse1 ungroupingField;
        
        private bool ungroupingFieldSpecified;
        
        private ST_TrueFalse1 rotationField;
        
        private bool rotationFieldSpecified;
        
        private ST_TrueFalse1 croppingField;
        
        private bool croppingFieldSpecified;
        
        private ST_TrueFalse1 verticiesField;
        
        private bool verticiesFieldSpecified;
        
        private ST_TrueFalse1 adjusthandlesField;
        
        private bool adjusthandlesFieldSpecified;
        
        private ST_TrueFalse1 textField;
        
        private bool textFieldSpecified;
        
        private ST_TrueFalse1 aspectratioField;
        
        private bool aspectratioFieldSpecified;
        
        private ST_TrueFalse1 shapetypeField;
        
        private bool shapetypeFieldSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="urn:schemas-microsoft-com:vml")]
        public ST_Ext ext {
            get {
                return this.extField;
            }
            set {
                this.extField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool extSpecified {
            get {
                return this.extFieldSpecified;
            }
            set {
                this.extFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 position {
            get {
                return this.positionField;
            }
            set {
                this.positionField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool positionSpecified {
            get {
                return this.positionFieldSpecified;
            }
            set {
                this.positionFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 selection {
            get {
                return this.selectionField;
            }
            set {
                this.selectionField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool selectionSpecified {
            get {
                return this.selectionFieldSpecified;
            }
            set {
                this.selectionFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 grouping {
            get {
                return this.groupingField;
            }
            set {
                this.groupingField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool groupingSpecified {
            get {
                return this.groupingFieldSpecified;
            }
            set {
                this.groupingFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 ungrouping {
            get {
                return this.ungroupingField;
            }
            set {
                this.ungroupingField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool ungroupingSpecified {
            get {
                return this.ungroupingFieldSpecified;
            }
            set {
                this.ungroupingFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 rotation {
            get {
                return this.rotationField;
            }
            set {
                this.rotationField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool rotationSpecified {
            get {
                return this.rotationFieldSpecified;
            }
            set {
                this.rotationFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 cropping {
            get {
                return this.croppingField;
            }
            set {
                this.croppingField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool croppingSpecified {
            get {
                return this.croppingFieldSpecified;
            }
            set {
                this.croppingFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 verticies {
            get {
                return this.verticiesField;
            }
            set {
                this.verticiesField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool verticiesSpecified {
            get {
                return this.verticiesFieldSpecified;
            }
            set {
                this.verticiesFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 adjusthandles {
            get {
                return this.adjusthandlesField;
            }
            set {
                this.adjusthandlesField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool adjusthandlesSpecified {
            get {
                return this.adjusthandlesFieldSpecified;
            }
            set {
                this.adjusthandlesFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 text {
            get {
                return this.textField;
            }
            set {
                this.textField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool textSpecified {
            get {
                return this.textFieldSpecified;
            }
            set {
                this.textFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 aspectratio {
            get {
                return this.aspectratioField;
            }
            set {
                this.aspectratioField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool aspectratioSpecified {
            get {
                return this.aspectratioFieldSpecified;
            }
            set {
                this.aspectratioFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 shapetype {
            get {
                return this.shapetypeField;
            }
            set {
                this.shapetypeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool shapetypeSpecified {
            get {
                return this.shapetypeFieldSpecified;
            }
            set {
                this.shapetypeFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_ColorMru {
        
        private ST_Ext extField;
        
        private bool extFieldSpecified;
        
        private string colorsField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="urn:schemas-microsoft-com:vml")]
        public ST_Ext ext {
            get {
                return this.extField;
            }
            set {
                this.extField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool extSpecified {
            get {
                return this.extFieldSpecified;
            }
            set {
                this.extFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string colors {
            get {
                return this.colorsField;
            }
            set {
                this.colorsField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_ColorMenu {
        
        private ST_Ext extField;
        
        private bool extFieldSpecified;
        
        private string strokecolorField;
        
        private string fillcolorField;
        
        private string shadowcolorField;
        
        private string extrusioncolorField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="urn:schemas-microsoft-com:vml")]
        public ST_Ext ext {
            get {
                return this.extField;
            }
            set {
                this.extField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool extSpecified {
            get {
                return this.extFieldSpecified;
            }
            set {
                this.extFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string strokecolor {
            get {
                return this.strokecolorField;
            }
            set {
                this.strokecolorField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string fillcolor {
            get {
                return this.fillcolorField;
            }
            set {
                this.fillcolorField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string shadowcolor {
            get {
                return this.shadowcolorField;
            }
            set {
                this.shadowcolorField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string extrusioncolor {
            get {
                return this.extrusioncolorField;
            }
            set {
                this.extrusioncolorField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_Ink {
        
        private byte[] iField;
        
        private ST_TrueFalse1 annotationField;
        
        private bool annotationFieldSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="base64Binary")]
        public byte[] i {
            get {
                return this.iField;
            }
            set {
                this.iField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 annotation {
            get {
                return this.annotationField;
            }
            set {
                this.annotationField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool annotationSpecified {
            get {
                return this.annotationFieldSpecified;
            }
            set {
                this.annotationFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_SignatureLine {
        
        private ST_Ext extField;
        
        private bool extFieldSpecified;
        
        private ST_TrueFalse1 issignaturelineField;
        
        private bool issignaturelineFieldSpecified;
        
        private string idField;
        
        private string providField;
        
        private ST_TrueFalse1 signinginstructionssetField;
        
        private bool signinginstructionssetFieldSpecified;
        
        private ST_TrueFalse1 allowcommentsField;
        
        private bool allowcommentsFieldSpecified;
        
        private ST_TrueFalse1 showsigndateField;
        
        private bool showsigndateFieldSpecified;
        
        private string suggestedsignerField;
        
        private string suggestedsigner2Field;
        
        private string suggestedsigneremailField;
        
        private string signinginstructionsField;
        
        private string addlxmlField;
        
        private string sigprovurlField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="urn:schemas-microsoft-com:vml")]
        public ST_Ext ext {
            get {
                return this.extField;
            }
            set {
                this.extField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool extSpecified {
            get {
                return this.extFieldSpecified;
            }
            set {
                this.extFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 issignatureline {
            get {
                return this.issignaturelineField;
            }
            set {
                this.issignaturelineField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool issignaturelineSpecified {
            get {
                return this.issignaturelineFieldSpecified;
            }
            set {
                this.issignaturelineFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="token")]
        public string id {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="token")]
        public string provid {
            get {
                return this.providField;
            }
            set {
                this.providField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 signinginstructionsset {
            get {
                return this.signinginstructionssetField;
            }
            set {
                this.signinginstructionssetField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool signinginstructionssetSpecified {
            get {
                return this.signinginstructionssetFieldSpecified;
            }
            set {
                this.signinginstructionssetFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 allowcomments {
            get {
                return this.allowcommentsField;
            }
            set {
                this.allowcommentsField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool allowcommentsSpecified {
            get {
                return this.allowcommentsFieldSpecified;
            }
            set {
                this.allowcommentsFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 showsigndate {
            get {
                return this.showsigndateField;
            }
            set {
                this.showsigndateField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool showsigndateSpecified {
            get {
                return this.showsigndateFieldSpecified;
            }
            set {
                this.showsigndateFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string suggestedsigner {
            get {
                return this.suggestedsignerField;
            }
            set {
                this.suggestedsignerField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string suggestedsigner2 {
            get {
                return this.suggestedsigner2Field;
            }
            set {
                this.suggestedsigner2Field = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string suggestedsigneremail {
            get {
                return this.suggestedsigneremailField;
            }
            set {
                this.suggestedsigneremailField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string signinginstructions {
            get {
                return this.signinginstructionsField;
            }
            set {
                this.signinginstructionsField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string addlxml {
            get {
                return this.addlxmlField;
            }
            set {
                this.addlxmlField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string sigprovurl {
            get {
                return this.sigprovurlField;
            }
            set {
                this.sigprovurlField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_ShapeLayout {
        
        private CT_IdMap idmapField;
        
        private CT_RegroupTable regrouptableField;
        
        private CT_Rules rulesField;
        
        private ST_Ext extField;
        
        private bool extFieldSpecified;
        
        /// <remarks/>
        public CT_IdMap idmap {
            get {
                return this.idmapField;
            }
            set {
                this.idmapField = value;
            }
        }
        
        /// <remarks/>
        public CT_RegroupTable regrouptable {
            get {
                return this.regrouptableField;
            }
            set {
                this.regrouptableField = value;
            }
        }
        
        /// <remarks/>
        public CT_Rules rules {
            get {
                return this.rulesField;
            }
            set {
                this.rulesField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="urn:schemas-microsoft-com:vml")]
        public ST_Ext ext {
            get {
                return this.extField;
            }
            set {
                this.extField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool extSpecified {
            get {
                return this.extFieldSpecified;
            }
            set {
                this.extFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_IdMap {
        
        private ST_Ext extField;
        
        private bool extFieldSpecified;
        
        private string dataField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="urn:schemas-microsoft-com:vml")]
        public ST_Ext ext {
            get {
                return this.extField;
            }
            set {
                this.extField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool extSpecified {
            get {
                return this.extFieldSpecified;
            }
            set {
                this.extFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string data {
            get {
                return this.dataField;
            }
            set {
                this.dataField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_RegroupTable {
        
        private CT_Entry[] entryField;
        
        private ST_Ext extField;
        
        private bool extFieldSpecified;
        
        /// <remarks/>
        [XmlElement("entry")]
        public CT_Entry[] entry {
            get {
                return this.entryField;
            }
            set {
                this.entryField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="urn:schemas-microsoft-com:vml")]
        public ST_Ext ext {
            get {
                return this.extField;
            }
            set {
                this.extField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool extSpecified {
            get {
                return this.extFieldSpecified;
            }
            set {
                this.extFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_Entry {
        
        private int newField;
        
        private bool newFieldSpecified;
        
        private int oldField;
        
        private bool oldFieldSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public int @new {
            get {
                return this.newField;
            }
            set {
                this.newField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool newSpecified {
            get {
                return this.newFieldSpecified;
            }
            set {
                this.newFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public int old {
            get {
                return this.oldField;
            }
            set {
                this.oldField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool oldSpecified {
            get {
                return this.oldFieldSpecified;
            }
            set {
                this.oldFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_Rules {
        
        private CT_R[] rField;
        
        private ST_Ext extField;
        
        private bool extFieldSpecified;
        
        /// <remarks/>
        [XmlElement("r")]
        public CT_R[] r {
            get {
                return this.rField;
            }
            set {
                this.rField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="urn:schemas-microsoft-com:vml")]
        public ST_Ext ext {
            get {
                return this.extField;
            }
            set {
                this.extField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool extSpecified {
            get {
                return this.extFieldSpecified;
            }
            set {
                this.extFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_R {
        
        private CT_Proxy[] proxyField;
        
        private string idField;
        
        private ST_RType typeField;
        
        private bool typeFieldSpecified;
        
        private ST_How howField;
        
        private bool howFieldSpecified;
        
        private string idrefField;
        
        /// <remarks/>
        [XmlElement("proxy")]
        public CT_Proxy[] proxy {
            get {
                return this.proxyField;
            }
            set {
                this.proxyField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string id {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_RType type {
            get {
                return this.typeField;
            }
            set {
                this.typeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool typeSpecified {
            get {
                return this.typeFieldSpecified;
            }
            set {
                this.typeFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_How how {
            get {
                return this.howField;
            }
            set {
                this.howField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool howSpecified {
            get {
                return this.howFieldSpecified;
            }
            set {
                this.howFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string idref {
            get {
                return this.idrefField;
            }
            set {
                this.idrefField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_Proxy {
        
        private ST_TrueFalseBlank startField;
        
        private ST_TrueFalseBlank endField;
        
        private string idrefField;
        
        private int connectlocField;
        
        private bool connectlocFieldSpecified;
        
        public CT_Proxy() {
            this.startField = ST_TrueFalseBlank.@false;
            this.endField = ST_TrueFalseBlank.@false;
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(ST_TrueFalseBlank.@false)]
        public ST_TrueFalseBlank start {
            get {
                return this.startField;
            }
            set {
                this.startField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(ST_TrueFalseBlank.@false)]
        public ST_TrueFalseBlank end {
            get {
                return this.endField;
            }
            set {
                this.endField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string idref {
            get {
                return this.idrefField;
            }
            set {
                this.idrefField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public int connectloc {
            get {
                return this.connectlocField;
            }
            set {
                this.connectlocField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool connectlocSpecified {
            get {
                return this.connectlocFieldSpecified;
            }
            set {
                this.connectlocFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_RType {
        
        /// <remarks/>
        arc,
        
        /// <remarks/>
        callout,
        
        /// <remarks/>
        connector,
        
        /// <remarks/>
        align,
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_How {
        
        /// <remarks/>
        top,
        
        /// <remarks/>
        middle,
        
        /// <remarks/>
        bottom,
        
        /// <remarks/>
        left,
        
        /// <remarks/>
        center,
        
        /// <remarks/>
        right,
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_Diagram {
        
        private CT_RelationTable relationtableField;
        
        private ST_Ext extField;
        
        private bool extFieldSpecified;
        
        private string dgmstyleField;
        
        private ST_TrueFalse1 autoformatField;
        
        private bool autoformatFieldSpecified;
        
        private ST_TrueFalse1 reverseField;
        
        private bool reverseFieldSpecified;
        
        private ST_TrueFalse1 autolayoutField;
        
        private bool autolayoutFieldSpecified;
        
        private string dgmscalexField;
        
        private string dgmscaleyField;
        
        private string dgmfontsizeField;
        
        private string constrainboundsField;
        
        private string dgmbasetextscaleField;
        
        /// <remarks/>
        public CT_RelationTable relationtable {
            get {
                return this.relationtableField;
            }
            set {
                this.relationtableField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="urn:schemas-microsoft-com:vml")]
        public ST_Ext ext {
            get {
                return this.extField;
            }
            set {
                this.extField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool extSpecified {
            get {
                return this.extFieldSpecified;
            }
            set {
                this.extFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="integer")]
        public string dgmstyle {
            get {
                return this.dgmstyleField;
            }
            set {
                this.dgmstyleField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 autoformat {
            get {
                return this.autoformatField;
            }
            set {
                this.autoformatField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool autoformatSpecified {
            get {
                return this.autoformatFieldSpecified;
            }
            set {
                this.autoformatFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 reverse {
            get {
                return this.reverseField;
            }
            set {
                this.reverseField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool reverseSpecified {
            get {
                return this.reverseFieldSpecified;
            }
            set {
                this.reverseFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 autolayout {
            get {
                return this.autolayoutField;
            }
            set {
                this.autolayoutField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool autolayoutSpecified {
            get {
                return this.autolayoutFieldSpecified;
            }
            set {
                this.autolayoutFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="integer")]
        public string dgmscalex {
            get {
                return this.dgmscalexField;
            }
            set {
                this.dgmscalexField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="integer")]
        public string dgmscaley {
            get {
                return this.dgmscaleyField;
            }
            set {
                this.dgmscaleyField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="integer")]
        public string dgmfontsize {
            get {
                return this.dgmfontsizeField;
            }
            set {
                this.dgmfontsizeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string constrainbounds {
            get {
                return this.constrainboundsField;
            }
            set {
                this.constrainboundsField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="integer")]
        public string dgmbasetextscale {
            get {
                return this.dgmbasetextscaleField;
            }
            set {
                this.dgmbasetextscaleField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_RelationTable {
        
        private CT_Relation[] relField;
        
        private ST_Ext extField;
        
        private bool extFieldSpecified;
        
        /// <remarks/>
        [XmlElement("rel")]
        public CT_Relation[] rel {
            get {
                return this.relField;
            }
            set {
                this.relField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="urn:schemas-microsoft-com:vml")]
        public ST_Ext ext {
            get {
                return this.extField;
            }
            set {
                this.extField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool extSpecified {
            get {
                return this.extFieldSpecified;
            }
            set {
                this.extFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_Relation {
        
        private ST_Ext extField;
        
        private bool extFieldSpecified;
        
        private string idsrcField;
        
        private string iddestField;
        
        private string idcntrField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="urn:schemas-microsoft-com:vml")]
        public ST_Ext ext {
            get {
                return this.extField;
            }
            set {
                this.extField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool extSpecified {
            get {
                return this.extFieldSpecified;
            }
            set {
                this.extFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string idsrc {
            get {
                return this.idsrcField;
            }
            set {
                this.idsrcField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string iddest {
            get {
                return this.iddestField;
            }
            set {
                this.iddestField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string idcntr {
            get {
                return this.idcntrField;
            }
            set {
                this.idcntrField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_OLEObject {
        
        private ST_OLELinkType linkTypeField;
        
        private bool linkTypeFieldSpecified;
        
        private ST_TrueFalseBlank lockedFieldField;
        
        private bool lockedFieldFieldSpecified;
        
        private string fieldCodesField;
        
        private ST_OLEType typeField;
        
        private bool typeFieldSpecified;
        
        private string progIDField;
        
        private string shapeIDField;
        
        private ST_OLEDrawAspect drawAspectField;
        
        private bool drawAspectFieldSpecified;
        
        private string objectIDField;
        
        private string idField;
        
        private ST_OLEUpdateMode updateModeField;
        
        private bool updateModeFieldSpecified;
        
        /// <remarks/>
        public ST_OLELinkType LinkType {
            get {
                return this.linkTypeField;
            }
            set {
                this.linkTypeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool LinkTypeSpecified {
            get {
                return this.linkTypeFieldSpecified;
            }
            set {
                this.linkTypeFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        public ST_TrueFalseBlank LockedField {
            get {
                return this.lockedFieldField;
            }
            set {
                this.lockedFieldField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool LockedFieldSpecified {
            get {
                return this.lockedFieldFieldSpecified;
            }
            set {
                this.lockedFieldFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        public string FieldCodes {
            get {
                return this.fieldCodesField;
            }
            set {
                this.fieldCodesField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_OLEType Type {
            get {
                return this.typeField;
            }
            set {
                this.typeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool TypeSpecified {
            get {
                return this.typeFieldSpecified;
            }
            set {
                this.typeFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string ProgID {
            get {
                return this.progIDField;
            }
            set {
                this.progIDField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string ShapeID {
            get {
                return this.shapeIDField;
            }
            set {
                this.shapeIDField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_OLEDrawAspect DrawAspect {
            get {
                return this.drawAspectField;
            }
            set {
                this.drawAspectField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool DrawAspectSpecified {
            get {
                return this.drawAspectFieldSpecified;
            }
            set {
                this.drawAspectFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string ObjectID {
            get {
                return this.objectIDField;
            }
            set {
                this.objectIDField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_OLEUpdateMode UpdateMode {
            get {
                return this.updateModeField;
            }
            set {
                this.updateModeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool UpdateModeSpecified {
            get {
                return this.updateModeFieldSpecified;
            }
            set {
                this.updateModeFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_OLELinkType {
        
        /// <remarks/>
        Picture,
        
        /// <remarks/>
        Bitmap,
        
        /// <remarks/>
        EnhancedMetaFile,
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_OLEType {
        
        /// <remarks/>
        Embed,
        
        /// <remarks/>
        Link,
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_OLEDrawAspect {
        
        /// <remarks/>
        Content,
        
        /// <remarks/>
        Icon,
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_OLEUpdateMode {
        
        /// <remarks/>
        Always,
        
        /// <remarks/>
        OnCall,
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_Complex {
        
        private ST_Ext extField;
        
        private bool extFieldSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="urn:schemas-microsoft-com:vml")]
        public ST_Ext ext {
            get {
                return this.extField;
            }
            set {
                this.extField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool extSpecified {
            get {
                return this.extFieldSpecified;
            }
            set {
                this.extFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_StrokeChild {
        
        private ST_Ext extField;
        
        private bool extFieldSpecified;
        
        private ST_TrueFalse1 onField;
        
        private bool onFieldSpecified;
        
        private string weightField;
        
        private string colorField;
        
        private string color2Field;
        
        private string opacityField;
        
        private ST_StrokeLineStyle linestyleField;
        
        private bool linestyleFieldSpecified;
        
        private decimal miterlimitField;
        
        private bool miterlimitFieldSpecified;
        
        private ST_StrokeJoinStyle joinstyleField;
        
        private bool joinstyleFieldSpecified;
        
        private ST_StrokeEndCap endcapField;
        
        private bool endcapFieldSpecified;
        
        private string dashstyleField;
        
        private ST_TrueFalse1 insetpenField;
        
        private bool insetpenFieldSpecified;
        
        private ST_FillType filltypeField;
        
        private bool filltypeFieldSpecified;
        
        private string srcField;
        
        private ST_ImageAspect imageaspectField;
        
        private bool imageaspectFieldSpecified;
        
        private string imagesizeField;
        
        private ST_TrueFalse1 imagealignshapeField;
        
        private bool imagealignshapeFieldSpecified;
        
        private ST_StrokeArrowType startarrowField;
        
        private bool startarrowFieldSpecified;
        
        private ST_StrokeArrowWidth startarrowwidthField;
        
        private bool startarrowwidthFieldSpecified;
        
        private ST_StrokeArrowLength startarrowlengthField;
        
        private bool startarrowlengthFieldSpecified;
        
        private ST_StrokeArrowType endarrowField;
        
        private bool endarrowFieldSpecified;
        
        private ST_StrokeArrowWidth endarrowwidthField;
        
        private bool endarrowwidthFieldSpecified;
        
        private ST_StrokeArrowLength endarrowlengthField;
        
        private bool endarrowlengthFieldSpecified;
        
        private string hrefField;
        
        private string althrefField;
        
        private string titleField;
        
        private ST_TrueFalse1 forcedashField;
        
        private bool forcedashFieldSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="urn:schemas-microsoft-com:vml")]
        public ST_Ext ext {
            get {
                return this.extField;
            }
            set {
                this.extField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool extSpecified {
            get {
                return this.extFieldSpecified;
            }
            set {
                this.extFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 on {
            get {
                return this.onField;
            }
            set {
                this.onField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool onSpecified {
            get {
                return this.onFieldSpecified;
            }
            set {
                this.onFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string weight {
            get {
                return this.weightField;
            }
            set {
                this.weightField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string color {
            get {
                return this.colorField;
            }
            set {
                this.colorField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string color2 {
            get {
                return this.color2Field;
            }
            set {
                this.color2Field = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string opacity {
            get {
                return this.opacityField;
            }
            set {
                this.opacityField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_StrokeLineStyle linestyle {
            get {
                return this.linestyleField;
            }
            set {
                this.linestyleField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool linestyleSpecified {
            get {
                return this.linestyleFieldSpecified;
            }
            set {
                this.linestyleFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public decimal miterlimit {
            get {
                return this.miterlimitField;
            }
            set {
                this.miterlimitField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool miterlimitSpecified {
            get {
                return this.miterlimitFieldSpecified;
            }
            set {
                this.miterlimitFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_StrokeJoinStyle joinstyle {
            get {
                return this.joinstyleField;
            }
            set {
                this.joinstyleField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool joinstyleSpecified {
            get {
                return this.joinstyleFieldSpecified;
            }
            set {
                this.joinstyleFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_StrokeEndCap endcap {
            get {
                return this.endcapField;
            }
            set {
                this.endcapField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool endcapSpecified {
            get {
                return this.endcapFieldSpecified;
            }
            set {
                this.endcapFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string dashstyle {
            get {
                return this.dashstyleField;
            }
            set {
                this.dashstyleField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 insetpen {
            get {
                return this.insetpenField;
            }
            set {
                this.insetpenField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool insetpenSpecified {
            get {
                return this.insetpenFieldSpecified;
            }
            set {
                this.insetpenFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_FillType filltype {
            get {
                return this.filltypeField;
            }
            set {
                this.filltypeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool filltypeSpecified {
            get {
                return this.filltypeFieldSpecified;
            }
            set {
                this.filltypeFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string src {
            get {
                return this.srcField;
            }
            set {
                this.srcField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_ImageAspect imageaspect {
            get {
                return this.imageaspectField;
            }
            set {
                this.imageaspectField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool imageaspectSpecified {
            get {
                return this.imageaspectFieldSpecified;
            }
            set {
                this.imageaspectFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public string imagesize {
            get {
                return this.imagesizeField;
            }
            set {
                this.imagesizeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TrueFalse1 imagealignshape {
            get {
                return this.imagealignshapeField;
            }
            set {
                this.imagealignshapeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool imagealignshapeSpecified {
            get {
                return this.imagealignshapeFieldSpecified;
            }
            set {
                this.imagealignshapeFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_StrokeArrowType startarrow {
            get {
                return this.startarrowField;
            }
            set {
                this.startarrowField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool startarrowSpecified {
            get {
                return this.startarrowFieldSpecified;
            }
            set {
                this.startarrowFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_StrokeArrowWidth startarrowwidth {
            get {
                return this.startarrowwidthField;
            }
            set {
                this.startarrowwidthField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool startarrowwidthSpecified {
            get {
                return this.startarrowwidthFieldSpecified;
            }
            set {
                this.startarrowwidthFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_StrokeArrowLength startarrowlength {
            get {
                return this.startarrowlengthField;
            }
            set {
                this.startarrowlengthField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool startarrowlengthSpecified {
            get {
                return this.startarrowlengthFieldSpecified;
            }
            set {
                this.startarrowlengthFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_StrokeArrowType endarrow {
            get {
                return this.endarrowField;
            }
            set {
                this.endarrowField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool endarrowSpecified {
            get {
                return this.endarrowFieldSpecified;
            }
            set {
                this.endarrowFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_StrokeArrowWidth endarrowwidth {
            get {
                return this.endarrowwidthField;
            }
            set {
                this.endarrowwidthField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool endarrowwidthSpecified {
            get {
                return this.endarrowwidthFieldSpecified;
            }
            set {
                this.endarrowwidthFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_StrokeArrowLength endarrowlength {
            get {
                return this.endarrowlengthField;
            }
            set {
                this.endarrowlengthField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool endarrowlengthSpecified {
            get {
                return this.endarrowlengthFieldSpecified;
            }
            set {
                this.endarrowlengthFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string href {
            get {
                return this.hrefField;
            }
            set {
                this.hrefField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string althref {
            get {
                return this.althrefField;
            }
            set {
                this.althrefField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string title {
            get {
                return this.titleField;
            }
            set {
                this.titleField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_TrueFalse1 forcedash {
            get {
                return this.forcedashField;
            }
            set {
                this.forcedashField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool forcedashSpecified {
            get {
                return this.forcedashFieldSpecified;
            }
            set {
                this.forcedashFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_ClipPath {
        
        private string vField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string v {
            get {
                return this.vField;
            }
            set {
                this.vField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(TypeName="CT_Fill", Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot("CT_Fill", Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public partial class CT_Fill1 {
        
        private ST_Ext extField;
        
        private bool extFieldSpecified;
        
        private ST_FillType1 typeField;
        
        private bool typeFieldSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="urn:schemas-microsoft-com:vml")]
        public ST_Ext ext {
            get {
                return this.extField;
            }
            set {
                this.extField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool extSpecified {
            get {
                return this.extFieldSpecified;
            }
            set {
                this.extFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_FillType1 type {
            get {
                return this.typeField;
            }
            set {
                this.typeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool typeSpecified {
            get {
                return this.typeFieldSpecified;
            }
            set {
                this.typeFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(TypeName="ST_FillType", Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot("ST_FillType", Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_FillType1 {
        
        /// <remarks/>
        gradientCenter,
        
        /// <remarks/>
        solid,
        
        /// <remarks/>
        pattern,
        
        /// <remarks/>
        tile,
        
        /// <remarks/>
        frame,
        
        /// <remarks/>
        gradientUnscaled,
        
        /// <remarks/>
        gradientRadial,
        
        /// <remarks/>
        gradient,
        
        /// <remarks/>
        background,
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_BWMode {
        
        /// <remarks/>
        color,
        
        /// <remarks/>
        auto,
        
        /// <remarks/>
        grayScale,
        
        /// <remarks/>
        lightGrayscale,
        
        /// <remarks/>
        inverseGray,
        
        /// <remarks/>
        grayOutline,
        
        /// <remarks/>
        highContrast,
        
        /// <remarks/>
        black,
        
        /// <remarks/>
        white,
        
        /// <remarks/>
        hide,
        
        /// <remarks/>
        undrawn,
        
        /// <remarks/>
        blackTextAndLines,
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_ScreenSize {
        
        /// <remarks/>
        [XmlEnum("544,376")]
        Item544376,
        
        /// <remarks/>
        [XmlEnum("640,480")]
        Item640480,
        
        /// <remarks/>
        [XmlEnum("720,512")]
        Item720512,
        
        /// <remarks/>
        [XmlEnum("800,600")]
        Item800600,
        
        /// <remarks/>
        [XmlEnum("1024,768")]
        Item1024768,
        
        /// <remarks/>
        [XmlEnum("1152,862")]
        Item1152862,
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_InsetMode {
        
        /// <remarks/>
        auto,
        
        /// <remarks/>
        custom,
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_CalloutPlacement {
        
        /// <remarks/>
        top,
        
        /// <remarks/>
        center,
        
        /// <remarks/>
        bottom,
        
        /// <remarks/>
        user,
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_ConnectorType {
        
        /// <remarks/>
        none,
        
        /// <remarks/>
        straight,
        
        /// <remarks/>
        elbow,
        
        /// <remarks/>
        curved,
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_HrAlign {
        
        /// <remarks/>
        left,
        
        /// <remarks/>
        right,
        
        /// <remarks/>
        center,
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_ConnectType {
        
        /// <remarks/>
        none,
        
        /// <remarks/>
        rect,
        
        /// <remarks/>
        segments,
        
        /// <remarks/>
        custom,
    }
}
