using System;
using System.ComponentModel;
using System.Xml.Serialization;
using NPOI.OpenXmlFormats.Vml.Spreadsheet;
using System.Collections.Generic;
using System.Xml;
using System.IO;
using System.Text;
using NPOI.OpenXml4Net.Util;

namespace NPOI.OpenXmlFormats.Vml.Office
{
    
    

    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public class CT_ShapeDefaults {
        
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
        
        private ST_Ext extField = ST_Ext.NONE;
        
        
        
        private string spidmaxField;
        
        private string styleField;
        
        private ST_TrueFalse fill1Field;
        
        private bool fill1FieldSpecified;
        
        private string fillcolorField;
        
        private ST_TrueFalse stroke1Field;
        
        private bool stroke1FieldSpecified;
        
        private string strokecolorField;
        
        private ST_TrueFalse allowincellField;
        
        private bool allowincellFieldSpecified;
        
    
        [XmlElement(Namespace="urn:schemas-microsoft-com:vml")]
        public CT_Fill fill {
            get {
                return this.fillField;
            }
            set {
                this.fillField = value;
            }
        }
        
    
        [XmlElement(Namespace="urn:schemas-microsoft-com:vml")]
        public CT_Stroke stroke {
            get {
                return this.strokeField;
            }
            set {
                this.strokeField = value;
            }
        }
        
    
        [XmlElement(Namespace="urn:schemas-microsoft-com:vml")]
        public CT_Textbox textbox {
            get {
                return this.textboxField;
            }
            set {
                this.textboxField = value;
            }
        }
        
    
        [XmlElement(Namespace="urn:schemas-microsoft-com:vml")]
        public CT_Shadow shadow {
            get {
                return this.shadowField;
            }
            set {
                this.shadowField = value;
            }
        }
        
    
        public CT_Skew skew {
            get {
                return this.skewField;
            }
            set {
                this.skewField = value;
            }
        }
        
    
        public CT_Extrusion extrusion {
            get {
                return this.extrusionField;
            }
            set {
                this.extrusionField = value;
            }
        }
        
    
        public CT_Callout callout {
            get {
                return this.calloutField;
            }
            set {
                this.calloutField = value;
            }
        }
        
    
        public CT_Lock @lock {
            get {
                return this.lockField;
            }
            set {
                this.lockField = value;
            }
        }
        
    
        public CT_ColorMru colormru {
            get {
                return this.colormruField;
            }
            set {
                this.colormruField = value;
            }
        }
        
    
        public CT_ColorMenu colormenu {
            get {
                return this.colormenuField;
            }
            set {
                this.colormenuField = value;
            }
        }


        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:schemas-microsoft-com:vml")]
        public ST_Ext ext
        {
            get { return this.extField; }
            set { this.extField = value; }
        }
        [XmlIgnore]
        public bool extSpecified
        {
            get { return ST_Ext.NONE != this.extField; }
        }        
    
        [XmlAttribute(DataType="integer")]
        public string spidmax {
            get {
                return this.spidmaxField;
            }
            set {
                this.spidmaxField = value;
            }
        }
        
    
        [XmlAttribute]
        public string style {
            get {
                return this.styleField;
            }
            set {
                this.styleField = value;
            }
        }
        
    
        [XmlAttribute("fill")]
        public ST_TrueFalse fill1 {
            get {
                return this.fill1Field;
            }
            set {
                this.fill1Field = value;
            }
        }
        
    
        [XmlIgnore]
        public bool fill1Specified {
            get {
                return this.fill1FieldSpecified;
            }
            set {
                this.fill1FieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public string fillcolor {
            get {
                return this.fillcolorField;
            }
            set {
                this.fillcolorField = value;
            }
        }
        
    
        [XmlAttribute("stroke")]
        public ST_TrueFalse stroke1 {
            get {
                return this.stroke1Field;
            }
            set {
                this.stroke1Field = value;
            }
        }
        
    
        [XmlIgnore]
        public bool stroke1Specified {
            get {
                return this.stroke1FieldSpecified;
            }
            set {
                this.stroke1FieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public string strokecolor {
            get {
                return this.strokecolorField;
            }
            set {
                this.strokecolorField = value;
            }
        }
        
    
        [XmlAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_TrueFalse allowincell {
            get {
                return this.allowincellField;
            }
            set {
                this.allowincellField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool allowincellSpecified {
            get {
                return this.allowincellFieldSpecified;
            }
            set {
                this.allowincellFieldSpecified = value;
            }
        }
    }
    



    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public class CT_Skew {
        
        private ST_Ext extField = ST_Ext.NONE;
        
        
        
        private string idField;
        
        private ST_TrueFalse onField;
        
        private bool onFieldSpecified;
        
        private string offsetField;
        
        private string originField;
        
        private string matrixField;


        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:schemas-microsoft-com:vml")]
        public ST_Ext ext
        {
            get { return this.extField; }
            set { this.extField = value; }
        }
        [XmlIgnore]
        public bool extSpecified
        {
            get { return ST_Ext.NONE != this.extField; }
        }

        [XmlAttribute]
        public string id
        {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse on {
            get {
                return this.onField;
            }
            set {
                this.onField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool onSpecified {
            get {
                return this.onFieldSpecified;
            }
            set {
                this.onFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public string offset {
            get {
                return this.offsetField;
            }
            set {
                this.offsetField = value;
            }
        }
        
    
        [XmlAttribute]
        public string origin {
            get {
                return this.originField;
            }
            set {
                this.originField = value;
            }
        }
        
    
        [XmlAttribute]
        public string matrix {
            get {
                return this.matrixField;
            }
            set {
                this.matrixField = value;
            }
        }
    }
        

    [Serializable]
    [XmlType(TypeName="ST_TrueFalse", Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot("ST_TrueFalse", Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_TrueFalse {
        f,
        t,
        @true,
        @false,
    }
    

    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public class CT_Extrusion {
        
        private ST_Ext extField = ST_Ext.NONE;
        
        
        
        private ST_TrueFalse onField;
        
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
        
        private ST_TrueFalse lockrotationcenterField;
        
        private bool lockrotationcenterFieldSpecified;
        
        private ST_TrueFalse autorotationcenterField;
        
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
        
        private ST_TrueFalse metalField;
        
        private bool metalFieldSpecified;
        
        private string edgeField;
        
        private string facetField;
        
        private ST_TrueFalse lightfaceField;
        
        private bool lightfaceFieldSpecified;
        
        private string brightnessField;
        
        private string lightpositionField;
        
        private string lightlevelField;
        
        private ST_TrueFalse lightharshField;
        
        private bool lightharshFieldSpecified;
        
        private string lightposition2Field;
        
        private string lightlevel2Field;
        
        private ST_TrueFalse lightharsh2Field;
        
        private bool lightharsh2FieldSpecified;
        
        public CT_Extrusion() {
            this.typeField = ST_ExtrusionType.parallel;
            this.renderField = ST_ExtrusionRender.solid;
            this.planeField = ST_ExtrusionPlane.XY;
        }


        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:schemas-microsoft-com:vml")]
        public ST_Ext ext
        {
            get { return this.extField; }
            set { this.extField = value; }
        }
        [XmlIgnore]
        public bool extSpecified
        {
            get { return ST_Ext.NONE != this.extField; }
        }        
    
        [XmlAttribute]
        public ST_TrueFalse on {
            get {
                return this.onField;
            }
            set {
                this.onField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool onSpecified {
            get {
                return this.onFieldSpecified;
            }
            set {
                this.onFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValue(ST_ExtrusionType.parallel)]
        public ST_ExtrusionType type {
            get {
                return this.typeField;
            }
            set {
                this.typeField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValue(ST_ExtrusionRender.solid)]
        public ST_ExtrusionRender render {
            get {
                return this.renderField;
            }
            set {
                this.renderField = value;
            }
        }
        
    
        [XmlAttribute]
        public string viewpointorigin {
            get {
                return this.viewpointoriginField;
            }
            set {
                this.viewpointoriginField = value;
            }
        }
        
    
        [XmlAttribute]
        public string viewpoint {
            get {
                return this.viewpointField;
            }
            set {
                this.viewpointField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValue(ST_ExtrusionPlane.XY)]
        public ST_ExtrusionPlane plane {
            get {
                return this.planeField;
            }
            set {
                this.planeField = value;
            }
        }
        
    
        [XmlAttribute]
        public float skewangle {
            get {
                return this.skewangleField;
            }
            set {
                this.skewangleField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool skewangleSpecified {
            get {
                return this.skewangleFieldSpecified;
            }
            set {
                this.skewangleFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public string skewamt {
            get {
                return this.skewamtField;
            }
            set {
                this.skewamtField = value;
            }
        }
        
    
        [XmlAttribute]
        public string foredepth {
            get {
                return this.foredepthField;
            }
            set {
                this.foredepthField = value;
            }
        }
        
    
        [XmlAttribute]
        public string backdepth {
            get {
                return this.backdepthField;
            }
            set {
                this.backdepthField = value;
            }
        }
        
    
        [XmlAttribute]
        public string orientation {
            get {
                return this.orientationField;
            }
            set {
                this.orientationField = value;
            }
        }
        
    
        [XmlAttribute]
        public float orientationangle {
            get {
                return this.orientationangleField;
            }
            set {
                this.orientationangleField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool orientationangleSpecified {
            get {
                return this.orientationangleFieldSpecified;
            }
            set {
                this.orientationangleFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse lockrotationcenter {
            get {
                return this.lockrotationcenterField;
            }
            set {
                this.lockrotationcenterField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool lockrotationcenterSpecified {
            get {
                return this.lockrotationcenterFieldSpecified;
            }
            set {
                this.lockrotationcenterFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse autorotationcenter {
            get {
                return this.autorotationcenterField;
            }
            set {
                this.autorotationcenterField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool autorotationcenterSpecified {
            get {
                return this.autorotationcenterFieldSpecified;
            }
            set {
                this.autorotationcenterFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public string rotationcenter {
            get {
                return this.rotationcenterField;
            }
            set {
                this.rotationcenterField = value;
            }
        }
        
    
        [XmlAttribute]
        public string rotationangle {
            get {
                return this.rotationangleField;
            }
            set {
                this.rotationangleField = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_ColorMode colormode {
            get {
                return this.colormodeField;
            }
            set {
                this.colormodeField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool colormodeSpecified {
            get {
                return this.colormodeFieldSpecified;
            }
            set {
                this.colormodeFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public string color {
            get {
                return this.colorField;
            }
            set {
                this.colorField = value;
            }
        }
        
    
        [XmlAttribute]
        public float shininess {
            get {
                return this.shininessField;
            }
            set {
                this.shininessField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool shininessSpecified {
            get {
                return this.shininessFieldSpecified;
            }
            set {
                this.shininessFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public string specularity {
            get {
                return this.specularityField;
            }
            set {
                this.specularityField = value;
            }
        }
        
    
        [XmlAttribute]
        public string diffusity {
            get {
                return this.diffusityField;
            }
            set {
                this.diffusityField = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse metal {
            get {
                return this.metalField;
            }
            set {
                this.metalField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool metalSpecified {
            get {
                return this.metalFieldSpecified;
            }
            set {
                this.metalFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public string edge {
            get {
                return this.edgeField;
            }
            set {
                this.edgeField = value;
            }
        }
        
    
        [XmlAttribute]
        public string facet {
            get {
                return this.facetField;
            }
            set {
                this.facetField = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse lightface {
            get {
                return this.lightfaceField;
            }
            set {
                this.lightfaceField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool lightfaceSpecified {
            get {
                return this.lightfaceFieldSpecified;
            }
            set {
                this.lightfaceFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public string brightness {
            get {
                return this.brightnessField;
            }
            set {
                this.brightnessField = value;
            }
        }
        
    
        [XmlAttribute]
        public string lightposition {
            get {
                return this.lightpositionField;
            }
            set {
                this.lightpositionField = value;
            }
        }
        
    
        [XmlAttribute]
        public string lightlevel {
            get {
                return this.lightlevelField;
            }
            set {
                this.lightlevelField = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse lightharsh {
            get {
                return this.lightharshField;
            }
            set {
                this.lightharshField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool lightharshSpecified {
            get {
                return this.lightharshFieldSpecified;
            }
            set {
                this.lightharshFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public string lightposition2 {
            get {
                return this.lightposition2Field;
            }
            set {
                this.lightposition2Field = value;
            }
        }
        
    
        [XmlAttribute]
        public string lightlevel2 {
            get {
                return this.lightlevel2Field;
            }
            set {
                this.lightlevel2Field = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse lightharsh2 {
            get {
                return this.lightharsh2Field;
            }
            set {
                this.lightharsh2Field = value;
            }
        }
        
    
        [XmlIgnore]
        public bool lightharsh2Specified {
            get {
                return this.lightharsh2FieldSpecified;
            }
            set {
                this.lightharsh2FieldSpecified = value;
            }
        }
    }
    

    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_ExtrusionType {
        
    
        perspective,
        
    
        parallel,
    }
    

    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_ExtrusionRender {
        
    
        solid,
        
    
        wireFrame,
        
    
        boundingCube,
    }
    

    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_ExtrusionPlane {
        
    
        XY,
        
    
        ZX,
        
    
        YZ,
    }
    

    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_ColorMode {
        
    
        auto,
        
    
        custom,
    }
    

    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public class CT_Callout {
        
        private ST_Ext extField = ST_Ext.NONE;
        
        
        
        private ST_TrueFalse onField;
        
        private bool onFieldSpecified;
        
        private string typeField;
        
        private string gapField;
        
        private ST_Angle angleField;
        
        private bool angleFieldSpecified;
        
        private ST_TrueFalse dropautoField;
        
        private bool dropautoFieldSpecified;
        
        private string dropField;
        
        private string distanceField;
        
        private ST_TrueFalse lengthspecifiedField;
        
        private string lengthField;
        
        private ST_TrueFalse accentbarField;
        
        private bool accentbarFieldSpecified;
        
        private ST_TrueFalse textborderField;
        
        private bool textborderFieldSpecified;
        
        private ST_TrueFalse minusxField;
        
        private bool minusxFieldSpecified;
        
        private ST_TrueFalse minusyField;
        
        private bool minusyFieldSpecified;
        
        public CT_Callout() {
            this.lengthspecifiedField = ST_TrueFalse.f;
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:schemas-microsoft-com:vml")]
        public ST_Ext ext
        {
            get { return this.extField; }
            set { this.extField = value; }
        }
        [XmlIgnore]
        public bool extSpecified
        {
            get { return ST_Ext.NONE != this.extField; }
        }
    
        [XmlAttribute]
        public ST_TrueFalse on {
            get {
                return this.onField;
            }
            set {
                this.onField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool onSpecified {
            get {
                return this.onFieldSpecified;
            }
            set {
                this.onFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public string type {
            get {
                return this.typeField;
            }
            set {
                this.typeField = value;
            }
        }
        
    
        [XmlAttribute]
        public string gap {
            get {
                return this.gapField;
            }
            set {
                this.gapField = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_Angle angle {
            get {
                return this.angleField;
            }
            set {
                this.angleField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool angleSpecified {
            get {
                return this.angleFieldSpecified;
            }
            set {
                this.angleFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse dropauto {
            get {
                return this.dropautoField;
            }
            set {
                this.dropautoField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool dropautoSpecified {
            get {
                return this.dropautoFieldSpecified;
            }
            set {
                this.dropautoFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public string drop {
            get {
                return this.dropField;
            }
            set {
                this.dropField = value;
            }
        }
        
    
        [XmlAttribute]
        public string distance {
            get {
                return this.distanceField;
            }
            set {
                this.distanceField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValue(ST_TrueFalse.f)]
        public ST_TrueFalse lengthspecified {
            get {
                return this.lengthspecifiedField;
            }
            set {
                this.lengthspecifiedField = value;
            }
        }
        
    
        [XmlAttribute]
        public string length {
            get {
                return this.lengthField;
            }
            set {
                this.lengthField = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse accentbar {
            get {
                return this.accentbarField;
            }
            set {
                this.accentbarField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool accentbarSpecified {
            get {
                return this.accentbarFieldSpecified;
            }
            set {
                this.accentbarFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse textborder {
            get {
                return this.textborderField;
            }
            set {
                this.textborderField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool textborderSpecified {
            get {
                return this.textborderFieldSpecified;
            }
            set {
                this.textborderFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse minusx {
            get {
                return this.minusxField;
            }
            set {
                this.minusxField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool minusxSpecified {
            get {
                return this.minusxFieldSpecified;
            }
            set {
                this.minusxFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse minusy {
            get {
                return this.minusyField;
            }
            set {
                this.minusyField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool minusySpecified {
            get {
                return this.minusyFieldSpecified;
            }
            set {
                this.minusyFieldSpecified = value;
            }
        }
    }
    

    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_Angle {
        
    
        any,
        
    
        [XmlEnum("30")]
        Item30,
        
    
        [XmlEnum("45")]
        Item45,
        
    
        [XmlEnum("60")]
        Item60,
        
    
        [XmlEnum("90")]
        Item90,
        
    
        auto,
    }
    

    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public class CT_Lock {
        
        private ST_Ext extField = ST_Ext.NONE;
        
        
        
        private ST_TrueFalse positionField;
        
        private bool positionFieldSpecified;
        
        private ST_TrueFalse selectionField;
        
        private bool selectionFieldSpecified;
        
        private ST_TrueFalse groupingField;
        
        private bool groupingFieldSpecified;
        
        private ST_TrueFalse ungroupingField;
        
        private bool ungroupingFieldSpecified;
        
        private ST_TrueFalse rotationField;
        
        private bool rotationFieldSpecified;
        
        private ST_TrueFalse croppingField;
        
        private bool croppingFieldSpecified;
        
        private ST_TrueFalse verticiesField;
        
        private bool verticiesFieldSpecified;
        
        private ST_TrueFalse adjusthandlesField;
        
        private bool adjusthandlesFieldSpecified;
        
        private ST_TrueFalse textField;
        
        private bool textFieldSpecified;
        
        private ST_TrueFalse aspectratioField;
        
        private bool aspectratioFieldSpecified;
        
        private ST_TrueFalse shapetypeField;
        
        private bool shapetypeFieldSpecified;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:schemas-microsoft-com:vml")]
        public ST_Ext ext
        {
            get { return this.extField; }
            set { this.extField = value; }
        }
        [XmlIgnore]
        public bool extSpecified
        {
            get { return ST_Ext.NONE != this.extField; }
        }
    
        [XmlAttribute]
        public ST_TrueFalse position {
            get {
                return this.positionField;
            }
            set {
                this.positionField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool positionSpecified {
            get {
                return this.positionFieldSpecified;
            }
            set {
                this.positionFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse selection {
            get {
                return this.selectionField;
            }
            set {
                this.selectionField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool selectionSpecified {
            get {
                return this.selectionFieldSpecified;
            }
            set {
                this.selectionFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse grouping {
            get {
                return this.groupingField;
            }
            set {
                this.groupingField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool groupingSpecified {
            get {
                return this.groupingFieldSpecified;
            }
            set {
                this.groupingFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse ungrouping {
            get {
                return this.ungroupingField;
            }
            set {
                this.ungroupingField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool ungroupingSpecified {
            get {
                return this.ungroupingFieldSpecified;
            }
            set {
                this.ungroupingFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse rotation {
            get {
                return this.rotationField;
            }
            set {
                this.rotationField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool rotationSpecified {
            get {
                return this.rotationFieldSpecified;
            }
            set {
                this.rotationFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse cropping {
            get {
                return this.croppingField;
            }
            set {
                this.croppingField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool croppingSpecified {
            get {
                return this.croppingFieldSpecified;
            }
            set {
                this.croppingFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse verticies {
            get {
                return this.verticiesField;
            }
            set {
                this.verticiesField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool verticiesSpecified {
            get {
                return this.verticiesFieldSpecified;
            }
            set {
                this.verticiesFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse adjusthandles {
            get {
                return this.adjusthandlesField;
            }
            set {
                this.adjusthandlesField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool adjusthandlesSpecified {
            get {
                return this.adjusthandlesFieldSpecified;
            }
            set {
                this.adjusthandlesFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse text {
            get {
                return this.textField;
            }
            set {
                this.textField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool textSpecified {
            get {
                return this.textFieldSpecified;
            }
            set {
                this.textFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse aspectratio {
            get {
                return this.aspectratioField;
            }
            set {
                this.aspectratioField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool aspectratioSpecified {
            get {
                return this.aspectratioFieldSpecified;
            }
            set {
                this.aspectratioFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse shapetype {
            get {
                return this.shapetypeField;
            }
            set {
                this.shapetypeField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool shapetypeSpecified {
            get {
                return this.shapetypeFieldSpecified;
            }
            set {
                this.shapetypeFieldSpecified = value;
            }
        }

        public static CT_Lock Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Lock ctObj = new CT_Lock();
            if (node.Attributes["v:ext"] != null)
                ctObj.ext = (ST_Ext)Enum.Parse(typeof(ST_Ext), node.Attributes["v:ext"].Value);
            if (node.Attributes["position"] != null)
                ctObj.position = NPOI.OpenXmlFormats.Util.XmlHelper.ReadTrueFalse(node.Attributes["position"]);
            if (node.Attributes["selection"] != null)
                ctObj.selection = NPOI.OpenXmlFormats.Util.XmlHelper.ReadTrueFalse(node.Attributes["selection"]);
            if (node.Attributes["grouping"] != null)
                ctObj.grouping = NPOI.OpenXmlFormats.Util.XmlHelper.ReadTrueFalse(node.Attributes["grouping"]);
            if (node.Attributes["ungrouping"] != null)
                ctObj.ungrouping = NPOI.OpenXmlFormats.Util.XmlHelper.ReadTrueFalse(node.Attributes["ungrouping"]);
            if (node.Attributes["rotation"] != null)
                ctObj.rotation = NPOI.OpenXmlFormats.Util.XmlHelper.ReadTrueFalse(node.Attributes["rotation"]);
            if (node.Attributes["cropping"] != null)
                ctObj.cropping = NPOI.OpenXmlFormats.Util.XmlHelper.ReadTrueFalse(node.Attributes["cropping"]);
            if (node.Attributes["verticies"] != null)
                ctObj.verticies = NPOI.OpenXmlFormats.Util.XmlHelper.ReadTrueFalse(node.Attributes["verticies"]);
            if (node.Attributes["adjusthandles"] != null)
                ctObj.adjusthandles = NPOI.OpenXmlFormats.Util.XmlHelper.ReadTrueFalse(node.Attributes["adjusthandles"]);
            if (node.Attributes["text"] != null)
                ctObj.text = NPOI.OpenXmlFormats.Util.XmlHelper.ReadTrueFalse(node.Attributes["text"]);
            if (node.Attributes["aspectratio"] != null)
                ctObj.aspectratio = NPOI.OpenXmlFormats.Util.XmlHelper.ReadTrueFalse(node.Attributes["aspectratio"]);
            if (node.Attributes["shapetype"] != null)
                ctObj.shapetype = NPOI.OpenXmlFormats.Util.XmlHelper.ReadTrueFalse(node.Attributes["shapetype"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<o:{0}", nodeName));
            if(this.ext != ST_Ext.NONE)
                XmlHelper.WriteAttribute(sw, "v:ext", this.ext.ToString());
            NPOI.OpenXmlFormats.Util.XmlHelper.WriteAttribute(sw, "position", this.position);
            NPOI.OpenXmlFormats.Util.XmlHelper.WriteAttribute(sw, "selection", this.selection);
            NPOI.OpenXmlFormats.Util.XmlHelper.WriteAttribute(sw, "grouping", this.grouping);
            NPOI.OpenXmlFormats.Util.XmlHelper.WriteAttribute(sw, "ungrouping", this.ungrouping);
            NPOI.OpenXmlFormats.Util.XmlHelper.WriteAttribute(sw, "rotation", this.rotation);
            NPOI.OpenXmlFormats.Util.XmlHelper.WriteAttribute(sw, "cropping", this.cropping);
            NPOI.OpenXmlFormats.Util.XmlHelper.WriteAttribute(sw, "verticies", this.verticies);
            NPOI.OpenXmlFormats.Util.XmlHelper.WriteAttribute(sw, "adjusthandles", this.adjusthandles);
            NPOI.OpenXmlFormats.Util.XmlHelper.WriteAttribute(sw, "text", this.text);
            NPOI.OpenXmlFormats.Util.XmlHelper.WriteAttribute(sw, "aspectratio", this.aspectratio);
            NPOI.OpenXmlFormats.Util.XmlHelper.WriteAttribute(sw, "shapetype", this.shapetype);
            sw.Write("/>");
        }
    }
    

    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public class CT_ColorMru {
        
        private ST_Ext extField = ST_Ext.NONE;
        
        
        
        private string colorsField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:schemas-microsoft-com:vml")]
        public ST_Ext ext
        {
            get { return this.extField; }
            set { this.extField = value; }
        }
        [XmlIgnore]
        public bool extSpecified
        {
            get { return ST_Ext.NONE != this.extField; }
        }
    
        [XmlAttribute]
        public string colors {
            get {
                return this.colorsField;
            }
            set {
                this.colorsField = value;
            }
        }
    }
    

    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public class CT_ColorMenu {
        
        private ST_Ext extField = ST_Ext.NONE;
        
        
        
        private string strokecolorField;
        
        private string fillcolorField;
        
        private string shadowcolorField;
        
        private string extrusioncolorField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:schemas-microsoft-com:vml")]
        public ST_Ext ext
        {
            get { return this.extField; }
            set { this.extField = value; }
        }
        [XmlIgnore]
        public bool extSpecified
        {
            get { return ST_Ext.NONE != this.extField; }
        }
    
        [XmlAttribute]
        public string strokecolor {
            get {
                return this.strokecolorField;
            }
            set {
                this.strokecolorField = value;
            }
        }
        
    
        [XmlAttribute]
        public string fillcolor {
            get {
                return this.fillcolorField;
            }
            set {
                this.fillcolorField = value;
            }
        }
        
    
        [XmlAttribute]
        public string shadowcolor {
            get {
                return this.shadowcolorField;
            }
            set {
                this.shadowcolorField = value;
            }
        }
        
    
        [XmlAttribute]
        public string extrusioncolor {
            get {
                return this.extrusioncolorField;
            }
            set {
                this.extrusioncolorField = value;
            }
        }
    }
    

    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public class CT_Ink {
        
        private byte[] iField;
        
        private ST_TrueFalse annotationField;
        
        private bool annotationFieldSpecified;
        
    
        [XmlAttribute(DataType="base64Binary")]
        public byte[] i {
            get {
                return this.iField;
            }
            set {
                this.iField = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse annotation {
            get {
                return this.annotationField;
            }
            set {
                this.annotationField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool annotationSpecified {
            get {
                return this.annotationFieldSpecified;
            }
            set {
                this.annotationFieldSpecified = value;
            }
        }
    }
    

    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public class CT_SignatureLine {
        
        private ST_Ext extField = ST_Ext.NONE;
        
        
        
        private ST_TrueFalse issignaturelineField;
        
        private bool issignaturelineFieldSpecified;
        
        private string idField;
        
        private string providField;
        
        private ST_TrueFalse signinginstructionssetField;
        
        private bool signinginstructionssetFieldSpecified;
        
        private ST_TrueFalse allowcommentsField;
        
        private bool allowcommentsFieldSpecified;
        
        private ST_TrueFalse showsigndateField;
        
        private bool showsigndateFieldSpecified;
        
        private string suggestedsignerField;
        
        private string suggestedsigner2Field;
        
        private string suggestedsigneremailField;
        
        private string signinginstructionsField;
        
        private string addlxmlField;
        
        private string sigprovurlField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:schemas-microsoft-com:vml")]
        public ST_Ext ext
        {
            get { return this.extField; }
            set { this.extField = value; }
        }
        [XmlIgnore]
        public bool extSpecified
        {
            get { return ST_Ext.NONE != this.extField; }
        }
    
        [XmlAttribute]
        public ST_TrueFalse issignatureline {
            get {
                return this.issignaturelineField;
            }
            set {
                this.issignaturelineField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool issignaturelineSpecified {
            get {
                return this.issignaturelineFieldSpecified;
            }
            set {
                this.issignaturelineFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute(DataType = "token")]
        public string id {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
        }
        
    
        [XmlAttribute(DataType="token")]
        public string provid {
            get {
                return this.providField;
            }
            set {
                this.providField = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse signinginstructionsset {
            get {
                return this.signinginstructionssetField;
            }
            set {
                this.signinginstructionssetField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool signinginstructionssetSpecified {
            get {
                return this.signinginstructionssetFieldSpecified;
            }
            set {
                this.signinginstructionssetFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse allowcomments {
            get {
                return this.allowcommentsField;
            }
            set {
                this.allowcommentsField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool allowcommentsSpecified {
            get {
                return this.allowcommentsFieldSpecified;
            }
            set {
                this.allowcommentsFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse showsigndate {
            get {
                return this.showsigndateField;
            }
            set {
                this.showsigndateField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool showsigndateSpecified {
            get {
                return this.showsigndateFieldSpecified;
            }
            set {
                this.showsigndateFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string suggestedsigner {
            get {
                return this.suggestedsignerField;
            }
            set {
                this.suggestedsignerField = value;
            }
        }
        
    
        [XmlAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string suggestedsigner2 {
            get {
                return this.suggestedsigner2Field;
            }
            set {
                this.suggestedsigner2Field = value;
            }
        }
        
    
        [XmlAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string suggestedsigneremail {
            get {
                return this.suggestedsigneremailField;
            }
            set {
                this.suggestedsigneremailField = value;
            }
        }
        
    
        [XmlAttribute]
        public string signinginstructions {
            get {
                return this.signinginstructionsField;
            }
            set {
                this.signinginstructionsField = value;
            }
        }
        
    
        [XmlAttribute]
        public string addlxml {
            get {
                return this.addlxmlField;
            }
            set {
                this.addlxmlField = value;
            }
        }
        
    
        [XmlAttribute]
        public string sigprovurl {
            get {
                return this.sigprovurlField;
            }
            set {
                this.sigprovurlField = value;
            }
        }
    }
    

    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "urn:schemas-microsoft-com:office:office")]
    [XmlRoot("shapelayout",Namespace = "urn:schemas-microsoft-com:office:office", IsNullable = true)]
    public class CT_ShapeLayout
    {
        private CT_IdMap idmapField = null;
        private CT_RegroupTable regrouptableField = null;
        private CT_Rules rulesField = null;
        private ST_Ext extField = ST_Ext.NONE;

        //static XmlSerializer serializer = new XmlSerializer(typeof(CT_ShapeLayout), "urn:schemas-microsoft-com:office:office");
        //public static CT_ShapeLayout Parse(string xmltext)
        //{
        //    TextReader tr = new StringReader(xmltext);
        //    CT_ShapeLayout obj = (CT_ShapeLayout)serializer.Deserialize(tr);
        //    return obj;
        //}

        public CT_IdMap AddNewIdmap()
        {
            idmapField = new CT_IdMap();
            return idmapField;
        }

        [XmlElement]
        public CT_IdMap idmap
        {
            get { return this.idmapField; }
            set { this.idmapField = value; }
        }

        [XmlElement]
        public CT_RegroupTable regrouptable
        {
            get { return this.regrouptableField; }
            set { this.regrouptableField = value; }
        }

        [XmlElement]
        public CT_Rules rules
        {
            get { return this.rulesField; }
            set { this.rulesField = value; }
        }


        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:schemas-microsoft-com:vml")]
        public ST_Ext ext
        {
            get { return this.extField; }
            set { this.extField = value; }
        }
        [XmlIgnore]
        public bool extSpecified
        {
            get { return ST_Ext.NONE != this.extField; }
        }
        //internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
        //    new XmlQualifiedName("o", "urn:schemas-microsoft-com:office:office"),
        //    new XmlQualifiedName("x", "urn:schemas-microsoft-com:office:excel"),
        //    new XmlQualifiedName("v", "urn:schemas-microsoft-com:vml")
        //});

        //public override string ToString()
        //{
        //    using (StringWriter stringWriter = new StringWriter())
        //    {
        //        XmlWriterSettings settings = new XmlWriterSettings();

        //        settings.Encoding = Encoding.UTF8;
        //        settings.OmitXmlDeclaration = true;
        //        using (XmlWriter writer = XmlWriter.Create(stringWriter, settings))
        //        {
        //            serializer.Serialize(writer, this, namespaces);
        //        }
        //        return stringWriter.ToString();
        //    }
        //}

        public static CT_ShapeLayout Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ShapeLayout ctObj = new CT_ShapeLayout();
            if (node.Attributes["v:ext"] != null)
                ctObj.ext = (ST_Ext)Enum.Parse(typeof(ST_Ext), node.Attributes["v:ext"].Value);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "idmap")
                    ctObj.idmap = CT_IdMap.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "regrouptable")
                    ctObj.regrouptable = CT_RegroupTable.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "rules")
                    ctObj.rules = CT_Rules.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        public void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<o:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "v:ext", this.ext.ToString());
            sw.Write(">");
            if (this.idmap != null)
                this.idmap.Write(sw, "idmap");
            if (this.regrouptable != null)
                this.regrouptable.Write(sw, "regrouptable");
            if (this.rules != null)
                this.rules.Write(sw, "rules");
            sw.Write(string.Format("</o:{0}>", nodeName));
        }

    }


    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace = "urn:schemas-microsoft-com:office:office", IsNullable = true)]
    public class CT_IdMap
    {
        private ST_Ext extField = ST_Ext.NONE;
        private string dataField = null;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:schemas-microsoft-com:vml")]
        public ST_Ext ext
        {
            get { return this.extField; }
            set { this.extField = value; }
        }
        [XmlIgnore]
        public bool extSpecified
        {
            get { return ST_Ext.NONE != this.extField; }
        }

        [XmlAttribute]
        public string data
        {
            get { return this.dataField; }
            set { this.dataField = value; }
        }
        [XmlIgnore]
        public bool dataSpecified
        {
            get { return null != this.dataField; }
        }

        public static CT_IdMap Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_IdMap ctObj = new CT_IdMap();
            if (node.Attributes["v:ext"] != null)
                ctObj.ext = (ST_Ext)Enum.Parse(typeof(ST_Ext), node.Attributes["v:ext"].Value);
            ctObj.data = XmlHelper.ReadString(node.Attributes["data"]);
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<o:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "v:ext", this.ext.ToString());
            XmlHelper.WriteAttribute(sw, "data", this.data);
            sw.Write("/>");
        }

    }


    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace = "urn:schemas-microsoft-com:office:office", IsNullable = true)]
    public class CT_RegroupTable
    {
        public static CT_RegroupTable Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_RegroupTable ctObj = new CT_RegroupTable();
            if (node.Attributes["v:ext"] != null)
                ctObj.ext = (ST_Ext)Enum.Parse(typeof(ST_Ext), node.Attributes["v:ext"].Value);
            ctObj.entry = new List<CT_Entry>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "entry")
                    ctObj.entry.Add(CT_Entry.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<o:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "v:ext", this.ext.ToString());
            sw.Write(">");
            if (this.entry != null)
            {
                foreach (CT_Entry x in this.entry)
                {
                    x.Write(sw, "entry");
                }
            }
            sw.Write(string.Format("</o:{0}>", nodeName));
        }


        private List<CT_Entry> entryField = null; // 0..*
        private ST_Ext extField = ST_Ext.NONE;

        [XmlElement("entry")]
        public List<CT_Entry> entry
        {
            get { return this.entryField; }
            set { this.entryField = value; }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:schemas-microsoft-com:vml")]
        public ST_Ext ext
        {
            get { return this.extField; }
            set { this.extField = value; }
        }
        [XmlIgnore]
        public bool extSpecified
        {
            get { return ST_Ext.NONE != this.extField; }
        }
    }
    

    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace = "urn:schemas-microsoft-com:office:office", IsNullable = true)]
    public class CT_Entry
    {
        private int? newField = null;
        private int? oldField = null;

        [XmlAttribute]
        public int @new
        {
            get { return (int)this.newField; }
            set { this.newField = value; }
        }
        [XmlIgnore]
        public bool newSpecified
        {
            get { return null != this.newField; }
        }

        [XmlAttribute]
        public int old
        {
            get { return (int)this.oldField; }
            set { this.oldField = value; }
        }
        [XmlIgnore]
        public bool oldSpecified
        {
            get { return null != this.oldField; }
        }
        public static CT_Entry Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Entry ctObj = new CT_Entry();
            if (node.Attributes["new"] != null)
                ctObj.@new = XmlHelper.ReadInt(node.Attributes["new"]);
            if (node.Attributes["old"] != null)
                ctObj.old = XmlHelper.ReadInt(node.Attributes["old"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<o:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "new", this.@new);
            XmlHelper.WriteAttribute(sw, "old", this.old);
            sw.Write(">");
            sw.Write(string.Format("</o:{0}>", nodeName));
        }

    }


    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace = "urn:schemas-microsoft-com:office:office", IsNullable = true)]
    public class CT_Rules
    {
        private List<CT_R> rField = null; // 0..*
        private ST_Ext extField = ST_Ext.NONE;


        [XmlElement("r")]
        public List<CT_R> r
        {
            get { return this.rField; }
            set { this.rField = value; }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:schemas-microsoft-com:vml")]
        public ST_Ext ext
        {
            get { return this.extField; }
            set { this.extField = value; }
        }
        [XmlIgnore]
        public bool extSpecified
        {
            get { return ST_Ext.NONE != this.extField; }
        }

        public static CT_Rules Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Rules ctObj = new CT_Rules();
            if (node.Attributes["v:ext"] != null)
                ctObj.ext = (ST_Ext)Enum.Parse(typeof(ST_Ext), node.Attributes["v:ext"].Value);
            ctObj.r = new List<CT_R>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "r")
                    ctObj.r.Add(CT_R.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<o:{0}", nodeName));
            if(this.ext!= ST_Ext.NONE)
                XmlHelper.WriteAttribute(sw, "v:ext", this.ext.ToString());
            sw.Write(">");
            if (this.r != null)
            {
                foreach (CT_R x in this.r)
                {
                    x.Write(sw, "r");
                }
            }
            sw.Write(string.Format("</o:{0}>", nodeName));
        }

    }


    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace = "urn:schemas-microsoft-com:office:office", IsNullable = true)]
    public class CT_R
    {

        private List<CT_Proxy> proxyField = null; // 0..*       
        private string idField = string.Empty; // 1..1
        private ST_RType typeField = ST_RType.NONE; // others optional
        private ST_How howField = ST_How.NONE;
        private string idrefField = null;

        [XmlElement("proxy")]
        public List<CT_Proxy> proxy
        {
            get { return this.proxyField; }
            set { this.proxyField = value; }
        }

        [XmlAttribute]
        public string id
        {
            get { return this.idField; }
            set { this.idField = value; }
        }

        [XmlAttribute]
        public ST_RType type
        {
            get { return this.typeField; }
            set { this.typeField = value; }
        }
        [XmlIgnore]
        public bool typeSpecified
        {
            get { return ST_RType.NONE != this.typeField; }
        }

        [XmlAttribute]
        public ST_How how
        {
            get { return this.howField; }
            set { this.howField = value; }
        }
        [XmlIgnore]
        public bool howSpecified
        {
            get { return ST_How.NONE != this.howField; }
        }

        [XmlAttribute]
        public string idref
        {
            get { return this.idrefField; }
            set { this.idrefField = value; }
        }
        [XmlIgnore]
        public bool idrefSpecified
        {
            get { return null != this.idrefField; }
        }

        public static CT_R Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_R ctObj = new CT_R();
            ctObj.id = XmlHelper.ReadString(node.Attributes["id"]);
            if (node.Attributes["type"] != null)
                ctObj.type = (ST_RType)Enum.Parse(typeof(ST_RType), node.Attributes["type"].Value);
            if (node.Attributes["how"] != null)
                ctObj.how = (ST_How)Enum.Parse(typeof(ST_How), node.Attributes["how"].Value);
            ctObj.idref = XmlHelper.ReadString(node.Attributes["idref"]);
            ctObj.proxy = new List<CT_Proxy>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "proxy")
                    ctObj.proxy.Add(CT_Proxy.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<o:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "id", this.id);
            if(this.type!=ST_RType.NONE)
                XmlHelper.WriteAttribute(sw, "type", this.type.ToString());
            if (this.how != ST_How.NONE)
                XmlHelper.WriteAttribute(sw, "how", this.how.ToString());
            XmlHelper.WriteAttribute(sw, "idref", this.idref);
            sw.Write(">");
            if (this.proxy != null)
            {
                foreach (CT_Proxy x in this.proxy)
                {
                    x.Write(sw, "proxy");
                }
            }
            sw.Write(string.Format("</o:{0}>", nodeName));
        }
    }


    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace = "urn:schemas-microsoft-com:office:office", IsNullable = true)]
    public class CT_Proxy
    {

        private ST_TrueFalseBlank startField = ST_TrueFalseBlank.NONE;
        private ST_TrueFalseBlank endField = ST_TrueFalseBlank.NONE;
        private string idrefField = null;
        private int? connectlocField;

        [XmlAttribute]
        [DefaultValue(ST_TrueFalseBlank.@false)]
        public ST_TrueFalseBlank start
        {
            get { return this.startField; }
            set { this.startField = value; }
        }
        [XmlIgnore]
        public bool startSpecified
        {
            get { return (ST_TrueFalseBlank.NONE != startField); }
        }

        [XmlAttribute]
        [DefaultValue(ST_TrueFalseBlank.@false)]
        public ST_TrueFalseBlank end
        {
            get { return this.endField; }
            set { this.endField = value; }
        }
        [XmlIgnore]
        public bool endSpecified
        {
            get { return (ST_TrueFalseBlank.NONE != endField); }
        }

        [XmlAttribute]
        public string idref
        {
            get { return this.idrefField; }
            set { this.idrefField = value; }
        }
        [XmlIgnore]
        public bool idrefSpecified
        {
            get { return (null != idrefField); }
        }

        [XmlAttribute]
        public int connectloc
        {
            get { return (int)this.connectlocField; }
            set { this.connectlocField = value; }
        }
        [XmlIgnore]
        public bool connectlocSpecified
        {
            get
            {
                return null != this.connectlocField;
            }
        }

        public static CT_Proxy Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Proxy ctObj = new CT_Proxy();
            if (node.Attributes["start"] != null)
                ctObj.start = NPOI.OpenXmlFormats.Util.XmlHelper.ReadTrueFalseBlank(node.Attributes["start"]);
            if (node.Attributes["end"] != null)
                ctObj.end = NPOI.OpenXmlFormats.Util.XmlHelper.ReadTrueFalseBlank(node.Attributes["end"]);
            ctObj.idref = XmlHelper.ReadString(node.Attributes["idref"]);
            if (node.Attributes["connectloc"] != null)
                ctObj.connectloc = XmlHelper.ReadInt(node.Attributes["connectloc"]);
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<o:{0}", nodeName));
            NPOI.OpenXmlFormats.Util.XmlHelper.WriteAttribute(sw, "start", this.start);
            NPOI.OpenXmlFormats.Util.XmlHelper.WriteAttribute(sw, "end", this.end);
            XmlHelper.WriteAttribute(sw, "idref", this.idref);
            XmlHelper.WriteAttribute(sw, "connectloc", this.connectloc);
            sw.Write(">");
            sw.Write(string.Format("</o:{0}>", nodeName));
        }


    }


    [Serializable]
    [XmlType(Namespace = "urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace = "urn:schemas-microsoft-com:office:office", IsNullable = false)]
    public enum ST_RType
    {
        NONE,
        arc,
        callout,
        connector,
        align,
    }


    [Serializable]
    [XmlType(Namespace = "urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace = "urn:schemas-microsoft-com:office:office", IsNullable = false)]
    public enum ST_How
    {
        NONE,
        top,
        middle,
        bottom,
        left,
        center,
        right,
    }
    

    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public class CT_Diagram {
        
        private CT_RelationTable relationtableField;
        
        private ST_Ext extField = ST_Ext.NONE;
        
        
        
        private string dgmstyleField;
        
        private ST_TrueFalse autoformatField;
        
        private bool autoformatFieldSpecified;
        
        private ST_TrueFalse reverseField;
        
        private bool reverseFieldSpecified;
        
        private ST_TrueFalse autolayoutField;
        
        private bool autolayoutFieldSpecified;
        
        private string dgmscalexField;
        
        private string dgmscaleyField;
        
        private string dgmfontsizeField;
        
        private string constrainboundsField;
        
        private string dgmbasetextscaleField;
        
    
        public CT_RelationTable relationtable {
            get {
                return this.relationtableField;
            }
            set {
                this.relationtableField = value;
            }
        }


        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:schemas-microsoft-com:vml")]
        public ST_Ext ext
        {
            get { return this.extField; }
            set { this.extField = value; }
        }
        [XmlIgnore]
        public bool extSpecified
        {
            get { return ST_Ext.NONE != this.extField; }
        }
        
    
        [XmlAttribute(DataType="integer")]
        public string dgmstyle {
            get {
                return this.dgmstyleField;
            }
            set {
                this.dgmstyleField = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse autoformat {
            get {
                return this.autoformatField;
            }
            set {
                this.autoformatField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool autoformatSpecified {
            get {
                return this.autoformatFieldSpecified;
            }
            set {
                this.autoformatFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse reverse {
            get {
                return this.reverseField;
            }
            set {
                this.reverseField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool reverseSpecified {
            get {
                return this.reverseFieldSpecified;
            }
            set {
                this.reverseFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse autolayout {
            get {
                return this.autolayoutField;
            }
            set {
                this.autolayoutField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool autolayoutSpecified {
            get {
                return this.autolayoutFieldSpecified;
            }
            set {
                this.autolayoutFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute(DataType="integer")]
        public string dgmscalex {
            get {
                return this.dgmscalexField;
            }
            set {
                this.dgmscalexField = value;
            }
        }
        
    
        [XmlAttribute(DataType="integer")]
        public string dgmscaley {
            get {
                return this.dgmscaleyField;
            }
            set {
                this.dgmscaleyField = value;
            }
        }
        
    
        [XmlAttribute(DataType="integer")]
        public string dgmfontsize {
            get {
                return this.dgmfontsizeField;
            }
            set {
                this.dgmfontsizeField = value;
            }
        }
        
    
        [XmlAttribute]
        public string constrainbounds {
            get {
                return this.constrainboundsField;
            }
            set {
                this.constrainboundsField = value;
            }
        }
        
    
        [XmlAttribute(DataType="integer")]
        public string dgmbasetextscale {
            get {
                return this.dgmbasetextscaleField;
            }
            set {
                this.dgmbasetextscaleField = value;
            }
        }
    }
    

    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public class CT_RelationTable {
        
        private  List<CT_Relation> relField;
        
        private ST_Ext extField = ST_Ext.NONE;
        
        
        
    
        [XmlElement("rel")]
        public  List<CT_Relation> rel {
            get {
                return this.relField;
            }
            set {
                this.relField = value;
            }
        }


        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:schemas-microsoft-com:vml")]
        public ST_Ext ext
        {
            get { return this.extField; }
            set { this.extField = value; }
        }
        [XmlIgnore]
        public bool extSpecified
        {
            get { return ST_Ext.NONE != this.extField; }
        }
    }
    

    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public class CT_Relation {
        
        private ST_Ext extField = ST_Ext.NONE;
        
        
        
        private string idsrcField;
        
        private string iddestField;
        
        private string idcntrField;


        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:schemas-microsoft-com:vml")]
        public ST_Ext ext
        {
            get { return this.extField; }
            set { this.extField = value; }
        }
        [XmlIgnore]
        public bool extSpecified
        {
            get { return ST_Ext.NONE != this.extField; }
        }       
    
        [XmlAttribute]
        public string idsrc {
            get {
                return this.idsrcField;
            }
            set {
                this.idsrcField = value;
            }
        }
        
    
        [XmlAttribute]
        public string iddest {
            get {
                return this.iddestField;
            }
            set {
                this.iddestField = value;
            }
        }
        
    
        [XmlAttribute]
        public string idcntr {
            get {
                return this.idcntrField;
            }
            set {
                this.idcntrField = value;
            }
        }
    }
    

    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public class CT_OLEObject {
        
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
        
    
        public ST_OLELinkType LinkType {
            get {
                return this.linkTypeField;
            }
            set {
                this.linkTypeField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool LinkTypeSpecified {
            get {
                return this.linkTypeFieldSpecified;
            }
            set {
                this.linkTypeFieldSpecified = value;
            }
        }
        
    
        public ST_TrueFalseBlank LockedField {
            get {
                return this.lockedFieldField;
            }
            set {
                this.lockedFieldField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool LockedFieldSpecified {
            get {
                return this.lockedFieldFieldSpecified;
            }
            set {
                this.lockedFieldFieldSpecified = value;
            }
        }
        
    
        public string FieldCodes {
            get {
                return this.fieldCodesField;
            }
            set {
                this.fieldCodesField = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_OLEType Type {
            get {
                return this.typeField;
            }
            set {
                this.typeField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool TypeSpecified {
            get {
                return this.typeFieldSpecified;
            }
            set {
                this.typeFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public string ProgID {
            get {
                return this.progIDField;
            }
            set {
                this.progIDField = value;
            }
        }
        
    
        [XmlAttribute]
        public string ShapeID {
            get {
                return this.shapeIDField;
            }
            set {
                this.shapeIDField = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_OLEDrawAspect DrawAspect {
            get {
                return this.drawAspectField;
            }
            set {
                this.drawAspectField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool DrawAspectSpecified {
            get {
                return this.drawAspectFieldSpecified;
            }
            set {
                this.drawAspectFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public string ObjectID {
            get {
                return this.objectIDField;
            }
            set {
                this.objectIDField = value;
            }
        }


        [XmlAttribute]
        public string id
        {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_OLEUpdateMode UpdateMode {
            get {
                return this.updateModeField;
            }
            set {
                this.updateModeField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool UpdateModeSpecified {
            get {
                return this.updateModeFieldSpecified;
            }
            set {
                this.updateModeFieldSpecified = value;
            }
        }
    }
    

    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_OLELinkType {
        
    
        Picture,
        
    
        Bitmap,
        
    
        EnhancedMetaFile,
    }
    

    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_OLEType {
        
    
        Embed,
        
    
        Link,
    }
    

    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_OLEDrawAspect {
        
    
        Content,
        
    
        Icon,
    }
    

    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_OLEUpdateMode {
        
    
        Always,
        
    
        OnCall,
    }


    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace = "urn:schemas-microsoft-com:office:office", IsNullable = true)]
    public class CT_Complex
    {

        private ST_Ext extField = ST_Ext.NONE;

        


        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:schemas-microsoft-com:vml")]
        public ST_Ext ext
        {
            get { return this.extField; }
            set { this.extField = value; }
        }
        [XmlIgnore]
        public bool extSpecified
        {
            get { return ST_Ext.NONE != this.extField; }
        }
    }    

    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public class CT_StrokeChild {
        
        private ST_Ext extField = ST_Ext.NONE;
        
        
        
        private ST_TrueFalse onField;
        
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
        
        private ST_TrueFalse insetpenField;
        
        private bool insetpenFieldSpecified;
        
        private ST_FillType filltypeField;
        
        private bool filltypeFieldSpecified;
        
        private string srcField;
        
        private ST_ImageAspect imageaspectField;
        
        private bool imageaspectFieldSpecified;
        
        private string imagesizeField;
        
        private ST_TrueFalse imagealignshapeField;
        
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
        
        private ST_TrueFalse forcedashField;
        
        private bool forcedashFieldSpecified;


        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:schemas-microsoft-com:vml")]
        public ST_Ext ext
        {
            get { return this.extField; }
            set { this.extField = value; }
        }
        [XmlIgnore]
        public bool extSpecified
        {
            get { return ST_Ext.NONE != this.extField; }
        }       
    
        [XmlAttribute]
        public ST_TrueFalse on {
            get {
                return this.onField;
            }
            set {
                this.onField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool onSpecified {
            get {
                return this.onFieldSpecified;
            }
            set {
                this.onFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public string weight {
            get {
                return this.weightField;
            }
            set {
                this.weightField = value;
            }
        }
        
    
        [XmlAttribute]
        public string color {
            get {
                return this.colorField;
            }
            set {
                this.colorField = value;
            }
        }
        
    
        [XmlAttribute]
        public string color2 {
            get {
                return this.color2Field;
            }
            set {
                this.color2Field = value;
            }
        }
        
    
        [XmlAttribute]
        public string opacity {
            get {
                return this.opacityField;
            }
            set {
                this.opacityField = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_StrokeLineStyle linestyle {
            get {
                return this.linestyleField;
            }
            set {
                this.linestyleField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool linestyleSpecified {
            get {
                return this.linestyleFieldSpecified;
            }
            set {
                this.linestyleFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public decimal miterlimit {
            get {
                return this.miterlimitField;
            }
            set {
                this.miterlimitField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool miterlimitSpecified {
            get {
                return this.miterlimitFieldSpecified;
            }
            set {
                this.miterlimitFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_StrokeJoinStyle joinstyle {
            get {
                return this.joinstyleField;
            }
            set {
                this.joinstyleField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool joinstyleSpecified {
            get {
                return this.joinstyleFieldSpecified;
            }
            set {
                this.joinstyleFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_StrokeEndCap endcap {
            get {
                return this.endcapField;
            }
            set {
                this.endcapField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool endcapSpecified {
            get {
                return this.endcapFieldSpecified;
            }
            set {
                this.endcapFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public string dashstyle {
            get {
                return this.dashstyleField;
            }
            set {
                this.dashstyleField = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse insetpen {
            get {
                return this.insetpenField;
            }
            set {
                this.insetpenField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool insetpenSpecified {
            get {
                return this.insetpenFieldSpecified;
            }
            set {
                this.insetpenFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_FillType filltype {
            get {
                return this.filltypeField;
            }
            set {
                this.filltypeField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool filltypeSpecified {
            get {
                return this.filltypeFieldSpecified;
            }
            set {
                this.filltypeFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public string src {
            get {
                return this.srcField;
            }
            set {
                this.srcField = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_ImageAspect imageaspect {
            get {
                return this.imageaspectField;
            }
            set {
                this.imageaspectField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool imageaspectSpecified {
            get {
                return this.imageaspectFieldSpecified;
            }
            set {
                this.imageaspectFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public string imagesize {
            get {
                return this.imagesizeField;
            }
            set {
                this.imagesizeField = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_TrueFalse imagealignshape {
            get {
                return this.imagealignshapeField;
            }
            set {
                this.imagealignshapeField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool imagealignshapeSpecified {
            get {
                return this.imagealignshapeFieldSpecified;
            }
            set {
                this.imagealignshapeFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_StrokeArrowType startarrow {
            get {
                return this.startarrowField;
            }
            set {
                this.startarrowField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool startarrowSpecified {
            get {
                return this.startarrowFieldSpecified;
            }
            set {
                this.startarrowFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_StrokeArrowWidth startarrowwidth {
            get {
                return this.startarrowwidthField;
            }
            set {
                this.startarrowwidthField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool startarrowwidthSpecified {
            get {
                return this.startarrowwidthFieldSpecified;
            }
            set {
                this.startarrowwidthFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_StrokeArrowLength startarrowlength {
            get {
                return this.startarrowlengthField;
            }
            set {
                this.startarrowlengthField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool startarrowlengthSpecified {
            get {
                return this.startarrowlengthFieldSpecified;
            }
            set {
                this.startarrowlengthFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_StrokeArrowType endarrow {
            get {
                return this.endarrowField;
            }
            set {
                this.endarrowField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool endarrowSpecified {
            get {
                return this.endarrowFieldSpecified;
            }
            set {
                this.endarrowFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_StrokeArrowWidth endarrowwidth {
            get {
                return this.endarrowwidthField;
            }
            set {
                this.endarrowwidthField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool endarrowwidthSpecified {
            get {
                return this.endarrowwidthFieldSpecified;
            }
            set {
                this.endarrowwidthFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_StrokeArrowLength endarrowlength {
            get {
                return this.endarrowlengthField;
            }
            set {
                this.endarrowlengthField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool endarrowlengthSpecified {
            get {
                return this.endarrowlengthFieldSpecified;
            }
            set {
                this.endarrowlengthFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string href {
            get {
                return this.hrefField;
            }
            set {
                this.hrefField = value;
            }
        }
        
    
        [XmlAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string althref {
            get {
                return this.althrefField;
            }
            set {
                this.althrefField = value;
            }
        }
        
    
        [XmlAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string title {
            get {
                return this.titleField;
            }
            set {
                this.titleField = value;
            }
        }
        
    
        [XmlAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_TrueFalse forcedash {
            get {
                return this.forcedashField;
            }
            set {
                this.forcedashField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool forcedashSpecified {
            get {
                return this.forcedashFieldSpecified;
            }
            set {
                this.forcedashFieldSpecified = value;
            }
        }
    }
    

    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public class CT_ClipPath {
        
        private string vField;
        
    
        [XmlAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string v {
            get {
                return this.vField;
            }
            set {
                this.vField = value;
            }
        }
    }
    

    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(TypeName="CT_Fill", Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot("CT_Fill", Namespace="urn:schemas-microsoft-com:office:office", IsNullable=true)]
    public class CT_Fill {
        
        private ST_Ext extField = ST_Ext.NONE;
        
        
        
        private ST_FillType1 typeField;
        
        private bool typeFieldSpecified;


        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:schemas-microsoft-com:vml")]
        public ST_Ext ext
        {
            get { return this.extField; }
            set { this.extField = value; }
        }
        [XmlIgnore]
        public bool extSpecified
        {
            get { return ST_Ext.NONE != this.extField; }
        }       
    
        [XmlAttribute]
        public ST_FillType1 type {
            get {
                return this.typeField;
            }
            set {
                this.typeField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool typeSpecified {
            get {
                return this.typeFieldSpecified;
            }
            set {
                this.typeFieldSpecified = value;
            }
        }
    }
    

    [Serializable]
    [XmlType(TypeName="ST_FillType", Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot("ST_FillType", Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_FillType1 {
        
    
        gradientCenter,
        
    
        solid,
        
    
        pattern,
        
    
        tile,
        
    
        frame,
        
    
        gradientUnscaled,
        
    
        gradientRadial,
        
    
        gradient,
        
    
        background,
    }
    

    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_BWMode {
        
    
        color,
        
    
        auto,
        
    
        grayScale,
        
    
        lightGrayscale,
        
    
        inverseGray,
        
    
        grayOutline,
        
    
        highContrast,
        
    
        black,
        
    
        white,
        
    
        hide,
        
    
        undrawn,
        
    
        blackTextAndLines,
    }
    

    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_ScreenSize {
        
    
        [XmlEnum("544,376")]
        Item544376,
        
    
        [XmlEnum("640,480")]
        Item640480,
        
    
        [XmlEnum("720,512")]
        Item720512,
        
    
        [XmlEnum("800,600")]
        Item800600,
        
    
        [XmlEnum("1024,768")]
        Item1024768,
        
    
        [XmlEnum("1152,862")]
        Item1152862,
    }
    

    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_InsetMode {
        auto,
        custom,
    }
    

    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_CalloutPlacement {
        
    
        top,
        
    
        center,
        
    
        bottom,
        
    
        user,
    }
    

    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_ConnectorType {
        
    
        none,
        
    
        straight,
        
    
        elbow,
        
    
        curved,
    }
    

    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_HrAlign {
        
    
        left,
        
    
        right,
        
    
        center,
    }
    

    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:office")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:office", IsNullable=false)]
    public enum ST_ConnectType {
        
    
        none,
        
    
        rect,
        
    
        segments,
        
    
        custom,
    }
}
