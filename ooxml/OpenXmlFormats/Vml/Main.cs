using System.Xml.Serialization;
using System.Collections.Generic;
using System.Xml.Schema;

namespace NPOI.OpenXmlFormats.Vml
{
    
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_Shape {
        
        private List<object> itemsField;
        
        private ItemsChoiceType1[] itemsElementNameField;
        
        private string typeField;
        
        private string adjField;
        private string styleField;
        private CT_Path pathField;
        
        private string equationxmlField;

        private CT_Wrap wrapField;
        private CT_Fill fillField;
        private CT_Formulas formulasField;
        private CT_Handles handlesField;
        private CT_ImageData imagedataField;
        private CT_Stroke strokeField;
        private CT_Shadow shadowField;
        private CT_Textbox textboxField;
        private CT_TextPath textpathField;

        
        
        [XmlElement("ClientData", typeof(CT_ClientData), Namespace="urn:schemas-microsoft-com:office:excel")]
        [XmlElement("iscomment", typeof(CT_Empty), Namespace="urn:schemas-microsoft-com:office:powerpoint")]
        [XmlElement("textdata", typeof(CT_Rel), Namespace="urn:schemas-microsoft-com:office:powerpoint")]
        [XmlElement("anchorlock", typeof(CT_AnchorLock), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderbottom", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderleft", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderright", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("bordertop", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemsElementName")]
        public List<object> Items {
            get {
                return this.itemsField;
            }
            set {
                this.itemsField = value;
            }
        }
        
        [XmlElement("ItemsElementName")]
        [System.Xml.Serialization.XmlIgnore]
        public ItemsChoiceType1[] ItemsElementName
        {
            get
            {
                return this.itemsElementNameField;
            }
            set
            {
                this.itemsElementNameField = value;
            }
        }

        [XmlElement(Namespace="urn:schemas-microsoft-com:office:word")]
        public CT_Wrap wrap
        {
            get { return this.wrapField; }
            set { this.wrapField = value; }
        }
        [XmlElement]
        public CT_Fill fill
        {
            get { return this.fillField; }
            set { this.fillField = value; }
        }
        [XmlElement]
        public CT_Formulas formulas
        {
            get { return this.formulasField; }
            set { this.formulasField = value; }
        }
        [XmlElement]
        public CT_Handles handles
        {
            get { return this.handlesField; }
            set { this.handlesField = value; }
        }
        [XmlElement]
        public CT_ImageData imagedata
        {
            get { return this.imagedataField; }
            set { this.imagedataField = value; }
        }
        [XmlElement]
        public CT_Stroke stroke
        {
            get { return this.strokeField; }
            set { this.strokeField = value; }
        }
        [XmlElement]
        public CT_Shadow shadow
        {
            get { return this.shadowField; }
            set { this.shadowField = value; }
        }

        public CT_Fill AddNewFill()
        {
            this.fillField=new CT_Fill();
            return this.fillField;
        }
        public CT_Shadow AddNewShadow()
        {
            this.shadowField = new CT_Shadow();
            return this.shadowField;
        }
        List<CT_ClientData> clientDataField = null;

        [XmlElement]
        public List<CT_ClientData> clientData
        {
            get { return clientDataField; }
            set { clientDataField = value; }
        }

        public CT_ClientData GetClientDataArray(int index)
        {
            return clientDataField != null ? this.clientDataField[index] : null;
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
        public string adj {
            get {
                return this.adjField;
            }
            set {
                this.adjField = value;
            }
        }
        
        
        [XmlElement]
        public CT_Path path {
            get {
                return this.pathField;
            }
            set {
                this.pathField = value;
            }
        }
        
        
        [XmlAttribute]
        public string equationxml {
            get {
                return this.equationxmlField;
            }
            set {
                this.equationxmlField = value;
            }
        }
        [XmlAttribute]
        public string style
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
    }
    
    
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "urn:schemas-microsoft-com:vml")]
    public partial class CT_Formulas
    {
        private List<CT_F> fField = null; // 0..* 

       
        [XmlElement("f", Form = XmlSchemaForm.Qualified, Namespace = "urn:schemas-microsoft-com:vml")]
        public List<CT_F> f
        {
            get { return this.fField; }
            set { this.fField = value; }
        }
        [XmlIgnore]
        public bool fSpecified
        {
            get { return (null != fField); }
        }
    }


    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "urn:schemas-microsoft-com:vml")]
    public partial class CT_F
    {
        private string eqnField = null;

        [XmlAttribute]
        public string eqn
        {
            get { return this.eqnField; }
            set { this.eqnField = value; }
        }
        [XmlIgnore]
        public bool eqnSpecified
        {
            get { return (null != eqnField); }
        }
    }


    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "urn:schemas-microsoft-com:vml")]
    public partial class CT_Handles
    {

        private List<CT_H> hField = null;

        [XmlElement("h")]
        public List<CT_H> h
        {
            get { return this.hField; }
            set { this.hField = value; }
        }
        [XmlIgnore]
        public bool hSpecified
        {
            get { return (null != hField); }
        }
    }
    
    
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    public partial class CT_H {
        
        private string positionField;
        
        private string polarField;
        
        private string mapField;
        
        private ST_TrueFalse invxField;
        
        private bool invxFieldSpecified;
        
        private ST_TrueFalse invyField;
        
        private bool invyFieldSpecified; // TODO remove
        
        private ST_TrueFalseBlank switchField;
        
        private bool switchFieldSpecified;
        
        private string xrangeField;
        
        private string yrangeField;
        
        private string radiusrangeField;
        
        
        [XmlAttribute]
        public string position {
            get {
                return this.positionField;
            }
            set {
                this.positionField = value;
            }
        }
        
        
        [XmlAttribute]
        public string polar {
            get {
                return this.polarField;
            }
            set {
                this.polarField = value;
            }
        }
        
        
        [XmlAttribute]
        public string map {
            get {
                return this.mapField;
            }
            set {
                this.mapField = value;
            }
        }
        
        
        [XmlAttribute]
        public ST_TrueFalse invx {
            get {
                return this.invxField;
            }
            set {
                this.invxField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool invxSpecified {
            get {
                return this.invxFieldSpecified;
            }
            set {
                this.invxFieldSpecified = value;
            }
        }
        
        
        [XmlAttribute]
        public ST_TrueFalse invy {
            get {
                return this.invyField;
            }
            set {
                this.invyField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool invySpecified {
            get {
                return this.invyFieldSpecified;
            }
            set {
                this.invyFieldSpecified = value;
            }
        }
        
        
        [XmlAttribute]
        public ST_TrueFalseBlank @switch {
            get {
                return this.switchField;
            }
            set {
                this.switchField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool switchSpecified {
            get {
                return this.switchFieldSpecified;
            }
            set {
                this.switchFieldSpecified = value;
            }
        }
        
        
        [XmlAttribute]
        public string xrange {
            get {
                return this.xrangeField;
            }
            set {
                this.xrangeField = value;
            }
        }
        
        
        [XmlAttribute]
        public string yrange {
            get {
                return this.yrangeField;
            }
            set {
                this.yrangeField = value;
            }
        }
        
        
        [XmlAttribute]
        public string radiusrange {
            get {
                return this.radiusrangeField;
            }
            set {
                this.radiusrangeField = value;
            }
        }
    }
    
    
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_ImageData {
        
        private string idField;
        
        private string srcField;
        
        private string cropleftField;
        
        private string croptopField;
        
        private string croprightField;
        
        private string cropbottomField;
        
        private string gainField;
        
        private string blacklevelField;
        
        private string gammaField;
        
        private ST_TrueFalse grayscaleField;
        
        private bool grayscaleFieldSpecified;
        
        private ST_TrueFalse bilevelField;
        
        private bool bilevelFieldSpecified;
        
        private string chromakeyField;
        
        private string embosscolorField;
        
        private string recolortargetField;
        
        private string id1Field;
        
        private string pictField;
        
        private string hrefField;
        
        
        // TODO is the following correct?
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
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
        public string src {
            get {
                return this.srcField;
            }
            set {
                this.srcField = value;
            }
        }
        
        
        [XmlAttribute]
        public string cropleft {
            get {
                return this.cropleftField;
            }
            set {
                this.cropleftField = value;
            }
        }
        
        
        [XmlAttribute]
        public string croptop {
            get {
                return this.croptopField;
            }
            set {
                this.croptopField = value;
            }
        }
        
        
        [XmlAttribute]
        public string cropright {
            get {
                return this.croprightField;
            }
            set {
                this.croprightField = value;
            }
        }
        
        
        [XmlAttribute]
        public string cropbottom {
            get {
                return this.cropbottomField;
            }
            set {
                this.cropbottomField = value;
            }
        }
        
        
        [XmlAttribute]
        public string gain {
            get {
                return this.gainField;
            }
            set {
                this.gainField = value;
            }
        }
        
        
        [XmlAttribute]
        public string blacklevel {
            get {
                return this.blacklevelField;
            }
            set {
                this.blacklevelField = value;
            }
        }
        
        
        [XmlAttribute]
        public string gamma {
            get {
                return this.gammaField;
            }
            set {
                this.gammaField = value;
            }
        }
        
        
        [XmlAttribute]
        public ST_TrueFalse grayscale {
            get {
                return this.grayscaleField;
            }
            set {
                this.grayscaleField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool grayscaleSpecified {
            get {
                return this.grayscaleFieldSpecified;
            }
            set {
                this.grayscaleFieldSpecified = value;
            }
        }
        
        
        [XmlAttribute]
        public ST_TrueFalse bilevel {
            get {
                return this.bilevelField;
            }
            set {
                this.bilevelField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool bilevelSpecified {
            get {
                return this.bilevelFieldSpecified;
            }
            set {
                this.bilevelFieldSpecified = value;
            }
        }
        
        
        [XmlAttribute]
        public string chromakey {
            get {
                return this.chromakeyField;
            }
            set {
                this.chromakeyField = value;
            }
        }
        
        
        [XmlAttribute]
        public string embosscolor {
            get {
                return this.embosscolorField;
            }
            set {
                this.embosscolorField = value;
            }
        }
        
        
        [XmlAttribute]
        public string recolortarget {
            get {
                return this.recolortargetField;
            }
            set {
                this.recolortargetField = value;
            }
        }
        
        
        [XmlAttributeAttribute("id", Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id1 {
            get {
                return this.id1Field;
            }
            set {
                this.id1Field = value;
            }
        }
        
        
        [XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string pict {
            get {
                return this.pictField;
            }
            set {
                this.pictField = value;
            }
        }
        
        
        [XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string href {
            get {
                return this.hrefField;
            }
            set {
                this.hrefField = value;
            }
        }
    }
    
    
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_Path {
        
        private string idField;
        
        private string vField;
        
        private string limoField;
        
        private string textboxrectField;
        
        private ST_TrueFalse fillokField;
        
        private bool fillokFieldSpecified;
        
        private ST_TrueFalse strokeokField;
        
        private bool strokeokFieldSpecified;
        
        private ST_TrueFalse shadowokField;
        
        private bool shadowokFieldSpecified;
        
        private ST_TrueFalse arrowokField;
        
        private bool arrowokFieldSpecified;
        
        private ST_TrueFalse gradientshapeokField;
        
        private bool gradientshapeokFieldSpecified;
        
        private ST_TrueFalse textpathokField;
        
        private bool textpathokFieldSpecified;
        
        private ST_TrueFalse insetpenokField;
        
        private bool insetpenokFieldSpecified;
        
        
        // TODO is the following correct?
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
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
        public string v {
            get {
                return this.vField;
            }
            set {
                this.vField = value;
            }
        }
        
        
        [XmlAttribute]
        public string limo {
            get {
                return this.limoField;
            }
            set {
                this.limoField = value;
            }
        }
        
        
        [XmlAttribute]
        public string textboxrect {
            get {
                return this.textboxrectField;
            }
            set {
                this.textboxrectField = value;
            }
        }
        
        
        [XmlAttribute]
        public ST_TrueFalse fillok {
            get {
                return this.fillokField;
            }
            set {
                this.fillokField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool fillokSpecified {
            get {
                return this.fillokFieldSpecified;
            }
            set {
                this.fillokFieldSpecified = value;
            }
        }
        
        
        [XmlAttribute]
        public ST_TrueFalse strokeok {
            get {
                return this.strokeokField;
            }
            set {
                this.strokeokField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool strokeokSpecified {
            get {
                return this.strokeokFieldSpecified;
            }
            set {
                this.strokeokFieldSpecified = value;
            }
        }
        
        
        [XmlAttribute]
        public ST_TrueFalse shadowok {
            get {
                return this.shadowokField;
            }
            set {
                this.shadowokField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool shadowokSpecified {
            get {
                return this.shadowokFieldSpecified;
            }
            set {
                this.shadowokFieldSpecified = value;
            }
        }
        
        
        [XmlAttribute]
        public ST_TrueFalse arrowok {
            get {
                return this.arrowokField;
            }
            set {
                this.arrowokField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool arrowokSpecified {
            get {
                return this.arrowokFieldSpecified;
            }
            set {
                this.arrowokFieldSpecified = value;
            }
        }
        
        
        [XmlAttribute]
        public ST_TrueFalse gradientshapeok {
            get {
                return this.gradientshapeokField;
            }
            set {
                this.gradientshapeokField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool gradientshapeokSpecified {
            get {
                return this.gradientshapeokFieldSpecified;
            }
            set {
                this.gradientshapeokFieldSpecified = value;
            }
        }
        
        
        [XmlAttribute]
        public ST_TrueFalse textpathok {
            get {
                return this.textpathokField;
            }
            set {
                this.textpathokField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool textpathokSpecified {
            get {
                return this.textpathokFieldSpecified;
            }
            set {
                this.textpathokFieldSpecified = value;
            }
        }
        
        
        [XmlAttribute]
        public ST_TrueFalse insetpenok {
            get {
                return this.insetpenokField;
            }
            set {
                this.insetpenokField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool insetpenokSpecified {
            get {
                return this.insetpenokFieldSpecified;
            }
            set {
                this.insetpenokFieldSpecified = value;
            }
        }
    }
    
    
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_Shadow {
        
        private string idField;
        
        private ST_TrueFalse onField;
        
        private bool onFieldSpecified;
        
        private ST_ShadowType typeField;
        
        private bool typeFieldSpecified;
        
        private ST_TrueFalse obscuredField;
        
        private bool obscuredFieldSpecified;
        
        private string colorField;
        
        private string opacityField;
        
        private string offsetField;
        
        private string color2Field;
        
        private string offset2Field;
        
        private string originField;
        
        private string matrixField;
        
        
        // TODO is the following correct?
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
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
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool onSpecified {
            get {
                return this.onFieldSpecified;
            }
            set {
                this.onFieldSpecified = value;
            }
        }
        
        
        [XmlAttribute]
        public ST_ShadowType type {
            get {
                return this.typeField;
            }
            set {
                this.typeField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool typeSpecified {
            get {
                return this.typeFieldSpecified;
            }
            set {
                this.typeFieldSpecified = value;
            }
        }
        
        
        [XmlAttribute]
        public ST_TrueFalse obscured {
            get {
                return this.obscuredField;
            }
            set {
                this.obscuredField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool obscuredSpecified {
            get {
                return this.obscuredFieldSpecified;
            }
            set {
                this.obscuredFieldSpecified = value;
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
        public string opacity {
            get {
                return this.opacityField;
            }
            set {
                this.opacityField = value;
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
        public string color2 {
            get {
                return this.color2Field;
            }
            set {
                this.color2Field = value;
            }
        }
        
        
        [XmlAttribute]
        public string offset2 {
            get {
                return this.offset2Field;
            }
            set {
                this.offset2Field = value;
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
    
    
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=false)]
    public enum ST_ShadowType {
        
        
        single,
        
        
        @double,
        
        
        emboss,
        
        
        perspective,
    }
    
    
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_Stroke {
        
        private string idField;
        
        private ST_TrueFalse onField;
        
        private bool onFieldSpecified;
        
        private string weightField;
        
        private string colorField;
        
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
        
        private ST_FillType filltypeField;
        
        private bool filltypeFieldSpecified;
        
        private string srcField;
        
        private ST_ImageAspect imageaspectField;
        
        private bool imageaspectFieldSpecified;
        
        private string imagesizeField;
        
        private ST_TrueFalse imagealignshapeField;
        
        private bool imagealignshapeFieldSpecified;
        
        private string color2Field;
        
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
        
        private string id1Field;
        
        private ST_TrueFalse insetpenField;
        
        private bool insetpenFieldSpecified;
        
        
        // TODO is the following correct?
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
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
        
        
        [System.Xml.Serialization.XmlIgnore]
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
        
        
        [System.Xml.Serialization.XmlIgnore]
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
        
        
        [System.Xml.Serialization.XmlIgnore]
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
        
        
        [System.Xml.Serialization.XmlIgnore]
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
        
        
        [System.Xml.Serialization.XmlIgnore]
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
        public ST_FillType filltype {
            get {
                return this.filltypeField;
            }
            set {
                this.filltypeField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
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
        
        
        [System.Xml.Serialization.XmlIgnore]
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
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool imagealignshapeSpecified {
            get {
                return this.imagealignshapeFieldSpecified;
            }
            set {
                this.imagealignshapeFieldSpecified = value;
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
        public ST_StrokeArrowType startarrow {
            get {
                return this.startarrowField;
            }
            set {
                this.startarrowField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
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
        
        
        [System.Xml.Serialization.XmlIgnore]
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
        
        
        [System.Xml.Serialization.XmlIgnore]
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
        
        
        [System.Xml.Serialization.XmlIgnore]
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
        
        
        [System.Xml.Serialization.XmlIgnore]
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
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool endarrowlengthSpecified {
            get {
                return this.endarrowlengthFieldSpecified;
            }
            set {
                this.endarrowlengthFieldSpecified = value;
            }
        }
        
        
        [XmlAttributeAttribute("id", Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id1 {
            get {
                return this.id1Field;
            }
            set {
                this.id1Field = value;
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
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool insetpenSpecified {
            get {
                return this.insetpenFieldSpecified;
            }
            set {
                this.insetpenFieldSpecified = value;
            }
        }
    }
    
    
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=false)]
    public enum ST_StrokeLineStyle {
        
        
        single,
        
        
        thinThin,
        
        
        thinThick,
        
        
        thickThin,
        
        
        thickBetweenThin,
    }
    
    
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=false)]
    public enum ST_StrokeJoinStyle {
        
        
        round,
        
        
        bevel,
        
        
        miter,
    }
    
    
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=false)]
    public enum ST_StrokeEndCap {
        
        
        flat,
        
        
        square,
        
        
        round,
    }
    
    
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=false)]
    public enum ST_StrokeArrowType {
        
        
        none,
        
        
        block,
        
        
        classic,
        
        
        oval,
        
        
        diamond,
        
        
        open,
    }
    
    
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=false)]
    public enum ST_StrokeArrowWidth {
        
        
        narrow,
        
        
        medium,
        
        
        wide,
    }
    
    
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=false)]
    public enum ST_StrokeArrowLength {
        
        
        @short,
        
        
        medium,
        
        
        @long,
    }
    
    
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_Textbox {
        
        private System.Xml.XmlElement itemField;
        
        private string idField;
        
        private string styleField;
        
        private string insetField;
        
        
        [System.Xml.Serialization.XmlAnyElementAttribute()]
        public System.Xml.XmlElement Item {
            get {
                return this.itemField;
            }
            set {
                this.itemField = value;
            }
        }
        
        
        // TODO is the following correct?
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
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
        public string style {
            get {
                return this.styleField;
            }
            set {
                this.styleField = value;
            }
        }
        
        
        [XmlAttribute]
        public string inset {
            get {
                return this.insetField;
            }
            set {
                this.insetField = value;
            }
        }
    }
    
    
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_TextPath {
        
        private string idField;
        
        private string styleField;
        
        private ST_TrueFalse onField;
        
        private bool onFieldSpecified;
        
        private ST_TrueFalse fitshapeField;
        
        private bool fitshapeFieldSpecified;
        
        private ST_TrueFalse fitpathField;
        
        private bool fitpathFieldSpecified;
        
        private ST_TrueFalse trimField;
        
        private bool trimFieldSpecified;
        
        private ST_TrueFalse xscaleField;
        
        private bool xscaleFieldSpecified;
        
        private string stringField;
        
        
        // TODO is the following correct?
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
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
        public string style {
            get {
                return this.styleField;
            }
            set {
                this.styleField = value;
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
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool onSpecified {
            get {
                return this.onFieldSpecified;
            }
            set {
                this.onFieldSpecified = value;
            }
        }
        
        
        [XmlAttribute]
        public ST_TrueFalse fitshape {
            get {
                return this.fitshapeField;
            }
            set {
                this.fitshapeField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool fitshapeSpecified {
            get {
                return this.fitshapeFieldSpecified;
            }
            set {
                this.fitshapeFieldSpecified = value;
            }
        }
        
        
        [XmlAttribute]
        public ST_TrueFalse fitpath {
            get {
                return this.fitpathField;
            }
            set {
                this.fitpathField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool fitpathSpecified {
            get {
                return this.fitpathFieldSpecified;
            }
            set {
                this.fitpathFieldSpecified = value;
            }
        }
        
        
        [XmlAttribute]
        public ST_TrueFalse trim {
            get {
                return this.trimField;
            }
            set {
                this.trimField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool trimSpecified {
            get {
                return this.trimFieldSpecified;
            }
            set {
                this.trimFieldSpecified = value;
            }
        }
        
        
        [XmlAttribute]
        public ST_TrueFalse xscale {
            get {
                return this.xscaleField;
            }
            set {
                this.xscaleField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool xscaleSpecified {
            get {
                return this.xscaleFieldSpecified;
            }
            set {
                this.xscaleFieldSpecified = value;
            }
        }
        
        
        [XmlAttribute]
        public string @string {
            get {
                return this.stringField;
            }
            set {
                this.stringField = value;
            }
        }
    }
    
    
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml", IncludeInSchema=false)]
    public enum ItemsChoiceType1 {
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:excel:ClientData")]
        ClientData,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:powerpoint:iscomment")]
        iscomment,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:powerpoint:textdata")]
        textdata,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:anchorlock")]
        anchorlock,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:borderbottom")]
        borderbottom,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:borderleft")]
        borderleft,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:borderright")]
        borderright,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:bordertop")]
        bordertop,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:wrap")]
        wrap,
        
        
        fill,
        
        
        formulas,
        
        
        handles,
        
        
        imagedata,
        
        
        path,
        
        
        shadow,
        
        
        stroke,
        
        
        textbox,
        
        
        textpath,
    }
    
    
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_Shapetype {
        
        private List<CT_Path> pathField;
        
        private List<CT_Formulas> formulasField;
        
        private List<CT_Handles> handlesField;
        
        private List<CT_Fill> fillField;
        
        private List<CT_Stroke> strokeField;
        
        private List<CT_Shadow> shadowField;
        
        private List<CT_Textbox> textboxField;
        
        private List<CT_TextPath> textpathField;
        
        private List<CT_ImageData> imagedataField;
        
        private List<CT_Wrap> wrapField;
        
        private List<CT_AnchorLock> anchorlockField;
        
        private List<CT_Border> bordertopField;
        
        private List<CT_Border> borderbottomField;
        
        private List<CT_Border> borderleftField;
        
        private List<CT_Border> borderrightField;
        
        private List<CT_ClientData> clientDataField;
        
        private List<CT_Rel> textdataField;
        
        private string adjField;
        private string idField;
        private string styleField;

        
        private string path1Field;

        // TODO is the following correct?
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id
        {
            get {
                return this.idField;
            }
            set 
            {
                this.idField = value;
            }
        }

        
        [XmlElement("path")]
        public List<CT_Path> path {
            get {
                return this.pathField;
            }
            set {
                this.pathField = value;
            }
        }
        
        
        [XmlElement("formulas")]
        public List<CT_Formulas> formulas {
            get {
                return this.formulasField;
            }
            set {
                this.formulasField = value;
            }
        }
        
        
        [XmlElement("handles")]
        public List<CT_Handles> handles {
            get {
                return this.handlesField;
            }
            set {
                this.handlesField = value;
            }
        }
        
        
        [XmlElement("fill")]
        public List<CT_Fill> fill {
            get {
                return this.fillField;
            }
            set {
                this.fillField = value;
            }
        }
        
        
        [XmlElement("stroke")]
        public List<CT_Stroke> stroke {
            get {
                return this.strokeField;
            }
            set {
                this.strokeField = value;
            }
        }
        
        
        [XmlElement("shadow")]
        public List<CT_Shadow> shadow {
            get {
                return this.shadowField;
            }
            set {
                this.shadowField = value;
            }
        }
        
        
        [XmlElement("textbox")]
        public List<CT_Textbox> textbox {
            get {
                return this.textboxField;
            }
            set {
                this.textboxField = value;
            }
        }
        
        
        [XmlElement("textpath")]
        public List<CT_TextPath> textpath {
            get {
                return this.textpathField;
            }
            set {
                this.textpathField = value;
            }
        }
        
        
        [XmlElement("imagedata")]
        public List<CT_ImageData> imagedata {
            get {
                return this.imagedataField;
            }
            set {
                this.imagedataField = value;
            }
        }
        
        
        [XmlElement("wrap", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Wrap> wrap {
            get {
                return this.wrapField;
            }
            set {
                this.wrapField = value;
            }
        }
        
        
        [XmlElement("anchorlock", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_AnchorLock> anchorlock {
            get {
                return this.anchorlockField;
            }
            set {
                this.anchorlockField = value;
            }
        }
        
        
        [XmlElement("bordertop", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Border> bordertop {
            get {
                return this.bordertopField;
            }
            set {
                this.bordertopField = value;
            }
        }
        
        
        [XmlElement("borderbottom", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Border> borderbottom {
            get {
                return this.borderbottomField;
            }
            set {
                this.borderbottomField = value;
            }
        }
        
        
        [XmlElement("borderleft", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Border> borderleft {
            get {
                return this.borderleftField;
            }
            set {
                this.borderleftField = value;
            }
        }
        
        
        [XmlElement("borderright", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Border> borderright {
            get {
                return this.borderrightField;
            }
            set {
                this.borderrightField = value;
            }
        }
        
        
        [XmlElement("ClientData", Namespace="urn:schemas-microsoft-com:office:excel")]
        public List<CT_ClientData> ClientData {
            get {
                return this.clientDataField;
            }
            set {
                this.clientDataField = value;
            }
        }
        
        
        [XmlElement("textdata", Namespace="urn:schemas-microsoft-com:office:powerpoint")]
        public List<CT_Rel> textdata {
            get {
                return this.textdataField;
            }
            set {
                this.textdataField = value;
            }
        }
        
        
        [XmlAttribute]
        public string adj {
            get {
                return this.adjField;
            }
            set {
                this.adjField = value;
            }
        }
        
        
        [XmlAttributeAttribute("path")]
        public string path1 {
            get {
                return this.path1Field;
            }
            set {
                this.path1Field = value;
            }
        }

        public CT_Stroke AddNewStroke()
        {
            CT_Stroke stk = new CT_Stroke();
            this.strokeField.Add(stk);
            return stk;
        }
        public CT_Path AddNewPath()
        {
            CT_Path p = new CT_Path();
            this.pathField.Add(p);
            return p;
        }

    }
    
    
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_Group {
        
        private List<object> itemsField;
        
        private ItemsChoiceType6[] itemsElementNameField;
        
        private ST_TrueFalse filledField;
        
        private bool filledFieldSpecified;
        
        private string fillcolorField;
        
        private ST_EditAs editasField;
        
        private bool editasFieldSpecified;
        
        
        [XmlElement("ClientData", typeof(CT_ClientData), Namespace="urn:schemas-microsoft-com:office:excel")]
        [XmlElement("textdata", typeof(CT_Rel), Namespace="urn:schemas-microsoft-com:office:powerpoint")]
        [XmlElement("anchorlock", typeof(CT_AnchorLock), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderbottom", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderleft", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderright", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("bordertop", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("wrap", typeof(CT_Wrap), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("arc", typeof(CT_Arc))]
        [XmlElement("curve", typeof(CT_Curve))]
        [XmlElement("fill", typeof(CT_Fill))]
        [XmlElement("formulas", typeof(CT_Formulas))]
        [XmlElement("group", typeof(CT_Group))]
        [XmlElement("handles", typeof(CT_Handles))]
        [XmlElement("image", typeof(CT_Image))]
        [XmlElement("imagedata", typeof(CT_ImageData))]
        [XmlElement("line", typeof(CT_Line))]
        [XmlElement("oval", typeof(CT_Oval))]
        [XmlElement("path", typeof(CT_Path))]
        [XmlElement("polyline", typeof(CT_PolyLine))]
        [XmlElement("rect", typeof(CT_Rect))]
        [XmlElement("roundrect", typeof(CT_RoundRect))]
        [XmlElement("shadow", typeof(CT_Shadow))]
        [XmlElement("shape", typeof(CT_Shape))]
        [XmlElement("shapetype", typeof(CT_Shapetype))]
        [XmlElement("stroke", typeof(CT_Stroke))]
        [XmlElement("textbox", typeof(CT_Textbox))]
        [XmlElement("textpath", typeof(CT_TextPath))]
        [System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemsElementName")]
        public List<object> Items {
            get {
                return this.itemsField;
            }
            set {
                this.itemsField = value;
            }
        }
        
        
        [XmlElement("ItemsElementName")]
        [System.Xml.Serialization.XmlIgnore]
        public ItemsChoiceType6[] ItemsElementName {
            get {
                return this.itemsElementNameField;
            }
            set {
                this.itemsElementNameField = value;
            }
        }
        
        
        [XmlAttribute]
        public ST_TrueFalse filled {
            get {
                return this.filledField;
            }
            set {
                this.filledField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool filledSpecified {
            get {
                return this.filledFieldSpecified;
            }
            set {
                this.filledFieldSpecified = value;
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
        public ST_EditAs editas {
            get {
                return this.editasField;
            }
            set {
                this.editasField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool editasSpecified {
            get {
                return this.editasFieldSpecified;
            }
            set {
                this.editasFieldSpecified = value;
            }
        }

        public CT_Shapetype AddNewShapetype()
        {
            throw new System.NotImplementedException();
        }
    }
    
    
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_Arc {
        
        private List<CT_Path> pathField;
        
        private List<CT_Formulas> formulasField;
        
        private List<CT_Handles> handlesField;
        
        private List<CT_Fill> fillField;
        
        private List<CT_Stroke> strokeField;
        
        private List<CT_Shadow> shadowField;
        
        private List<CT_Textbox> textboxField;
        
        private List<CT_TextPath> textpathField;
        
        private List<CT_ImageData> imagedataField;
        
        private List<CT_Wrap> wrapField;
        
        private List<CT_AnchorLock> anchorlockField;
        
        private List<CT_Border> bordertopField;
        
        private List<CT_Border> borderbottomField;
        
        private List<CT_Border> borderleftField;
        
        private List<CT_Border> borderrightField;
        
        private List<CT_ClientData> clientDataField;
        
        private List<CT_Rel> textdataField;
        
        private decimal startAngleField;
        
        private bool startAngleFieldSpecified;
        
        private decimal endAngleField;
        
        private bool endAngleFieldSpecified;
        
        
        [XmlElement("path")]
        public List<CT_Path> path {
            get {
                return this.pathField;
            }
            set {
                this.pathField = value;
            }
        }
        
        
        [XmlElement("formulas")]
        public List<CT_Formulas> formulas {
            get {
                return this.formulasField;
            }
            set {
                this.formulasField = value;
            }
        }
        
        
        [XmlElement("handles")]
        public List<CT_Handles> handles {
            get {
                return this.handlesField;
            }
            set {
                this.handlesField = value;
            }
        }
        
        
        [XmlElement("fill")]
        public List<CT_Fill> fill {
            get {
                return this.fillField;
            }
            set {
                this.fillField = value;
            }
        }
        
        
        [XmlElement("stroke")]
        public List<CT_Stroke> stroke {
            get {
                return this.strokeField;
            }
            set {
                this.strokeField = value;
            }
        }
        
        
        [XmlElement("shadow")]
        public List<CT_Shadow> shadow {
            get {
                return this.shadowField;
            }
            set {
                this.shadowField = value;
            }
        }
        
        
        [XmlElement("textbox")]
        public List<CT_Textbox> textbox {
            get {
                return this.textboxField;
            }
            set {
                this.textboxField = value;
            }
        }
        
        
        [XmlElement("textpath")]
        public List<CT_TextPath> textpath {
            get {
                return this.textpathField;
            }
            set {
                this.textpathField = value;
            }
        }
        
        
        [XmlElement("imagedata")]
        public List<CT_ImageData> imagedata {
            get {
                return this.imagedataField;
            }
            set {
                this.imagedataField = value;
            }
        }
        
        
        [XmlElement("wrap", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Wrap> wrap {
            get {
                return this.wrapField;
            }
            set {
                this.wrapField = value;
            }
        }
        
        
        [XmlElement("anchorlock", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_AnchorLock> anchorlock {
            get {
                return this.anchorlockField;
            }
            set {
                this.anchorlockField = value;
            }
        }
        
        
        [XmlElement("bordertop", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Border> bordertop {
            get {
                return this.bordertopField;
            }
            set {
                this.bordertopField = value;
            }
        }
        
        
        [XmlElement("borderbottom", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Border> borderbottom {
            get {
                return this.borderbottomField;
            }
            set {
                this.borderbottomField = value;
            }
        }
        
        
        [XmlElement("borderleft", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Border> borderleft {
            get {
                return this.borderleftField;
            }
            set {
                this.borderleftField = value;
            }
        }
        
        
        [XmlElement("borderright", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Border> borderright {
            get {
                return this.borderrightField;
            }
            set {
                this.borderrightField = value;
            }
        }
        
        
        [XmlElement("ClientData", Namespace="urn:schemas-microsoft-com:office:excel")]
        public List<CT_ClientData> ClientData {
            get {
                return this.clientDataField;
            }
            set {
                this.clientDataField = value;
            }
        }
        
        
        [XmlElement("textdata", Namespace="urn:schemas-microsoft-com:office:powerpoint")]
        public List<CT_Rel> textdata {
            get {
                return this.textdataField;
            }
            set {
                this.textdataField = value;
            }
        }
        
        
        [XmlAttribute]
        public decimal startAngle {
            get {
                return this.startAngleField;
            }
            set {
                this.startAngleField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool startAngleSpecified {
            get {
                return this.startAngleFieldSpecified;
            }
            set {
                this.startAngleFieldSpecified = value;
            }
        }
        
        
        [XmlAttribute]
        public decimal endAngle {
            get {
                return this.endAngleField;
            }
            set {
                this.endAngleField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool endAngleSpecified {
            get {
                return this.endAngleFieldSpecified;
            }
            set {
                this.endAngleFieldSpecified = value;
            }
        }
    }
    
    
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_Curve {
        
        private List<CT_Path> pathField;
        
        private List<CT_Formulas> formulasField;
        
        private List<CT_Handles> handlesField;
        
        private List<CT_Fill> fillField;
        
        private List<CT_Stroke> strokeField;
        
        private List<CT_Shadow> shadowField;
        
        private List<CT_Textbox> textboxField;
        
        private List<CT_TextPath> textpathField;
        
        private List<CT_ImageData> imagedataField;
        
        private List<CT_Wrap> wrapField;
        
        private List<CT_AnchorLock> anchorlockField;
        
        private List<CT_Border> bordertopField;
        
        private List<CT_Border> borderbottomField;
        
        private List<CT_Border> borderleftField;
        
        private List<CT_Border> borderrightField;
        
        private List<CT_ClientData> clientDataField;
        
        private List<CT_Rel> textdataField;
        
        private string fromField;
        
        private string control1Field;
        
        private string control2Field;
        
        private string toField;
        
        
        [XmlElement("path")]
        public List<CT_Path> path {
            get {
                return this.pathField;
            }
            set {
                this.pathField = value;
            }
        }
        
        
        [XmlElement("formulas")]
        public List<CT_Formulas> formulas {
            get {
                return this.formulasField;
            }
            set {
                this.formulasField = value;
            }
        }
        
        
        [XmlElement("handles")]
        public List<CT_Handles> handles {
            get {
                return this.handlesField;
            }
            set {
                this.handlesField = value;
            }
        }
        
        
        [XmlElement("fill")]
        public List<CT_Fill> fill {
            get {
                return this.fillField;
            }
            set {
                this.fillField = value;
            }
        }
        
        
        [XmlElement("stroke")]
        public List<CT_Stroke> stroke {
            get {
                return this.strokeField;
            }
            set {
                this.strokeField = value;
            }
        }
        
        
        [XmlElement("shadow")]
        public List<CT_Shadow> shadow {
            get {
                return this.shadowField;
            }
            set {
                this.shadowField = value;
            }
        }
        
        
        [XmlElement("textbox")]
        public List<CT_Textbox> textbox {
            get {
                return this.textboxField;
            }
            set {
                this.textboxField = value;
            }
        }
        
        
        [XmlElement("textpath")]
        public List<CT_TextPath> textpath {
            get {
                return this.textpathField;
            }
            set {
                this.textpathField = value;
            }
        }
        
        
        [XmlElement("imagedata")]
        public List<CT_ImageData> imagedata {
            get {
                return this.imagedataField;
            }
            set {
                this.imagedataField = value;
            }
        }
        
        
        [XmlElement("wrap", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Wrap> wrap {
            get {
                return this.wrapField;
            }
            set {
                this.wrapField = value;
            }
        }
        
        
        [XmlElement("anchorlock", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_AnchorLock> anchorlock {
            get {
                return this.anchorlockField;
            }
            set {
                this.anchorlockField = value;
            }
        }
        
        
        [XmlElement("bordertop", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Border> bordertop {
            get {
                return this.bordertopField;
            }
            set {
                this.bordertopField = value;
            }
        }
        
        
        [XmlElement("borderbottom", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Border> borderbottom {
            get {
                return this.borderbottomField;
            }
            set {
                this.borderbottomField = value;
            }
        }
        
        
        [XmlElement("borderleft", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Border> borderleft {
            get {
                return this.borderleftField;
            }
            set {
                this.borderleftField = value;
            }
        }
        
        
        [XmlElement("borderright", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Border> borderright {
            get {
                return this.borderrightField;
            }
            set {
                this.borderrightField = value;
            }
        }
        
        
        [XmlElement("ClientData", Namespace="urn:schemas-microsoft-com:office:excel")]
        public List<CT_ClientData> ClientData {
            get {
                return this.clientDataField;
            }
            set {
                this.clientDataField = value;
            }
        }
        
        
        [XmlElement("textdata", Namespace="urn:schemas-microsoft-com:office:powerpoint")]
        public List<CT_Rel> textdata {
            get {
                return this.textdataField;
            }
            set {
                this.textdataField = value;
            }
        }
        
        
        [XmlAttribute]
        public string from {
            get {
                return this.fromField;
            }
            set {
                this.fromField = value;
            }
        }
        
        
        [XmlAttribute]
        public string control1 {
            get {
                return this.control1Field;
            }
            set {
                this.control1Field = value;
            }
        }
        
        
        [XmlAttribute]
        public string control2 {
            get {
                return this.control2Field;
            }
            set {
                this.control2Field = value;
            }
        }
        
        
        [XmlAttribute]
        public string to {
            get {
                return this.toField;
            }
            set {
                this.toField = value;
            }
        }
    }
    
    
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_Image {
        
        private List<CT_Path> pathField;
        
        private List<CT_Formulas> formulasField;
        
        private List<CT_Handles> handlesField;
        
        private List<CT_Fill> fillField;
        
        private List<CT_Stroke> strokeField;
        
        private List<CT_Shadow> shadowField;
        
        private List<CT_Textbox> textboxField;
        
        private List<CT_TextPath> textpathField;
        
        private List<CT_ImageData> imagedataField;
        
        private List<CT_Wrap> wrapField;
        
        private List<CT_AnchorLock> anchorlockField;
        
        private List<CT_Border> bordertopField;
        
        private List<CT_Border> borderbottomField;
        
        private List<CT_Border> borderleftField;
        
        private List<CT_Border> borderrightField;
        
        private List<CT_ClientData> clientDataField;
        
        private List<CT_Rel> textdataField;
        
        private string srcField;
        
        private string cropleftField;
        
        private string croptopField;
        
        private string croprightField;
        
        private string cropbottomField;
        
        private string gainField;
        
        private string blacklevelField;
        
        private string gammaField;
        
        private ST_TrueFalse grayscaleField;
        
        private bool grayscaleFieldSpecified;
        
        private ST_TrueFalse bilevelField;
        
        private bool bilevelFieldSpecified;
        
        
        [XmlElement("path")]
        public List<CT_Path> path {
            get {
                return this.pathField;
            }
            set {
                this.pathField = value;
            }
        }
        
        
        [XmlElement("formulas")]
        public List<CT_Formulas> formulas {
            get {
                return this.formulasField;
            }
            set {
                this.formulasField = value;
            }
        }
        
        
        [XmlElement("handles")]
        public List<CT_Handles> handles {
            get {
                return this.handlesField;
            }
            set {
                this.handlesField = value;
            }
        }
        
        
        [XmlElement("fill")]
        public List<CT_Fill> fill {
            get {
                return this.fillField;
            }
            set {
                this.fillField = value;
            }
        }
        
        
        [XmlElement("stroke")]
        public List<CT_Stroke> stroke {
            get {
                return this.strokeField;
            }
            set {
                this.strokeField = value;
            }
        }
        
        
        [XmlElement("shadow")]
        public List<CT_Shadow> shadow {
            get {
                return this.shadowField;
            }
            set {
                this.shadowField = value;
            }
        }
        
        
        [XmlElement("textbox")]
        public List<CT_Textbox> textbox {
            get {
                return this.textboxField;
            }
            set {
                this.textboxField = value;
            }
        }
        
        
        [XmlElement("textpath")]
        public List<CT_TextPath> textpath {
            get {
                return this.textpathField;
            }
            set {
                this.textpathField = value;
            }
        }
        
        
        [XmlElement("imagedata")]
        public List<CT_ImageData> imagedata {
            get {
                return this.imagedataField;
            }
            set {
                this.imagedataField = value;
            }
        }
        
        
        [XmlElement("wrap", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Wrap> wrap {
            get {
                return this.wrapField;
            }
            set {
                this.wrapField = value;
            }
        }
        
        
        [XmlElement("anchorlock", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_AnchorLock> anchorlock {
            get {
                return this.anchorlockField;
            }
            set {
                this.anchorlockField = value;
            }
        }
        
        
        [XmlElement("bordertop", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Border> bordertop {
            get {
                return this.bordertopField;
            }
            set {
                this.bordertopField = value;
            }
        }
        
        
        [XmlElement("borderbottom", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Border> borderbottom {
            get {
                return this.borderbottomField;
            }
            set {
                this.borderbottomField = value;
            }
        }
        
        
        [XmlElement("borderleft", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Border> borderleft {
            get {
                return this.borderleftField;
            }
            set {
                this.borderleftField = value;
            }
        }
        
        
        [XmlElement("borderright", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Border> borderright {
            get {
                return this.borderrightField;
            }
            set {
                this.borderrightField = value;
            }
        }
        
        
        [XmlElement("ClientData", Namespace="urn:schemas-microsoft-com:office:excel")]
        public List<CT_ClientData> ClientData {
            get {
                return this.clientDataField;
            }
            set {
                this.clientDataField = value;
            }
        }
        
        
        [XmlElement("textdata", Namespace="urn:schemas-microsoft-com:office:powerpoint")]
        public List<CT_Rel> textdata {
            get {
                return this.textdataField;
            }
            set {
                this.textdataField = value;
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
        public string cropleft {
            get {
                return this.cropleftField;
            }
            set {
                this.cropleftField = value;
            }
        }
        
        
        [XmlAttribute]
        public string croptop {
            get {
                return this.croptopField;
            }
            set {
                this.croptopField = value;
            }
        }
        
        
        [XmlAttribute]
        public string cropright {
            get {
                return this.croprightField;
            }
            set {
                this.croprightField = value;
            }
        }
        
        
        [XmlAttribute]
        public string cropbottom {
            get {
                return this.cropbottomField;
            }
            set {
                this.cropbottomField = value;
            }
        }
        
        
        [XmlAttribute]
        public string gain {
            get {
                return this.gainField;
            }
            set {
                this.gainField = value;
            }
        }
        
        
        [XmlAttribute]
        public string blacklevel {
            get {
                return this.blacklevelField;
            }
            set {
                this.blacklevelField = value;
            }
        }
        
        
        [XmlAttribute]
        public string gamma {
            get {
                return this.gammaField;
            }
            set {
                this.gammaField = value;
            }
        }
        
        
        [XmlAttribute]
        public ST_TrueFalse grayscale {
            get {
                return this.grayscaleField;
            }
            set {
                this.grayscaleField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool grayscaleSpecified {
            get {
                return this.grayscaleFieldSpecified;
            }
            set {
                this.grayscaleFieldSpecified = value;
            }
        }
        
        
        [XmlAttribute]
        public ST_TrueFalse bilevel {
            get {
                return this.bilevelField;
            }
            set {
                this.bilevelField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool bilevelSpecified {
            get {
                return this.bilevelFieldSpecified;
            }
            set {
                this.bilevelFieldSpecified = value;
            }
        }
    }
    
    
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_Line {
        
        private List<CT_Path> pathField;
        
        private List<CT_Formulas> formulasField;
        
        private List<CT_Handles> handlesField;
        
        private List<CT_Fill> fillField;
        
        private List<CT_Stroke> strokeField;
        
        private List<CT_Shadow> shadowField;
        
        private List<CT_Textbox> textboxField;
        
        private List<CT_TextPath> textpathField;
        
        private List<CT_ImageData> imagedataField;
        
        private List<CT_Wrap> wrapField;
        
        private List<CT_AnchorLock> anchorlockField;
        
        private List<CT_Border> bordertopField;
        
        private List<CT_Border> borderbottomField;
        
        private List<CT_Border> borderleftField;
        
        private List<CT_Border> borderrightField;
        
        private List<CT_ClientData> clientDataField;
        
        private List<CT_Rel> textdataField;
        
        private string fromField;
        
        private string toField;
        
        
        [XmlElement("path")]
        public List<CT_Path> path {
            get {
                return this.pathField;
            }
            set {
                this.pathField = value;
            }
        }
        
        
        [XmlElement("formulas")]
        public List<CT_Formulas> formulas {
            get {
                return this.formulasField;
            }
            set {
                this.formulasField = value;
            }
        }
        
        
        [XmlElement("handles")]
        public List<CT_Handles> handles {
            get {
                return this.handlesField;
            }
            set {
                this.handlesField = value;
            }
        }
        
        
        [XmlElement("fill")]
        public List<CT_Fill> fill {
            get {
                return this.fillField;
            }
            set {
                this.fillField = value;
            }
        }
        
        
        [XmlElement("stroke")]
        public List<CT_Stroke> stroke {
            get {
                return this.strokeField;
            }
            set {
                this.strokeField = value;
            }
        }
        
        
        [XmlElement("shadow")]
        public List<CT_Shadow> shadow {
            get {
                return this.shadowField;
            }
            set {
                this.shadowField = value;
            }
        }
        
        
        [XmlElement("textbox")]
        public List<CT_Textbox> textbox {
            get {
                return this.textboxField;
            }
            set {
                this.textboxField = value;
            }
        }
        
        
        [XmlElement("textpath")]
        public List<CT_TextPath> textpath {
            get {
                return this.textpathField;
            }
            set {
                this.textpathField = value;
            }
        }
        
        
        [XmlElement("imagedata")]
        public List<CT_ImageData> imagedata {
            get {
                return this.imagedataField;
            }
            set {
                this.imagedataField = value;
            }
        }
        
        
        [XmlElement("wrap", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Wrap> wrap {
            get {
                return this.wrapField;
            }
            set {
                this.wrapField = value;
            }
        }
        
        
        [XmlElement("anchorlock", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_AnchorLock> anchorlock {
            get {
                return this.anchorlockField;
            }
            set {
                this.anchorlockField = value;
            }
        }
        
        
        [XmlElement("bordertop", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Border> bordertop {
            get {
                return this.bordertopField;
            }
            set {
                this.bordertopField = value;
            }
        }
        
        
        [XmlElement("borderbottom", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Border> borderbottom {
            get {
                return this.borderbottomField;
            }
            set {
                this.borderbottomField = value;
            }
        }
        
        
        [XmlElement("borderleft", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Border> borderleft {
            get {
                return this.borderleftField;
            }
            set {
                this.borderleftField = value;
            }
        }
        
        
        [XmlElement("borderright", Namespace="urn:schemas-microsoft-com:office:word")]
        public List<CT_Border> borderright {
            get {
                return this.borderrightField;
            }
            set {
                this.borderrightField = value;
            }
        }
        
        
        [XmlElement("ClientData", Namespace="urn:schemas-microsoft-com:office:excel")]
        public List<CT_ClientData> ClientData {
            get {
                return this.clientDataField;
            }
            set {
                this.clientDataField = value;
            }
        }
        
        
        [XmlElement("textdata", Namespace="urn:schemas-microsoft-com:office:powerpoint")]
        public List<CT_Rel> textdata {
            get {
                return this.textdataField;
            }
            set {
                this.textdataField = value;
            }
        }
        
        
        [XmlAttribute]
        public string from {
            get {
                return this.fromField;
            }
            set {
                this.fromField = value;
            }
        }
        
        
        [XmlAttribute]
        public string to {
            get {
                return this.toField;
            }
            set {
                this.toField = value;
            }
        }
    }
    
    
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_Oval {
        
        private List<object> itemsField;
        
        private ItemsChoiceType2[] itemsElementNameField;
        
        
        [XmlElement("ClientData", typeof(CT_ClientData), Namespace="urn:schemas-microsoft-com:office:excel")]
        [XmlElement("textdata", typeof(CT_Rel), Namespace="urn:schemas-microsoft-com:office:powerpoint")]
        [XmlElement("anchorlock", typeof(CT_AnchorLock), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderbottom", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderleft", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderright", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("bordertop", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("wrap", typeof(CT_Wrap), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("fill", typeof(CT_Fill))]
        [XmlElement("formulas", typeof(CT_Formulas))]
        [XmlElement("handles", typeof(CT_Handles))]
        [XmlElement("imagedata", typeof(CT_ImageData))]
        [XmlElement("path", typeof(CT_Path))]
        [XmlElement("shadow", typeof(CT_Shadow))]
        [XmlElement("stroke", typeof(CT_Stroke))]
        [XmlElement("textbox", typeof(CT_Textbox))]
        [XmlElement("textpath", typeof(CT_TextPath))]
        [System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemsElementName")]
        public List<object> Items {
            get {
                return this.itemsField;
            }
            set {
                this.itemsField = value;
            }
        }
        
        
        [XmlElement("ItemsElementName")]
        [System.Xml.Serialization.XmlIgnore]
        public ItemsChoiceType2[] ItemsElementName {
            get {
                return this.itemsElementNameField;
            }
            set {
                this.itemsElementNameField = value;
            }
        }
    }
    
    
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml", IncludeInSchema=false)]
    public enum ItemsChoiceType2 {
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:excel:ClientData")]
        ClientData,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:powerpoint:textdata")]
        textdata,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:anchorlock")]
        anchorlock,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:borderbottom")]
        borderbottom,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:borderleft")]
        borderleft,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:borderright")]
        borderright,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:bordertop")]
        bordertop,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:wrap")]
        wrap,
        
        
        fill,
        
        
        formulas,
        
        
        handles,
        
        
        imagedata,
        
        
        path,
        
        
        shadow,
        
        
        stroke,
        
        
        textbox,
        
        
        textpath,
    }
    
    
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_PolyLine {
        
        private List<object> itemsField;
        
        private ItemsChoiceType3[] itemsElementNameField;
        
        private string pointsField;
        
        
        [XmlElement("ClientData", typeof(CT_ClientData), Namespace="urn:schemas-microsoft-com:office:excel")]
        [XmlElement("textdata", typeof(CT_Rel), Namespace="urn:schemas-microsoft-com:office:powerpoint")]
        [XmlElement("anchorlock", typeof(CT_AnchorLock), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderbottom", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderleft", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderright", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("bordertop", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("wrap", typeof(CT_Wrap), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("fill", typeof(CT_Fill))]
        [XmlElement("formulas", typeof(CT_Formulas))]
        [XmlElement("handles", typeof(CT_Handles))]
        [XmlElement("imagedata", typeof(CT_ImageData))]
        [XmlElement("path", typeof(CT_Path))]
        [XmlElement("shadow", typeof(CT_Shadow))]
        [XmlElement("stroke", typeof(CT_Stroke))]
        [XmlElement("textbox", typeof(CT_Textbox))]
        [XmlElement("textpath", typeof(CT_TextPath))]
        [System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemsElementName")]
        public List<object> Items {
            get {
                return this.itemsField;
            }
            set {
                this.itemsField = value;
            }
        }
        
        
        [XmlElement("ItemsElementName")]
        [System.Xml.Serialization.XmlIgnore]
        public ItemsChoiceType3[] ItemsElementName {
            get {
                return this.itemsElementNameField;
            }
            set {
                this.itemsElementNameField = value;
            }
        }
        
        
        [XmlAttribute]
        public string points {
            get {
                return this.pointsField;
            }
            set {
                this.pointsField = value;
            }
        }
    }
    
    
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml", IncludeInSchema=false)]
    public enum ItemsChoiceType3 {
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:excel:ClientData")]
        ClientData,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:powerpoint:textdata")]
        textdata,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:anchorlock")]
        anchorlock,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:borderbottom")]
        borderbottom,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:borderleft")]
        borderleft,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:borderright")]
        borderright,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:bordertop")]
        bordertop,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:wrap")]
        wrap,
        
        
        fill,
        
        
        formulas,
        
        
        handles,
        
        
        imagedata,
        
        
        path,
        
        
        shadow,
        
        
        stroke,
        
        
        textbox,
        
        
        textpath,
    }
    
    
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_Rect {
        
        private List<object> itemsField;
        
        private ItemsChoiceType4[] itemsElementNameField;
        
        
        [XmlElement("ClientData", typeof(CT_ClientData), Namespace="urn:schemas-microsoft-com:office:excel")]
        [XmlElement("textdata", typeof(CT_Rel), Namespace="urn:schemas-microsoft-com:office:powerpoint")]
        [XmlElement("anchorlock", typeof(CT_AnchorLock), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderbottom", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderleft", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderright", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("bordertop", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("wrap", typeof(CT_Wrap), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("fill", typeof(CT_Fill))]
        [XmlElement("formulas", typeof(CT_Formulas))]
        [XmlElement("handles", typeof(CT_Handles))]
        [XmlElement("imagedata", typeof(CT_ImageData))]
        [XmlElement("path", typeof(CT_Path))]
        [XmlElement("shadow", typeof(CT_Shadow))]
        [XmlElement("stroke", typeof(CT_Stroke))]
        [XmlElement("textbox", typeof(CT_Textbox))]
        [XmlElement("textpath", typeof(CT_TextPath))]
        [System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemsElementName")]
        public List<object> Items {
            get {
                return this.itemsField;
            }
            set {
                this.itemsField = value;
            }
        }
        
        
        [XmlElement("ItemsElementName")]
        [System.Xml.Serialization.XmlIgnore]
        public ItemsChoiceType4[] ItemsElementName {
            get {
                return this.itemsElementNameField;
            }
            set {
                this.itemsElementNameField = value;
            }
        }
    }
    
    
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml", IncludeInSchema=false)]
    public enum ItemsChoiceType4 {
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:excel:ClientData")]
        ClientData,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:powerpoint:textdata")]
        textdata,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:anchorlock")]
        anchorlock,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:borderbottom")]
        borderbottom,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:borderleft")]
        borderleft,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:borderright")]
        borderright,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:bordertop")]
        bordertop,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:wrap")]
        wrap,
        
        
        fill,
        
        
        formulas,
        
        
        handles,
        
        
        imagedata,
        
        
        path,
        
        
        shadow,
        
        
        stroke,
        
        
        textbox,
        
        
        textpath,
    }
    
    
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_RoundRect {
        
        private List<object> itemsField;
        
        private ItemsChoiceType5[] itemsElementNameField;
        
        private string arcsizeField;
        
        
        [XmlElement("ClientData", typeof(CT_ClientData), Namespace="urn:schemas-microsoft-com:office:excel")]
        [XmlElement("textdata", typeof(CT_Rel), Namespace="urn:schemas-microsoft-com:office:powerpoint")]
        [XmlElement("anchorlock", typeof(CT_AnchorLock), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderbottom", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderleft", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderright", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("bordertop", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("wrap", typeof(CT_Wrap), Namespace="urn:schemas-microsoft-com:office:word")]
        [XmlElement("fill", typeof(CT_Fill))]
        [XmlElement("formulas", typeof(CT_Formulas))]
        [XmlElement("handles", typeof(CT_Handles))]
        [XmlElement("imagedata", typeof(CT_ImageData))]
        [XmlElement("path", typeof(CT_Path))]
        [XmlElement("shadow", typeof(CT_Shadow))]
        [XmlElement("stroke", typeof(CT_Stroke))]
        [XmlElement("textbox", typeof(CT_Textbox))]
        [XmlElement("textpath", typeof(CT_TextPath))]
        [System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemsElementName")]
        public List<object> Items {
            get {
                return this.itemsField;
            }
            set {
                this.itemsField = value;
            }
        }
        
        
        [XmlElement("ItemsElementName")]
        [System.Xml.Serialization.XmlIgnore]
        public ItemsChoiceType5[] ItemsElementName {
            get {
                return this.itemsElementNameField;
            }
            set {
                this.itemsElementNameField = value;
            }
        }
        
        
        [XmlAttribute]
        public string arcsize {
            get {
                return this.arcsizeField;
            }
            set {
                this.arcsizeField = value;
            }
        }
    }
    
    
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml", IncludeInSchema=false)]
    public enum ItemsChoiceType5 {
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:excel:ClientData")]
        ClientData,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:powerpoint:textdata")]
        textdata,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:anchorlock")]
        anchorlock,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:borderbottom")]
        borderbottom,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:borderleft")]
        borderleft,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:borderright")]
        borderright,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:bordertop")]
        bordertop,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:wrap")]
        wrap,
        
        
        fill,
        
        
        formulas,
        
        
        handles,
        
        
        imagedata,
        
        
        path,
        
        
        shadow,
        
        
        stroke,
        
        
        textbox,
        
        
        textpath,
    }
    
    
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml", IncludeInSchema=false)]
    public enum ItemsChoiceType6 {
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:excel:ClientData")]
        ClientData,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:powerpoint:textdata")]
        textdata,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:anchorlock")]
        anchorlock,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:borderbottom")]
        borderbottom,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:borderleft")]
        borderleft,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:borderright")]
        borderright,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:bordertop")]
        bordertop,
        
        
        [XmlEnum("urn:schemas-microsoft-com:office:word:wrap")]
        wrap,
        
        
        arc,
        
        
        curve,
        
        
        fill,
        
        
        formulas,
        
        
        group,
        
        
        handles,
        
        
        image,
        
        
        imagedata,
        
        
        line,
        
        
        oval,
        
        
        path,
        
        
        polyline,
        
        
        rect,
        
        
        roundrect,
        
        
        shadow,
        
        
        shape,
        
        
        shapetype,
        
        
        stroke,
        
        
        textbox,
        
        
        textpath,
    }
    
    
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=false)]
    public enum ST_EditAs {
        
        
        canvas,
        
        
        orgchart,
        
        
        radial,
        
        
        cycle,
        
        
        stacked,
        
        
        venn,
        
        
        bullseye,
    }
    
    
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_Background {
        
        private CT_Fill fillField;
        
        private string idField;
        
        private ST_TrueFalse filledField;
        
        private bool filledFieldSpecified;
        
        private string fillcolorField;
        
        
        public CT_Fill fill {
            get {
                return this.fillField;
            }
            set {
                this.fillField = value;
            }
        }
        
        
        // TODO is the following correct?
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
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
        public ST_TrueFalse filled {
            get {
                return this.filledField;
            }
            set {
                this.filledField = value;
            }
        }
        
        
        [System.Xml.Serialization.XmlIgnore]
        public bool filledSpecified {
            get {
                return this.filledFieldSpecified;
            }
            set {
                this.filledFieldSpecified = value;
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
    }
    
    
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=false)]
    public enum ST_Ext {
        
        
        view,
        
        
        edit,
        
        
        backwardCompatible,
    }
}
