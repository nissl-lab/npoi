using System;
using System.Collections.Generic;
using System.Xml.Schema;
using System.Xml.Serialization;
using NPOI.OpenXmlFormats.Vml.Wordprocessing;
using NPOI.OpenXmlFormats.Vml.Office;
using NPOI.OpenXmlFormats.Vml.Spreadsheet;
using NPOI.OpenXmlFormats.Vml.Presentation;
using System.IO;
using System.Xml;
using System.Text;

namespace NPOI.OpenXmlFormats.Vml
{

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot("shape",Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class  CT_Shape {
        
        
        private string typeField;
        
        private string adjField;
        private string styleField;
        private CT_Path pathField;
        
        private string equationxmlField;

        private string idField;
        private string fillcolorField;
        private ST_InsetMode insetmodeField;

        private ST_TrueFalse strokedField;
        private string wrapcoordsField;
        [XmlAttribute]
        public string wrapcoords
        {
            get { return wrapcoordsField; }
            set { wrapcoordsField = value; }
        }
        [XmlAttribute]
        public ST_TrueFalse stroked
        {
            get { return strokedField; }
            set { strokedField = value; }
        }


        static XmlSerializer serializer = new XmlSerializer(typeof(CT_Shape), "urn:schemas-microsoft-com:vml");
        public static CT_Shape Parse(string xmltext)
        {
            TextReader tr = new StringReader(xmltext);
            CT_Shape obj = (CT_Shape)serializer.Deserialize(tr);
            return obj;
        }
        private string spidField;

        [XmlAttribute(Namespace = "urn:schemas-microsoft-com:office:office")]
        public string spid
        {
            get { return this.spidField; }
            set { this.spidField = value; }
        }
        [XmlAttribute]
        public string id
        {
            get { return idField; }
            set { idField = value; }
        }

        [XmlAttribute]
        public string fillcolor
        {
            get { return fillcolorField; }
            set { fillcolorField = value; }
        }

        [XmlAttribute(Namespace = "urn:schemas-microsoft-com:office:office")]
        public ST_InsetMode insetmode
        {
            get { return insetmodeField; }
            set { insetmodeField = value; }
        }

        public CT_Textbox AddNewTextbox()
        {
            textboxField = new CT_Textbox();
            return this.textboxField;
            }

        private CT_Wrap wrapField;
        private CT_Fill fillField;
        private CT_Formulas formulasField;
        private CT_Handles handlesField;
        private CT_ImageData imagedataField;
        private CT_Stroke strokeField;
        private CT_Shadow shadowField;
        private CT_Textbox textboxField;
        private CT_TextPath textpathField;
        private CT_Empty iscommentField;

        /*[XmlElement("textdata", typeof(CT_Rel), Namespace = "urn:schemas-microsoft-com:office:powerpoint")]
        [XmlElement("anchorlock", typeof(CT_AnchorLock), Namespace = "urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderbottom", typeof(CT_Border), Namespace = "urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderleft", typeof(CT_Border), Namespace = "urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderright", typeof(CT_Border), Namespace = "urn:schemas-microsoft-com:office:word")]
        [XmlElement("bordertop", typeof(CT_Border), Namespace = "urn:schemas-microsoft-com:office:word")]*/

        [XmlElement(Namespace = "urn:schemas-microsoft-com:office:powerpoint")]
        public CT_Empty iscomment
        {
            get { return this.iscommentField; }
            set { this.iscommentField = value; }
        }

        [XmlElement(Namespace="urn:schemas-microsoft-com:office:word")]
        public CT_Wrap wrap
        {
            get { return this.wrapField; }
            set { this.wrapField = value; }
        }
        [XmlElement]
        public CT_Textbox textbox
        {
            get { return this.textboxField; }
            set { this.textboxField = value; }
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
        public CT_Path AddNewPath()
        {
            this.pathField = new CT_Path();
            return this.pathField;       
        }
        List<CT_ClientData> clientDataField = null;

        [XmlElement("ClientData",Namespace = "urn:schemas-microsoft-com:office:excel")]
        public CT_ClientData[] ClientData
        {
            get
            {
                if (clientDataField == null)
                    return null;
                return clientDataField.ToArray();
            }
            set
            {
                if (value == null)
                    this.clientDataField = new List<CT_ClientData>();
                else
                    this.clientDataField = new List<CT_ClientData>(value);
            }
        }
        public CT_ClientData GetClientDataArray(int index)
        {
            return clientDataField != null ? this.clientDataField[index] : null;
        }
        public int sizeOfClientDataArray()
        {
            if (clientDataField == null)
                return 0;
            return clientDataField.Count;
        }
        public CT_ClientData AddNewClientData()
        {
            CT_ClientData cd=new CT_ClientData();
            if (clientDataField == null)
                this.clientDataField = new List<CT_ClientData>();
            this.clientDataField.Add(cd);
            return cd;
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
        internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
            new XmlQualifiedName("o", "urn:schemas-microsoft-com:office:office"),
            new XmlQualifiedName("x", "urn:schemas-microsoft-com:office:excel"),
            new XmlQualifiedName("v", "urn:schemas-microsoft-com:vml")
        });

        public override string ToString()
        {
            using (StringWriter stringWriter = new StringWriter())
            {
                XmlWriterSettings settings = new XmlWriterSettings();
                
                settings.Encoding = Encoding.UTF8;
                settings.OmitXmlDeclaration = true;

                using (XmlWriter writer = XmlWriter.Create(stringWriter, settings))
                {
                    serializer.Serialize(writer, this, namespaces);
                }
                return stringWriter.ToString();
            }
        }
        [XmlElement]
        public CT_TextPath textpath
        {
            get
            {
                return this.textpathField;
            }
            set
            {
                this.textpathField = value;
            }
        }

        public CT_TextPath AddNewTextpath()
        {
            this.textpathField = new CT_TextPath();
            return this.textpathField;
        }
    }
    
    
    [Serializable]
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

        public CT_F AddNewF()
        {
            if (this.fField == null)
                this.fField = new List<CT_F>();
            this.fField.Add(new CT_F());
            return this.fField[this.fField.Count - 1];
        }
    }


    [Serializable]
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


    [Serializable]
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

        public CT_H AddNewH()
        {
            if (hField == null)
                hField = new List<CT_H>();
            CT_H h = new CT_H();
            hField.Add(h);
            return h;
        }
    }
    
    
    [Serializable]
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
        
        
        [XmlIgnore]
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
        
        
        [XmlIgnore]
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
        
        
        [XmlIgnore]
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
    
    
    [Serializable]
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
        
        
        [XmlIgnore]
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
        
        
        [XmlIgnore]
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
        
        
        [XmlAttribute]//(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string pict {
            get {
                return this.pictField;
            }
            set {
                this.pictField = value;
            }
        }
        
        
        [XmlAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string href {
            get {
                return this.hrefField;
            }
            set {
                this.hrefField = value;
            }
        }
    }
    
    
    [Serializable]
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

        private ST_ConnectType connecttypeField;

        private string connectlocsField;

        private bool connectlocsFieldSpecified;

        private string connectanglesField;

        private bool connectanglesFieldSpecified;

        private ST_TrueFalse extrusionokField;

        private bool extrusionokFieldSpecified;

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
        [XmlAttribute(Namespace = "urn:schemas-microsoft-com:office:office")]
        public ST_ConnectType connecttype
        {
            get
            {
                return this.connecttypeField;
            }
            set
            {
                this.connecttypeField = value;
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
        
        
        [XmlIgnore]
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
        
        
        [XmlIgnore]
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
        
        
        [XmlIgnore]
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
        
        
        [XmlIgnore]
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
        
        
        [XmlIgnore]
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
        
        
        [XmlIgnore]
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
        
        
        [XmlIgnore]
        public bool insetpenokSpecified {
            get {
                return this.insetpenokFieldSpecified;
            }
            set {
                this.insetpenokFieldSpecified = value;
            }
        }
        [XmlAttribute]
        public string connectlocs
        {
            get
            {
                return this.connectlocsField;
            }
            set
            {
                this.connectlocsField = value;
            }
        }
        [XmlIgnore]
        public bool connectlocsSpecified
        {
            get
            {
                return this.connectlocsFieldSpecified;
            }
            set
            {
                this.connectlocsFieldSpecified = value;
            }
        }
        [XmlAttribute]
        public string connectangles
        {
            get
            {
                return this.connectanglesField;
            }
            set
            {
                this.connectanglesField = value;
            }
        }
        [XmlIgnore]
        public bool connectanglesSpecified
        {
            get
            {
                return this.connectanglesFieldSpecified;
            }
            set
            {
                this.connectanglesFieldSpecified = value;
            }
        }
        [XmlAttribute]
        public ST_TrueFalse extrusionok
        {
            get
            {
                return this.extrusionokField;
            }
            set
            {
                this.extrusionokField = value;
            }
        }
        [XmlIgnore]
        public bool extrusionokSpecified
        {
            get
            {
                return this.extrusionokFieldSpecified;
            }
            set
            {
                this.extrusionokFieldSpecified = value;
            }
        }
    }
    
    
    [Serializable]
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
        public ST_ShadowType type {
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
        
        
        [XmlAttribute]
        public ST_TrueFalse obscured {
            get {
                return this.obscuredField;
            }
            set {
                this.obscuredField = value;
            }
        }
        
        
        [XmlIgnore]
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
    
    
    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=false)]
    public enum ST_ShadowType {
        
        
        single,
        
        
        @double,
        
        
        emboss,
        
        
        perspective,
    }
    
    
    [Serializable]
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
    }
    
    
    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=false)]
    public enum ST_StrokeLineStyle {
        
        
        single,
        
        
        thinThin,
        
        
        thinThick,
        
        
        thickThin,
        
        
        thickBetweenThin,
    }
    
    
    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=false)]
    public enum ST_StrokeJoinStyle {
        
        
        round,
        
        
        bevel,
        
        
        miter,
    }
    
    
    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=false)]
    public enum ST_StrokeEndCap {
        
        
        flat,
        
        
        square,
        
        
        round,
    }
    
    
    [Serializable]
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
    
    
    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=false)]
    public enum ST_StrokeArrowWidth {
        
        
        narrow,
        
        
        medium,
        
        
        wide,
    }
    
    
    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=false)]
    public enum ST_StrokeArrowLength {
        
        
        @short,
        
        
        medium,
        
        
        @long,
    }
    
    
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_Textbox {
        
        private System.Xml.XmlElement itemField;
        
        private string idField;
        
        private string styleField;
        
        private string insetField;
        
        
        [XmlAnyElement()]
        public System.Xml.XmlElement Item {
            get {
                return this.itemField;
            }
            set {
                this.itemField = value;
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
    
    
    [Serializable]
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
        public ST_TrueFalse fitshape {
            get {
                return this.fitshapeField;
            }
            set {
                this.fitshapeField = value;
            }
        }
        
        
        [XmlIgnore]
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
        
        
        [XmlIgnore]
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
        
        
        [XmlIgnore]
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
        
        
        [XmlIgnore]
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
    
    
    [Serializable]
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
    
    
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot("shapetype",Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_Shapetype {
        
        private CT_Path pathField;
        
        private List<CT_Formulas> formulasField = new List<CT_Formulas>();
        
        private List<CT_Handles> handlesField = new List<CT_Handles>();
        
        private List<CT_Fill> fillField = new List<CT_Fill>();
        
        private CT_Stroke strokeField;
        
        private List<CT_Shadow> shadowField;
        
        private List<CT_Textbox> textboxField;
        
        private List<CT_TextPath> textpathField = new List<CT_TextPath>();
        
        private List<CT_ImageData> imagedataField;
        
        private List<CT_Wrap> wrapField;
        
        private List<CT_AnchorLock> anchorlockField;

        private List<CT_Lock> lockField = new List<CT_Lock>();
        
        private List<CT_Border> bordertopField;
        
        private List<CT_Border> borderbottomField;
        
        private List<CT_Border> borderleftField;
        
        private List<CT_Border> borderrightField;
        
        private List<CT_ClientData> clientDataField;
        
        private List<CT_Rel> textdataField;
        
        private string adjField;
        private string idField;
        private string styleField;
        private float sptField;
        private string coordsizeField;

        static XmlSerializer serializer = new XmlSerializer(typeof(CT_Shapetype), "urn:schemas-microsoft-com:vml");

        public static CT_Shapetype Parse(string xmltext)
        {
            TextReader tr = new StringReader(xmltext);
            CT_Shapetype obj = (CT_Shapetype)serializer.Deserialize(tr);
            return obj;
        }

        [XmlAttribute(Namespace = "urn:schemas-microsoft-com:office:office")]
        public string coordsize
        {
            get
            {
                return this.coordsizeField;
            }
            set
            {
                this.coordsizeField = value;
            }
        }

        [XmlAttribute(Namespace = "urn:schemas-microsoft-com:office:office")]
        public float spt
        {
            get
            {
                return this.sptField;
            }
            set
            {
                this.sptField = value;
            }
        }
        private string path1Field;

        [XmlAttribute]
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
        public CT_Path path {
            get {
                return this.pathField;
            }
            set {

                this.pathField = value;
            }
        }


        [XmlElement("formulas")]
        public CT_Formulas[] formulas
        {
            get
            {
                return this.formulasField.ToArray();
            }
            set
            {
                if (value == null)
                    this.formulasField = new List<CT_Formulas>();
                else
                    this.formulasField = new List<CT_Formulas>(value);
            }
        }


        [XmlElement("handles")]
        public CT_Handles[] handles
        {
            get
            {
                return this.handlesField.ToArray();
            }
            set
            {
                if (value == null)
                    this.handlesField = new List<CT_Handles>();
                else
                    this.handlesField = new List<CT_Handles>(value);
            }
        }


        [XmlElement("fill")]
        public CT_Fill[] fill
        {
            get
            {
                return this.fillField.ToArray();
            }
            set
            {
                if (value == null)
                    this.fillField = new List<CT_Fill>();
                else
                    this.fillField = new List<CT_Fill>(value);
            }
        }
        
        
        [XmlElement("stroke")]
        public CT_Stroke stroke {
            get {
                return this.strokeField;
            }
            set {
                this.strokeField = value;
            }
        }
        
        
        //[XmlElement("shadow")]
        //public List<CT_Shadow> shadow {
        //    get {
        //        return this.shadowField;
        //    }
        //    set {
        //        this.shadowField = value;
        //    }
        //}
        
        
        //[XmlElement("textbox")]
        //public List<CT_Textbox> textbox {
        //    get {
        //        return this.textboxField;
        //    }
        //    set {
        //        this.textboxField = value;
        //    }
        //}


        [XmlElement("textpath")]
        public CT_TextPath[] textpath
        {
            get
            {
                return this.textpathField.ToArray();
            }
            set
            {
                if (value == null)
                    this.textpathField = new List<CT_TextPath>();
                else
                    this.textpathField = new List<CT_TextPath>(value);
            }
        }
        
        
        //[XmlElement("imagedata")]
        //public List<CT_ImageData> imagedata {
        //    get {
        //        return this.imagedataField;
        //    }
        //    set {
        //        this.imagedataField = value;
        //    }
        //}
        
        
        //[XmlElement("wrap", Namespace="urn:schemas-microsoft-com:office:word")]
        //public List<CT_Wrap> wrap {
        //    get {
        //        return this.wrapField;
        //    }
        //    set {
        //        this.wrapField = value;
        //    }
        //}
        
        
        //[XmlElement("anchorlock", Namespace="urn:schemas-microsoft-com:office:word")]
        //public List<CT_AnchorLock> anchorlock {
        //    get {
        //        return this.anchorlockField;
        //    }
        //    set {
        //        this.anchorlockField = value;
        //    }
        //}

        [XmlElement("lock", Namespace = "urn:schemas-microsoft-com:office:word")]
        public CT_Lock[] @lock
        {
            get
            {
                return this.lockField.ToArray();
            }
            set
            {
                if (value == null)
                    this.lockField = new List<CT_Lock>();
                else
                    this.lockField = new List<CT_Lock>(value);
            }
        }
        
        //[XmlElement("bordertop", Namespace="urn:schemas-microsoft-com:office:word")]
        //public List<CT_Border> bordertop {
        //    get {
        //        return this.bordertopField;
        //    }
        //    set {
        //        this.bordertopField = value;
        //    }
        //}
        
        
        //[XmlElement("borderbottom", Namespace="urn:schemas-microsoft-com:office:word")]
        //public List<CT_Border> borderbottom {
        //    get {
        //        return this.borderbottomField;
        //    }
        //    set {
        //        this.borderbottomField = value;
        //    }
        //}
        
        
        //[XmlElement("borderleft", Namespace="urn:schemas-microsoft-com:office:word")]
        //public List<CT_Border> borderleft {
        //    get {
        //        return this.borderleftField;
        //    }
        //    set {
        //        this.borderleftField = value;
        //    }
        //}
        
        
        //[XmlElement("borderright", Namespace="urn:schemas-microsoft-com:office:word")]
        //public List<CT_Border> borderright {
        //    get {
        //        return this.borderrightField;
        //    }
        //    set {
        //        this.borderrightField = value;
        //    }
        //}
        
        
        //[XmlElement("ClientData", Namespace="urn:schemas-microsoft-com:office:excel")]
        //public List<CT_ClientData> ClientData {
        //    get {
        //        return this.clientDataField;
        //    }
        //    set {
        //        this.clientDataField = value;
        //    }
        //}
        
        
        //[XmlElement("textdata", Namespace="urn:schemas-microsoft-com:office:powerpoint")]
        //public List<CT_Rel> textdata {
        //    get {
        //        return this.textdataField;
        //    }
        //    set {
        //        this.textdataField = value;
        //    }
        //}
        
        
        [XmlAttribute]
        public string adj {
            get {
                return this.adjField;
            }
            set {
                this.adjField = value;
            }
        }
        
        
        [XmlAttribute("path")]
        public string path2 {
            get {
                return this.path1Field;
            }
            set {
                this.path1Field = value;
            }
        }

        public CT_Stroke AddNewStroke()
        {
            this.strokeField = new CT_Stroke();
            return strokeField;
        }
        public CT_Path AddNewPath()
        {
                this.pathField = new CT_Path();
            return this.pathField;
        }
        internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
            new XmlQualifiedName("o", "urn:schemas-microsoft-com:office:office"),
            new XmlQualifiedName("x", "urn:schemas-microsoft-com:office:excel"),
            new XmlQualifiedName("v", "urn:schemas-microsoft-com:vml")
        });

        public override string ToString()
        {
            using (StringWriter stringWriter = new StringWriter())
            {
                XmlWriterSettings settings = new XmlWriterSettings();
                settings.Encoding = Encoding.UTF8;
                settings.OmitXmlDeclaration = true;

                using (XmlWriter writer = XmlWriter.Create(stringWriter, settings))
                {
                    serializer.Serialize(writer, this, namespaces);
                }
                return stringWriter.ToString();
            }
        }

        public CT_Formulas AddNewFormulas()
        {
            if (this.formulasField == null)
                this.formulasField = new List<CT_Formulas>();
            CT_Formulas obj = new CT_Formulas();
            this.formulasField.Add(obj);
            return obj;
        }

        public CT_TextPath AddNewTextpath()
        {
            if (this.textpathField == null)
                this.textpathField = new List<CT_TextPath>();
            CT_TextPath obj = new CT_TextPath();
            this.textpathField.Add(obj);
            return obj;
        }

        public CT_Handles AddNewHandles()
        {
            if (this.handlesField == null)
                this.handlesField = new List<CT_Handles>();
            CT_Handles obj = new CT_Handles();
            this.handlesField.Add(obj);
            return obj;
        }

        public CT_Lock AddNewLock()
        {
            if (this.lockField == null)
                this.lockField = new List<CT_Lock>();
            CT_Lock obj = new CT_Lock();
            this.lockField.Add(obj);
            return obj;
        }
    }
    
    
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_Group {
        
        private List<object> itemsField = new List<object>();
        
        private List<ItemsChoiceType6> itemsElementNameField = new List<ItemsChoiceType6>();
        
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
        [XmlChoiceIdentifier("ItemsElementName")]
        public object[] Items {
            get {
                if (this.itemsField == null)
                    return null;
                return this.itemsField.ToArray();
            }
            set {
                if (value == null)
                    this.itemsField = new List<object>();
                else
                    this.itemsField = new List<object>(value);
            }
        }
        
        
        [XmlElement("ItemsElementName")]
        [XmlIgnore]
        public ItemsChoiceType6[] ItemsElementName {
            get {
                return this.itemsElementNameField.ToArray();
            }
            set {
                if (value == null)
                    this.itemsElementNameField = new List<ItemsChoiceType6>();
                else
                    this.itemsElementNameField = new List<ItemsChoiceType6>(value);
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
        
        
        [XmlIgnore]
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
        
        
        [XmlIgnore]
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
            return AddNewObject<CT_Shapetype>(ItemsChoiceType6.shapetype);
        }
        #region Generic methods for object operation

        private List<T> GetObjectList<T>(ItemsChoiceType6 type) where T : class
        {
            lock (this)
            {
                List<T> list = new List<T>();
                for (int i = 0; i < itemsElementNameField.Count; i++)
                {
                    if (itemsElementNameField[i] == type)
                        list.Add(itemsField[i] as T);
                }
                return list;
            }
        }
        private int SizeOfObjectArray(ItemsChoiceType6 type)
        {
            lock (this)
            {
                int size = 0;
                for (int i = 0; i < itemsElementNameField.Count; i++)
                {
                    if (itemsElementNameField[i] == type)
                        size++;
                }
                return size;
            }
        }
        private T GetObjectArray<T>(int p, ItemsChoiceType6 type) where T : class
        {
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                if (pos < 0 || pos >= this.itemsField.Count)
                    return null;
                return itemsField[pos] as T;
            }
        }
        private T InsertNewObject<T>(ItemsChoiceType6 type, int p) where T : class, new()
        {
            T t = new T();
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                this.itemsElementNameField.Insert(pos, type);
                this.itemsField.Insert(pos, t);
            }
            return t;
        }
        private T AddNewObject<T>(ItemsChoiceType6 type) where T : class, new()
        {
            T t = new T();
            lock (this)
            {
                this.itemsElementNameField.Add(type);
                this.itemsField.Add(t);
            }
            return t;
        }
        private void SetObjectArray<T>(ItemsChoiceType6 type, int p, T obj) where T : class
        {
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                if (pos < 0 || pos >= this.itemsField.Count)
                    return;
                if (this.itemsField[pos] is T)
                    this.itemsField[pos] = obj;
                else
                    throw new Exception(string.Format(@"object types are difference, itemsField[{0}] is {1}, and parameter obj is {2}",
                        pos, this.itemsField[pos].GetType().Name, typeof(T).Name));
            }
        }
        private int GetObjectIndex(ItemsChoiceType6 type, int p)
        {
            int index = -1;
            int pos = 0;
            for (int i = 0; i < itemsElementNameField.Count; i++)
            {
                if (itemsElementNameField[i] == type)
                {
                    if (pos == p)
                    {
                        //return itemsField[p] as T;
                        index = i;
                        break;
                    }
                    else
                        pos++;
                }
            }
            return index;
        }
        private void RemoveObject(ItemsChoiceType6 type, int p)
        {
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                if (pos < 0 || pos >= this.itemsField.Count)
                    return;
                itemsElementNameField.RemoveAt(pos);
                itemsField.RemoveAt(pos);
            }
        }
        #endregion

        public CT_Shape AddNewShape()
        {
            return AddNewObject<CT_Shape>(ItemsChoiceType6.shape);
        }
    }
    
    
    [Serializable]
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
        
        
        [XmlIgnore]
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
        
        
        [XmlIgnore]
        public bool endAngleSpecified {
            get {
                return this.endAngleFieldSpecified;
            }
            set {
                this.endAngleFieldSpecified = value;
            }
        }
    }
    
    
    [Serializable]
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
    
    
    [Serializable]
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
        
        
        [XmlIgnore]
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
        
        
        [XmlIgnore]
        public bool bilevelSpecified {
            get {
                return this.bilevelFieldSpecified;
            }
            set {
                this.bilevelFieldSpecified = value;
            }
        }
    }
    
    
    [Serializable]
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
    
    
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_Oval {
        
        private List<object> itemsField = new List<object>();
        
        private List<ItemsChoiceType2> itemsElementNameField = new List<ItemsChoiceType2>();


        [XmlElement("ClientData", typeof(CT_ClientData), Namespace = "urn:schemas-microsoft-com:office:excel")]
        [XmlElement("textdata", typeof(CT_Rel), Namespace = "urn:schemas-microsoft-com:office:powerpoint")]
        [XmlElement("anchorlock", typeof(CT_AnchorLock), Namespace = "urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderbottom", typeof(CT_Border), Namespace = "urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderleft", typeof(CT_Border), Namespace = "urn:schemas-microsoft-com:office:word")]
        [XmlElement("borderright", typeof(CT_Border), Namespace = "urn:schemas-microsoft-com:office:word")]
        [XmlElement("bordertop", typeof(CT_Border), Namespace = "urn:schemas-microsoft-com:office:word")]
        [XmlElement("wrap", typeof(CT_Wrap), Namespace = "urn:schemas-microsoft-com:office:word")]
        [XmlElement("fill", typeof(CT_Fill))]
        [XmlElement("formulas", typeof(CT_Formulas))]
        [XmlElement("handles", typeof(CT_Handles))]
        [XmlElement("imagedata", typeof(CT_ImageData))]
        [XmlElement("path", typeof(CT_Path))]
        [XmlElement("shadow", typeof(CT_Shadow))]
        [XmlElement("stroke", typeof(CT_Stroke))]
        [XmlElement("textbox", typeof(CT_Textbox))]
        [XmlElement("textpath", typeof(CT_TextPath))]
        [XmlChoiceIdentifier("ItemsElementName")]
        public object[] Items
        {
            get
            {
                return this.itemsField.ToArray();
            }
            set
            {
                if (value == null)
                    this.itemsField = new List<object>();
                else
                    this.itemsField = new List<object>(value);
            }
        }
        
        
        [XmlElement("ItemsElementName")]
        [XmlIgnore]
        public ItemsChoiceType2[] ItemsElementName {
            get {
                return this.itemsElementNameField.ToArray();
            }
            set {
                if (value == null)
                    this.itemsElementNameField = new List<ItemsChoiceType2>();
                else
                    this.itemsElementNameField = new List<ItemsChoiceType2>(value);
            }
        }
    }
    
    
    [Serializable]
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
    
    
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_PolyLine {
        
        private List<object> itemsField = new List<object>();

        private List<ItemsChoiceType3> itemsElementNameField = new List<ItemsChoiceType3>();
        
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
        [XmlChoiceIdentifier("ItemsElementName")]
        public object[] Items {
            get {
                return this.itemsField.ToArray();
            }
            set {
                if (value == null)
                    this.itemsField = new List<object>();
                else
                    this.itemsField = new List<object>(value);
            }
        }
        
        
        [XmlElement("ItemsElementName")]
        [XmlIgnore]
        public ItemsChoiceType3[] ItemsElementName {
            get {
                return this.itemsElementNameField.ToArray();
            }
            set {
                if (value == null)
                    this.itemsElementNameField = new List<ItemsChoiceType3>();
                else
                    this.itemsElementNameField = new List<ItemsChoiceType3>(value);
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
    
    
    [Serializable]
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
    
    
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_Rect {
        
        private List<object> itemsField;
        
        private ItemsChoiceType4[] itemsElementNameField;
        
        
        //[XmlElement("ClientData", typeof(CT_ClientData), Namespace="urn:schemas-microsoft-com:office:excel")]
        //[XmlElement("textdata", typeof(CT_Rel), Namespace="urn:schemas-microsoft-com:office:powerpoint")]
        //[XmlElement("anchorlock", typeof(CT_AnchorLock), Namespace="urn:schemas-microsoft-com:office:word")]
        //[XmlElement("borderbottom", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        //[XmlElement("borderleft", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        //[XmlElement("borderright", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        //[XmlElement("bordertop", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        //[XmlElement("wrap", typeof(CT_Wrap), Namespace="urn:schemas-microsoft-com:office:word")]
        //[XmlElement("fill", typeof(CT_Fill))]
        //[XmlElement("formulas", typeof(CT_Formulas))]
        //[XmlElement("handles", typeof(CT_Handles))]
        //[XmlElement("imagedata", typeof(CT_ImageData))]
        //[XmlElement("path", typeof(CT_Path))]
        //[XmlElement("shadow", typeof(CT_Shadow))]
        //[XmlElement("stroke", typeof(CT_Stroke))]
        //[XmlElement("textbox", typeof(CT_Textbox))]
        //[XmlElement("textpath", typeof(CT_TextPath))]
        //[XmlChoiceIdentifier("ItemsElementName")]
        //public List<object> Items {
        //    get {
        //        return this.itemsField;
        //    }
        //    set {
        //        this.itemsField = value;
        //    }
        //}
        
        
        //[XmlElement("ItemsElementName")]
        //[XmlIgnore]
        //public ItemsChoiceType4[] ItemsElementName {
        //    get {
        //        return this.itemsElementNameField;
        //    }
        //    set {
        //        this.itemsElementNameField = value;
        //    }
        //}
    }
    
    
    [Serializable]
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
    
    
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=true)]
    public partial class CT_RoundRect {
        
        private List<object> itemsField;
        
        private ItemsChoiceType5[] itemsElementNameField;
        
        private string arcsizeField;
        
        
        //[XmlElement("ClientData", typeof(CT_ClientData), Namespace="urn:schemas-microsoft-com:office:excel")]
        //[XmlElement("textdata", typeof(CT_Rel), Namespace="urn:schemas-microsoft-com:office:powerpoint")]
        //[XmlElement("anchorlock", typeof(CT_AnchorLock), Namespace="urn:schemas-microsoft-com:office:word")]
        //[XmlElement("borderbottom", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        //[XmlElement("borderleft", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        //[XmlElement("borderright", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        //[XmlElement("bordertop", typeof(CT_Border), Namespace="urn:schemas-microsoft-com:office:word")]
        //[XmlElement("wrap", typeof(CT_Wrap), Namespace="urn:schemas-microsoft-com:office:word")]
        //[XmlElement("fill", typeof(CT_Fill))]
        //[XmlElement("formulas", typeof(CT_Formulas))]
        //[XmlElement("handles", typeof(CT_Handles))]
        //[XmlElement("imagedata", typeof(CT_ImageData))]
        //[XmlElement("path", typeof(CT_Path))]
        //[XmlElement("shadow", typeof(CT_Shadow))]
        //[XmlElement("stroke", typeof(CT_Stroke))]
        //[XmlElement("textbox", typeof(CT_Textbox))]
        //[XmlElement("textpath", typeof(CT_TextPath))]
        //[XmlChoiceIdentifier("ItemsElementName")]
        //public List<object> Items {
        //    get {
        //        return this.itemsField;
        //    }
        //    set {
        //        this.itemsField = value;
        //    }
        //}
        
        
        //[XmlElement("ItemsElementName")]
        //[XmlIgnore]
        //public ItemsChoiceType5[] ItemsElementName {
        //    get {
        //        return this.itemsElementNameField;
        //    }
        //    set {
        //        this.itemsElementNameField = value;
        //    }
        //}
        
        
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
    
    
    [Serializable]
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
    
    
    [Serializable]
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
    
    
    [Serializable]
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
    
    
    [Serializable]
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
        public ST_TrueFalse filled {
            get {
                return this.filledField;
            }
            set {
                this.filledField = value;
            }
        }
        
        
        [XmlIgnore]
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
    
    
    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:vml")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:vml", IsNullable=false)]
    public enum ST_Ext 
    {       
        NONE,
        view,
        edit,
        backwardCompatible,
    }
}
