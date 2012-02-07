namespace NPOI.OpenXmlFormats.Dml {
    
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/compatibility")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/drawingml/2006/compatibility", IsNullable=true)]
    public partial class CT_Compat {
        
        private string spidField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="token")]
        public string spid {
            get {
                return this.spidField;
            }
            set {
                this.spidField = value;
            }
        }
    }
}
