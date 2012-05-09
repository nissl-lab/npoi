namespace NPOI.OpenXmlFormats.Dml {
    
    

    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/compatibility")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/compatibility", IsNullable=true)]
    public partial class CT_Compat {
        
        private string spidField;
        
    
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
