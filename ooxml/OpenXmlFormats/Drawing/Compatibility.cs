
using System;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Dml {
    
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/compatibility")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/compatibility", IsNullable=true)]
    public partial class CT_Compat {
        
        private string spidField;
        
    
        [XmlAttribute(DataType="token")]
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
