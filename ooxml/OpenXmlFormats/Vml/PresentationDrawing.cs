using System;
namespace NPOI.OpenXmlFormats.Vml
{
    
    /// <remarks/>
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="urn:schemas-microsoft-com:office:powerpoint")]
    [System.Xml.Serialization.XmlRoot(Namespace="urn:schemas-microsoft-com:office:powerpoint", IsNullable=true)]
    public partial class CT_Empty {
    }
    
    /// <remarks/>
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="urn:schemas-microsoft-com:office:powerpoint")]
    [System.Xml.Serialization.XmlRoot(Namespace="urn:schemas-microsoft-com:office:powerpoint", IsNullable=true)]
    public partial class CT_Rel {
        
        private string idField;
        
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
    }
}
