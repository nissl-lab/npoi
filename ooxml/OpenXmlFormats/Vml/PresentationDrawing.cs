using System;
using System.Xml.Serialization;
namespace NPOI.OpenXmlFormats.Vml
{
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="urn:schemas-microsoft-com:office:powerpoint")]
    [System.Xml.Serialization.XmlRoot(Namespace="urn:schemas-microsoft-com:office:powerpoint", IsNullable=true)]
    public partial class CT_Empty {
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="urn:schemas-microsoft-com:office:powerpoint")]
    [System.Xml.Serialization.XmlRoot(Namespace="urn:schemas-microsoft-com:office:powerpoint", IsNullable=true)]
    public partial class CT_Rel {
        
        private string idField;
        
    
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
    }
}
