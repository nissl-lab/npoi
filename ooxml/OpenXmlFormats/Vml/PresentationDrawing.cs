using System;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Vml.Presentation
{
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:powerpoint")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:powerpoint", IsNullable=true)]
    public partial class CT_Empty {
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:powerpoint")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:powerpoint", IsNullable=true)]
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
