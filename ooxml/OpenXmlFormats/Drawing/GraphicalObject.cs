
using System;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Dml
{
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_GraphicalObjectData
    {

        private System.Xml.XmlElement[] anyField;

        private string uriField;


        [XmlAnyElement()]
        public System.Xml.XmlElement[] Any
        {
            get
            {
                return this.anyField;
            }
            set
            {
                this.anyField = value;
            }
        }


        [XmlAttribute(DataType = "token")]
        public string uri
        {
            get
            {
                return this.uriField;
            }
            set
            {
                this.uriField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_GraphicalObject
    {

        private CT_GraphicalObjectData graphicDataField;

        [XmlElement(Order = 0)]
        public CT_GraphicalObjectData graphicData
        {
            get
            {
                return this.graphicDataField;
            }
            set
            {
                this.graphicDataField = value;
            }
        }
    }
}
