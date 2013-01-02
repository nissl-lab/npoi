
using System;
using System.Xml.Serialization;
using System.Collections.Generic;
using System.Xml;

namespace NPOI.OpenXmlFormats.Dml
{
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_GraphicalObjectData
    {

        private List<XmlElement> anyField = new List<XmlElement>();

        private string uriField;
        
        public void AddPicElement(XmlElement el)
        {
            anyField.Add(el);
        }
        [XmlAnyElement()]
        public List<XmlElement> Any
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
    public class CT_GraphicalObject
    {
        public CT_GraphicalObjectData AddNewGraphicData()
        {
            this.graphicDataField = new CT_GraphicalObjectData();
            return this.graphicDataField;
        }
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
