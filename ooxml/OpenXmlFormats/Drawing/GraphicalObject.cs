
using System;
using System.Xml.Serialization;
using System.Collections.Generic;

namespace NPOI.OpenXmlFormats.Dml
{
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_GraphicalObjectData
    {

        private List<System.Xml.XmlElement> anyField = new List<System.Xml.XmlElement>();

        private string uriField;
        
        public void AddPicElement(System.Xml.XmlElement el)
        {
            anyField.Add(el);
        }
        [XmlAnyElement()]
        public System.Xml.XmlElement[] Any
        {
            get
            {
                return this.anyField.ToArray();
            }
            set
            {
                if (value == null)
                    this.anyField = new List<System.Xml.XmlElement>();
                else
                    this.anyField = new List<System.Xml.XmlElement>(value);
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
