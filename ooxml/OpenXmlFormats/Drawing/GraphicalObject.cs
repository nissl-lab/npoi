
using System;
using System.Xml.Serialization;
using System.Collections.Generic;
using System.Xml;
using System.IO;
using NPOI.OpenXml4Net.Util;

namespace NPOI.OpenXmlFormats.Dml
{
    [Serializable]
    
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_GraphicalObjectData
    {
        public static CT_GraphicalObjectData Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_GraphicalObjectData ctObj = new CT_GraphicalObjectData();
            ctObj.uri = XmlHelper.ReadString(node.Attributes["uri"]);
            ctObj.Any = new List<string>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                ctObj.Any.Add(childNode.OuterXml);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "uri", this.uri);
            sw.Write(">");
            foreach (string x in this.Any)
            {
                sw.Write(x);
            }
            sw.Write(string.Format("</a:{0}>", nodeName));
        }

        private List<string> anyField = new List<string>();

        private string uriField;

        public void AddChartElement(string el)
        {
            anyField.Add(el);
        }
        public void AddPicElement(string el)
        {
            anyField.Add(el);
        }
        //[XmlAnyElement()]
        [XmlIgnore]
        public List<string> Any
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
        public static CT_GraphicalObject Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_GraphicalObject ctObj = new CT_GraphicalObject();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "graphicData")
                    ctObj.graphicData = CT_GraphicalObjectData.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<a:{0}", nodeName));
            sw.Write(">");
            if (this.graphicData != null)
                this.graphicData.Write(sw, "graphicData");
            sw.Write(string.Format("</a:{0}>", nodeName));
        }

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
