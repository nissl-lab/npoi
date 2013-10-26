using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXml4Net.Util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_ExternalReference
    {

        private string idField;
        public static CT_ExternalReference Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ExternalReference ctObj = new CT_ExternalReference();
            ctObj.id = XmlHelper.ReadString(node.Attributes["id", PackageNamespaces.SCHEMA_RELATIONSHIPS]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "r:id", this.id);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id
        {
            get
            {
                return this.idField;
            }
            set
            {
                this.idField = value;
            }
        }
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_ExternalReferences
    {

        private List<CT_ExternalReference> externalReferenceField;
        public static CT_ExternalReferences Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ExternalReferences ctObj = new CT_ExternalReferences();
            ctObj.externalReference = new List<CT_ExternalReference>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "externalReference")
                    ctObj.externalReference.Add(CT_ExternalReference.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            sw.Write(">");
            if (this.externalReference != null)
            {
                foreach (CT_ExternalReference x in this.externalReference)
                {
                    x.Write(sw, "externalReference");
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_ExternalReferences()
        {
            //this.externalReferenceField = new List<CT_ExternalReference>();
        }
        [XmlElement]
        public List<CT_ExternalReference> externalReference
        {
            get
            {
                return this.externalReferenceField;
            }
            set
            {
                this.externalReferenceField = value;
            }
        }
    }
}
