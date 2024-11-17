using NPOI.OpenXml4Net.Util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot("MapInfo", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public class CT_MapInfo
    {
        //  TODO the initial elements of schemaField and mapField must be ensured somewhere else - or is there a save default!?

        private List<CT_Schema> schemaField = new List<CT_Schema>(); // 1..* 

        private List<CT_Map> mapField = new List<CT_Map>(); // 1..*

        private string selectionNamespacesField = string.Empty; // 1..1

        public static CT_MapInfo Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_MapInfo ctObj = new CT_MapInfo();
            ctObj.SelectionNamespaces = XmlHelper.ReadString(node.Attributes["SelectionNamespaces"]);
            foreach(XmlNode cn in node.ChildNodes)
            {
                if(cn.LocalName == "Schema")
                {
                    ctObj.Schema.Add(CT_Schema.Parse(cn, namespaceManager));
                }
                else if(cn.LocalName == "Map")
                {
                    ctObj.Map.Add(CT_Map.Parse(cn, namespaceManager));
                }
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "SelectionNamespaces", this.SelectionNamespaces);
            sw.Write(">");
            foreach(CT_Schema ctSchema in Schema)
            {
                ctSchema.Write(sw, "Schema");
            }

            foreach(CT_Map ctMap in Map)
            {
                ctMap.Write(sw, "Map");
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }
        [XmlElement("Schema")]
        public List<CT_Schema> Schema
        {
            get
            {
                return this.schemaField;
            }
            set
            {
                this.schemaField = value;
            }
        }

        [XmlElement("Map")]
        public List<CT_Map> Map
        {
            get
            {
                return this.mapField;
            }
            set
            {
                this.mapField = value;
            }
        }

        [XmlAttribute]
        public string SelectionNamespaces
        {
            get
            {
                return this.selectionNamespacesField;
            }
            set
            {
                this.selectionNamespacesField = value;
            }
        }
    }

    [Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Schema
    {
        private System.Xml.XmlElement anyField; // TODO ensure initialization = new XmlElement(); // 1..1

        private string idField = string.Empty;  // 1..1

        private string schemaRefField = null; // 0..1

        private string namespaceField = null; // 0..1

        private string schemaLanguageField;

        public static CT_Schema Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_Schema ctObj = new CT_Schema();
            ctObj.ID = node.Attributes["ID"].Value;
            ctObj.Namespace = XmlHelper.ReadString(node.Attributes["Namespace"]);
            ctObj.SchemaRef = XmlHelper.ReadString(node.Attributes["SchemaRef"]);
            ctObj.SchemaLanguage = XmlHelper.ReadString(node.Attributes["SchemaLanguage"]);
            ctObj.Any = node.FirstChild as XmlElement;
            return ctObj;
        }
        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "ID", this.ID);
            XmlHelper.WriteAttribute(sw, "Namespace", this.Namespace);
            XmlHelper.WriteAttribute(sw, "SchemaRef", this.SchemaRef);
            XmlHelper.WriteAttribute(sw, "SchemaLanguage", this.SchemaLanguage);
            sw.Write(">");
            if(anyField != null)
                sw.Write(anyField.OuterXml);
            sw.Write(string.Format("</{0}>", nodeName));
        }
        [XmlAnyElement]
        public System.Xml.XmlElement Any
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

        // uppercase ID is correct!
        // (Form = XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        // namespace is really http://schemas.openxmlformats.org/spreadsheetml/2006/main
        [XmlAttribute]
        public string ID
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


        [XmlAttribute]
        public string SchemaRef
        {
            get
            {
                return this.schemaRefField;
            }
            set
            {
                this.schemaRefField = value;
            }
        }
        [XmlIgnore]
        public bool SchemaRefSpecified
        {
            get { return null != this.schemaRefField; }
        }


        [XmlAttribute]
        public string Namespace
        {
            get
            {
                return this.namespaceField;
            }
            set
            {
                this.namespaceField = value;
            }
        }
        [XmlIgnore]
        public bool NamespaceSpecified
        {
            get { return null != this.namespaceField; }
        }

        public string SchemaLanguage
        {
            get
            {
                return this.schemaLanguageField;
            }
            set
            {
                schemaLanguageField = value;
            }
        }

        [XmlIgnore]
        public bool SchemaLanguageSpecified
        {
            get { return null != this.schemaLanguageField; }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Map
    {
        private CT_DataBinding dataBindingField = null; // 0..1  element

        // the following are all required attributes

        private uint idField;

        private string nameField;

        private string rootElementField;

        private string schemaIDField;

        private bool showImportExportValidationErrorsField;

        private bool autoFitField;

        private bool appendField;

        private bool preserveSortAFLayoutField;

        private bool preserveFormatField;

        public static CT_Map Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_Map ctMap = new CT_Map();
            ctMap.ID = XmlHelper.ReadUInt(node.Attributes["ID"]);
            ctMap.Name = XmlHelper.ReadString(node.Attributes["Name"]);
            ctMap.RootElement = XmlHelper.ReadString(node.Attributes["RootElement"]);
            ctMap.SchemaID = XmlHelper.ReadString(node.Attributes["SchemaID"]);
            ctMap.ShowImportExportValidationErrors = XmlHelper.ReadBool(node.Attributes["ShowImportExportValidationErrors"]);
            ctMap.PreserveFormat = XmlHelper.ReadBool(node.Attributes["PreserveFormat"]);
            ctMap.PreserveSortAFLayout = XmlHelper.ReadBool(node.Attributes["PreserveSortAFLayout"]);
            ctMap.Append = XmlHelper.ReadBool(node.Attributes["Append"]);
            ctMap.AutoFit = XmlHelper.ReadBool(node.Attributes["AutoFit"]);
            foreach(XmlElement ele in node)
            {
                if(ele.LocalName == "DataBinding")
                {
                    ctMap.DataBinding = CT_DataBinding.Parse(ele, namespaceManager);
                }
            }
            return ctMap;
        }
        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "ID", this.ID);
            XmlHelper.WriteAttribute(sw, "Name", this.Name);
            XmlHelper.WriteAttribute(sw, "RootElement", this.RootElement);
            XmlHelper.WriteAttribute(sw, "SchemaID", this.SchemaID);
            XmlHelper.WriteAttribute(sw, "ShowImportExportValidationErrors", this.ShowImportExportValidationErrors);
            XmlHelper.WriteAttribute(sw, "PreserveFormat", this.PreserveFormat);
            XmlHelper.WriteAttribute(sw, "PreserveSortAFLayout", this.PreserveSortAFLayout);
            XmlHelper.WriteAttribute(sw, "Append", this.Append);
            XmlHelper.WriteAttribute(sw, "AutoFit", this.AutoFit);
            
            if(dataBindingField != null)
            {
                sw.Write(">");
                dataBindingField.Write(sw, "DataBinding");
                sw.Write(string.Format("</{0}>", nodeName));
            }
            else
                sw.Write("/>");
        }
        [XmlElement]
        public CT_DataBinding DataBinding
        {
            get
            {
                return this.dataBindingField;
            }
            set
            {
                this.dataBindingField = value;
            }
        }
        [XmlIgnore]
        public bool DataBindingSpecified
        {
            get { return (null != dataBindingField); }
        }


        [XmlAttribute]
        public uint ID
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


        [XmlAttribute]
        public string Name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }


        [XmlAttribute]
        public string RootElement
        {
            get
            {
                return this.rootElementField;
            }
            set
            {
                this.rootElementField = value;
            }
        }


        [XmlAttribute]
        public string SchemaID
        {
            get
            {
                return this.schemaIDField;
            }
            set
            {
                this.schemaIDField = value;
            }
        }


        [XmlAttribute]
        public bool ShowImportExportValidationErrors
        {
            get
            {
                return this.showImportExportValidationErrorsField;
            }
            set
            {
                this.showImportExportValidationErrorsField = value;
            }
        }


        [XmlAttribute]
        public bool AutoFit
        {
            get
            {
                return this.autoFitField;
            }
            set
            {
                this.autoFitField = value;
            }
        }


        [XmlAttribute]
        public bool Append
        {
            get
            {
                return this.appendField;
            }
            set
            {
                this.appendField = value;
            }
        }


        [XmlAttribute]
        public bool PreserveSortAFLayout
        {
            get
            {
                return this.preserveSortAFLayoutField;
            }
            set
            {
                this.preserveSortAFLayoutField = value;
            }
        }


        [XmlAttribute]
        public bool PreserveFormat
        {
            get
            {
                return this.preserveFormatField;
            }
            set
            {
                this.preserveFormatField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_DataBinding
    {

        private System.Xml.XmlElement anyField; // 1..1 element

        // now attributes

        private string dataBindingNameField = null; // 0..1

        private bool? fileBindingField = null; // 0..1

        private uint? connectionIDField = null; // 0..1

        private string fileBindingNameField = null; // 0..1

        private uint dataBindingLoadModeField = 0; // 1..1 - default value not defined in xsd

        public static CT_DataBinding Parse(XmlNode ele, XmlNamespaceManager namespaceManager)
        {
            CT_DataBinding ctObj = new CT_DataBinding();
            ctObj.Any = ele.FirstChild as XmlElement;
            ctObj.connectionIDField = ele.Attributes["ConnectionID"] ==null ? null : XmlHelper.ReadUInt(ele.Attributes["ConnectionID"]);
            ctObj.DataBindingName = XmlHelper.ReadString(ele.Attributes["DataBindingName"]);
            ctObj.FileBindingName = XmlHelper.ReadString(ele.Attributes["FileBindingName"]);
            ctObj.DataBindingLoadMode = XmlHelper.ReadUInt(ele.Attributes["DataBindingLoadMode"]);
            ctObj.fileBindingField = ele.Attributes["FileBinding"] ==null ? null : XmlHelper.ReadBool(ele.Attributes["FileBinding"]);
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            if(this.FileBindingSpecified)
                XmlHelper.WriteAttribute(sw, "FileBinding", FileBinding);
            if(this.ConnectionIDSpecified)
                XmlHelper.WriteAttribute(sw, "ConnectionID", ConnectionID);
            XmlHelper.WriteAttribute(sw, "DataBindingLoadMode", DataBindingLoadMode);
            XmlHelper.WriteAttribute(sw, "DataBindingName", DataBindingName);
            XmlHelper.WriteAttribute(sw, "FileBindingName", FileBindingName);
            sw.Write(">");
            if(anyField != null)
                sw.Write(anyField.OuterXml);
            sw.Write(string.Format("</{0}>", nodeName));
        }
        [XmlAnyElement]
        public System.Xml.XmlElement Any
        {
            get { return this.anyField; }
            set { this.anyField = value; }
        }


        [XmlAttribute]
        public string DataBindingName
        {
            get { return this.dataBindingNameField; }
            set { this.dataBindingNameField = value; }
        }
        [XmlIgnore]
        public bool outlineSpecified
        {
            get { return (null != dataBindingNameField); }
        }


        [XmlAttribute]
        public bool FileBinding
        {
            get { return null == this.fileBindingField ? false : (bool)this.fileBindingField; } // default value not defined in xsd
            set { this.fileBindingField = value; }
        }
        [XmlIgnore]
        public bool FileBindingSpecified
        {
            get { return (null != fileBindingField); }
        }


        [XmlAttribute]
        public uint ConnectionID
        {
            get { return null == this.connectionIDField ? 0 : (uint)this.connectionIDField; } // default value not defined in xsd
            set { this.connectionIDField = value; }
        }

        [XmlIgnore]
        public bool ConnectionIDSpecified
        {
            get { return (null != connectionIDField); }
        }


        [XmlAttribute]
        public string FileBindingName
        {
            get { return this.fileBindingNameField; }
            set { this.fileBindingNameField = value; }
        }
        [XmlIgnore]
        public bool FileBindingNameSpecified
        {
            get { return (null != fileBindingNameField); }
        }


        [XmlAttribute]
        public uint DataBindingLoadMode
        {
            get { return this.dataBindingLoadModeField; }
            set { this.dataBindingLoadModeField = value; }
        }
    }
}
