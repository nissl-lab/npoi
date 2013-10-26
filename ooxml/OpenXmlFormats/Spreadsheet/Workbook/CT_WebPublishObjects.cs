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

    public enum ST_TargetScreenSize
    {


        [XmlEnum("544x376")]
        Item544x376,


        [XmlEnum("640x480")]
        Item640x480,


        [XmlEnum("720x512")]
        Item720x512,


        [XmlEnum("800x600")]
        Item800x600,


        [XmlEnum("1024x768")]
        Item1024x768,


        [XmlEnum("1152x882")]
        Item1152x882,


        [XmlEnum("1152x900")]
        Item1152x900,


        [XmlEnum("1280x1024")]
        Item1280x1024,


        [XmlEnum("1600x1200")]
        Item1600x1200,


        [XmlEnum("1800x1440")]
        Item1800x1440,


        [XmlEnum("1920x1200")]
        Item1920x1200,
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_WebPublishing
    {

        private bool cssField;

        private bool thicketField;

        private bool longFileNamesField;

        private bool vmlField;

        private bool allowPngField;

        private ST_TargetScreenSize targetScreenSizeField;

        private uint dpiField;

        private uint codePageField;

        private bool codePageFieldSpecified;
        public static CT_WebPublishing Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_WebPublishing ctObj = new CT_WebPublishing();
            ctObj.css = XmlHelper.ReadBool(node.Attributes["css"]);
            ctObj.thicket = XmlHelper.ReadBool(node.Attributes["thicket"]);
            ctObj.longFileNames = XmlHelper.ReadBool(node.Attributes["longFileNames"]);
            ctObj.vml = XmlHelper.ReadBool(node.Attributes["vml"]);
            ctObj.allowPng = XmlHelper.ReadBool(node.Attributes["allowPng"]);
            if (node.Attributes["targetScreenSize"] != null)
                ctObj.targetScreenSize = (ST_TargetScreenSize)Enum.Parse(typeof(ST_TargetScreenSize), node.Attributes["targetScreenSize"].Value);
            ctObj.dpi = XmlHelper.ReadUInt(node.Attributes["dpi"]);
            ctObj.codePage = XmlHelper.ReadUInt(node.Attributes["codePage"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "css", this.css);
            XmlHelper.WriteAttribute(sw, "thicket", this.thicket);
            XmlHelper.WriteAttribute(sw, "longFileNames", this.longFileNames);
            XmlHelper.WriteAttribute(sw, "vml", this.vml);
            XmlHelper.WriteAttribute(sw, "allowPng", this.allowPng);
            XmlHelper.WriteAttribute(sw, "targetScreenSize", this.targetScreenSize.ToString());
            XmlHelper.WriteAttribute(sw, "dpi", this.dpi);
            XmlHelper.WriteAttribute(sw, "codePage", this.codePage);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_WebPublishing()
        {
            this.cssField = true;
            this.thicketField = true;
            this.longFileNamesField = true;
            this.vmlField = false;
            this.allowPngField = false;
            this.targetScreenSizeField = ST_TargetScreenSize.Item800x600;
            this.dpiField = ((uint)(96));
        }

        [DefaultValue(true)]
        public bool css
        {
            get
            {
                return this.cssField;
            }
            set
            {
                this.cssField = value;
            }
        }

        [DefaultValue(true)]
        public bool thicket
        {
            get
            {
                return this.thicketField;
            }
            set
            {
                this.thicketField = value;
            }
        }

        [DefaultValue(true)]
        public bool longFileNames
        {
            get
            {
                return this.longFileNamesField;
            }
            set
            {
                this.longFileNamesField = value;
            }
        }

        [DefaultValue(false)]
        public bool vml
        {
            get
            {
                return this.vmlField;
            }
            set
            {
                this.vmlField = value;
            }
        }

        [DefaultValue(false)]
        public bool allowPng
        {
            get
            {
                return this.allowPngField;
            }
            set
            {
                this.allowPngField = value;
            }
        }

        [DefaultValue(ST_TargetScreenSize.Item800x600)]
        public ST_TargetScreenSize targetScreenSize
        {
            get
            {
                return this.targetScreenSizeField;
            }
            set
            {
                this.targetScreenSizeField = value;
            }
        }

        [DefaultValue(typeof(uint), "96")]
        public uint dpi
        {
            get
            {
                return this.dpiField;
            }
            set
            {
                this.dpiField = value;
            }
        }

        public uint codePage
        {
            get
            {
                return this.codePageField;
            }
            set
            {
                this.codePageField = value;
            }
        }

        [XmlIgnore]
        public bool codePageSpecified
        {
            get
            {
                return this.codePageFieldSpecified;
            }
            set
            {
                this.codePageFieldSpecified = value;
            }
        }
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_WebPublishObject
    {

        private uint idField;

        private string divIdField;

        private string sourceObjectField;

        private string destinationFileField;

        private string titleField;

        private bool autoRepublishField;

        public CT_WebPublishObject()
        {
            this.autoRepublishField = false;
        }
        public static CT_WebPublishObject Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_WebPublishObject ctObj = new CT_WebPublishObject();
            ctObj.id = XmlHelper.ReadUInt(node.Attributes["id"]);
            ctObj.divId = XmlHelper.ReadString(node.Attributes["divId"]);
            ctObj.sourceObject = XmlHelper.ReadString(node.Attributes["sourceObject"]);
            ctObj.destinationFile = XmlHelper.ReadString(node.Attributes["destinationFile"]);
            ctObj.title = XmlHelper.ReadString(node.Attributes["title"]);
            ctObj.autoRepublish = XmlHelper.ReadBool(node.Attributes["autoRepublish"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "id", this.id);
            XmlHelper.WriteAttribute(sw, "divId", this.divId);
            XmlHelper.WriteAttribute(sw, "sourceObject", this.sourceObject);
            XmlHelper.WriteAttribute(sw, "destinationFile", this.destinationFile);
            XmlHelper.WriteAttribute(sw, "title", this.title);
            XmlHelper.WriteAttribute(sw, "autoRepublish", this.autoRepublish);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }
        [XmlAnyAttribute]
        public uint id
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
        [XmlAnyAttribute]
        public string divId
        {
            get
            {
                return this.divIdField;
            }
            set
            {
                this.divIdField = value;
            }
        }
        [XmlAnyAttribute]
        public string sourceObject
        {
            get
            {
                return this.sourceObjectField;
            }
            set
            {
                this.sourceObjectField = value;
            }
        }
        [XmlAnyAttribute]
        public string destinationFile
        {
            get
            {
                return this.destinationFileField;
            }
            set
            {
                this.destinationFileField = value;
            }
        }
        [XmlAnyAttribute]
        public string title
        {
            get
            {
                return this.titleField;
            }
            set
            {
                this.titleField = value;
            }
        }
        [XmlAnyAttribute]
        [DefaultValue(false)]
        public bool autoRepublish
        {
            get
            {
                return this.autoRepublishField;
            }
            set
            {
                this.autoRepublishField = value;
            }
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_WebPublishObjects
    {

        private List<CT_WebPublishObject> webPublishObjectField;

        private uint countField;

        private bool countFieldSpecified;
        public static CT_WebPublishObjects Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_WebPublishObjects ctObj = new CT_WebPublishObjects();
            ctObj.count = XmlHelper.ReadUInt(node.Attributes["count"]);
            ctObj.countSpecified = node.Attributes["count"] != null;
            ctObj.webPublishObject = new List<CT_WebPublishObject>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "webPublishObject")
                    ctObj.webPublishObject.Add(CT_WebPublishObject.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "count", this.count);
            sw.Write(">");
            if (this.webPublishObject != null)
            {
                foreach (CT_WebPublishObject x in this.webPublishObject)
                {
                    x.Write(sw, "webPublishObject");
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_WebPublishObjects()
        {
            //this.webPublishObjectField = new List<CT_WebPublishObject>();
        }
        [XmlElement]
        public List<CT_WebPublishObject> webPublishObject
        {
            get
            {
                return this.webPublishObjectField;
            }
            set
            {
                this.webPublishObjectField = value;
            }
        }
        [XmlAttribute]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        [XmlIgnore]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
    }
}
