using NPOI.OpenXml4Net.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_CellStyle
    {

        private CT_ExtensionList extLstField;

        private string nameField;

        private uint xfIdField;

        private uint builtinIdField;

        private bool builtinIdFieldSpecified;

        private uint iLevelField;

        private bool iLevelFieldSpecified;

        private bool hiddenField;

        private bool hiddenFieldSpecified;

        private bool customBuiltinField;

        private bool customBuiltinFieldSpecified;

        public CT_CellStyle()
        {
           // this.extLstField = new CT_ExtensionList();
        }
        public static CT_CellStyle Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CellStyle ctObj = new CT_CellStyle();
            ctObj.name = XmlHelper.ReadString(node.Attributes["name"]);
            ctObj.xfId = XmlHelper.ReadUInt(node.Attributes["xfId"]);
            ctObj.builtinId = XmlHelper.ReadUInt(node.Attributes["builtinId"]);
            ctObj.iLevel = XmlHelper.ReadUInt(node.Attributes["iLevel"]);
            ctObj.hidden = XmlHelper.ReadBool(node.Attributes["hidden"]);
            ctObj.customBuiltin = XmlHelper.ReadBool(node.Attributes["customBuiltin"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "name", this.name);
            XmlHelper.WriteAttribute(sw, "xfId", this.xfId, true);
            XmlHelper.WriteAttribute(sw, "builtinId", this.builtinId,true);
            XmlHelper.WriteAttribute(sw, "iLevel", this.iLevel);
            XmlHelper.WriteAttribute(sw, "hidden", this.hidden, false);
            XmlHelper.WriteAttribute(sw, "customBuiltin", this.customBuiltin, false);
            if (this.extLst != null)
            {
                sw.Write(">");
                this.extLst.Write(sw, "extLst");
                sw.Write(string.Format("</{0}>", nodeName));
            }
            else
            {
                sw.Write("/>");
            }
        }

        [XmlElement]
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }
        [XmlAttribute]
        public string name
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
        public uint xfId
        {
            get
            {
                return this.xfIdField;
            }
            set
            {
                this.xfIdField = value;
            }
        }
        [XmlAttribute]
        public uint builtinId
        {
            get
            {
                return this.builtinIdField;
            }
            set
            {
                this.builtinIdField = value;
            }
        }

        [XmlIgnore]
        public bool builtinIdSpecified
        {
            get
            {
                return this.builtinIdFieldSpecified;
            }
            set
            {
                this.builtinIdFieldSpecified = value;
            }
        }
        [XmlAttribute]
        public uint iLevel
        {
            get
            {
                return this.iLevelField;
            }
            set
            {
                this.iLevelField = value;
            }
        }

        [XmlIgnore]
        public bool iLevelSpecified
        {
            get
            {
                return this.iLevelFieldSpecified;
            }
            set
            {
                this.iLevelFieldSpecified = value;
            }
        }
        [XmlAttribute]
        public bool hidden
        {
            get
            {
                return this.hiddenField;
            }
            set
            {
                this.hiddenField = value;
            }
        }

        [XmlIgnore]
        public bool hiddenSpecified
        {
            get
            {
                return this.hiddenFieldSpecified;
            }
            set
            {
                this.hiddenFieldSpecified = value;
            }
        }
        [XmlAttribute]
        public bool customBuiltin
        {
            get
            {
                return this.customBuiltinField;
            }
            set
            {
                this.customBuiltinField = value;
            }
        }

        [XmlIgnore]
        public bool customBuiltinSpecified
        {
            get
            {
                return this.customBuiltinFieldSpecified;
            }
            set
            {
                this.customBuiltinFieldSpecified = value;
            }
        }
    }


}
