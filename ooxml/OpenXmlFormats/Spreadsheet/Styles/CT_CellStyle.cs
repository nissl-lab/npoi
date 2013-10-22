using System;
using System.Collections.Generic;

using System.Text;
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
            this.extLstField = new CT_ExtensionList();
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
