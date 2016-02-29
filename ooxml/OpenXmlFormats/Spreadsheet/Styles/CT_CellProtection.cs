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
    public class CT_CellProtection
    {

        private bool lockedField = false;

        private bool hiddenField = false;

        public bool IsSetHidden()
        {
            return this.hiddenField != false;
        }
        public bool IsSetLocked()
        {
            return this.lockedField != false;
        }
        public static CT_CellProtection Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CellProtection ctObj = new CT_CellProtection();
            ctObj.locked = XmlHelper.ReadBool(node.Attributes["locked"]);
            ctObj.hidden = XmlHelper.ReadBool(node.Attributes["hidden"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "locked", this.locked);
            if(this.hidden)
                XmlHelper.WriteAttribute(sw, "hidden", this.hidden);
            sw.Write("/>");
        }

        [XmlAttribute]
        public bool locked
        {
            get
            {
                return this.lockedField;
            }
            set
            {
                this.lockedField = value;
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
    }
}
