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
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot("sheet", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public class CT_Sheet
    {
        public static CT_Sheet Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Sheet ctObj = new CT_Sheet
            {
                name = XmlHelper.ReadString(node.Attributes[nameof(name)]),
                sheetId = XmlHelper.ReadUInt(node.Attributes[nameof(sheetId)])
            };
            if (node.Attributes[nameof(state)] != null)
                ctObj.state = (ST_SheetState)Enum.Parse(typeof(ST_SheetState), node.Attributes[nameof(state)].Value);
            ctObj.id = XmlHelper.ReadString(node.Attributes[nameof(id), PackageNamespaces.SCHEMA_RELATIONSHIPS]);
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");
            XmlHelper.WriteAttribute(sw, nameof(name), this.name);
            XmlHelper.WriteAttribute(sw, nameof(sheetId), this.sheetId);
            if (state != ST_SheetState.visible)
                XmlHelper.WriteAttribute(sw, nameof(state), this.state.ToString());
            XmlHelper.WriteAttribute(sw, $"r:{nameof(id)}", this.id);
            sw.Write(">");
            sw.Write($"</{nodeName}>");
        }

        public CT_Sheet()
        {
            this.state = ST_SheetState.visible;
        }
        public void Set(CT_Sheet sheet)
        {
            this.name = sheet.name;
            this.sheetId = sheet.sheetId;
            this.state = sheet.state;
            this.id = sheet.id;
        }
        public CT_Sheet Copy()
        {
            CT_Sheet obj = new CT_Sheet();
            obj.id = this.id;
            obj.sheetId = this.sheetId;
            obj.name = this.name;
            obj.state = this.state;
            return obj;
        }
        [XmlAttribute(nameof(name))]
        public string name { get; set; }
        [XmlAttribute(nameof(sheetId))]
        public uint sheetId { get; set; }
        [XmlAttribute(nameof(state))]
        [DefaultValue(ST_SheetState.visible)]
        public ST_SheetState state { get; set; }
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id { get; set; }
    }
}
