using System;
using System.Collections.Generic;
using System.ComponentModel;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot("sheet", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public class CT_Sheet
    {

        private string nameField;

        private uint sheetIdField;

        private ST_SheetState stateField;

        private string idField;

        public CT_Sheet()
        {
            this.stateField = ST_SheetState.visible;
        }
        public void Set(CT_Sheet sheet)
        {
            this.nameField = sheet.nameField;
            this.sheetIdField = sheet.sheetIdField;
            this.stateField = sheet.stateField;
            this.idField = sheet.idField;
        }
        public CT_Sheet Copy()
        {
            CT_Sheet obj = new CT_Sheet();
            obj.idField = this.idField;
            obj.nameField = this.nameField;
            obj.stateField = this.stateField;
            return obj;
        }
        [XmlAttribute("name")]
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
        [XmlAttribute("sheetId")]
        public uint sheetId
        {
            get
            {
                return this.sheetIdField;
            }
            set
            {
                this.sheetIdField = value;
            }
        }
        [XmlAttribute("state")]
        [DefaultValue(ST_SheetState.visible)]
        public ST_SheetState state
        {
            get
            {
                return this.stateField;
            }
            set
            {
                this.stateField = value;
            }
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
}
