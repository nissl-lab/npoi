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
    [XmlRoot("sheets", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public class CT_Sheets
    {
        private List<CT_Sheet> sheetField; // required field

        public CT_Sheets()
        {
            this.sheetField = new List<CT_Sheet>();
        }
        public static CT_Sheets Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Sheets ctObj = new CT_Sheets();
            ctObj.sheet = new List<CT_Sheet>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "sheet")
                    ctObj.sheet.Add(CT_Sheet.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            sw.Write(">");
            if (this.sheet != null)
            {
                foreach (CT_Sheet x in this.sheet)
                {
                    x.Write(sw, "sheet");
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_Sheet AddNewSheet()
        {
            if (this.sheetField == null)
                this.sheetField = new List<CT_Sheet>();
            CT_Sheet newsheet = new CT_Sheet();
            this.sheetField.Add(newsheet);
            return newsheet;
        }
        public void RemoveSheet(int index)
        {
            if (sheetField == null)
                return;
            sheetField.RemoveAt(index);
        }
        public CT_Sheet InsertNewSheet(int index)
        {
            if (this.sheetField == null)
                this.sheetField = new List<CT_Sheet>();
            CT_Sheet newsheet = new CT_Sheet();
            this.sheetField.Insert(index, newsheet);
            return newsheet;
        }
        public CT_Sheet GetSheetArray(int index)
        {
            if (this.sheetField == null)
                return null;

            return this.sheetField[index];
        }
        [XmlElement("sheet")]
        public List<CT_Sheet> sheet
        {
            get
            {
                return this.sheetField;
            }
            set
            {
                this.sheetField = value;
            }
        }
    }
}
