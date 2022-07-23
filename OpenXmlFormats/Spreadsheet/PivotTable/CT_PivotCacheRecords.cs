using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;
using NPOI.OpenXml4Net.Util;
using System.Xml;
using System.IO;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot("pivotCacheRecords", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public partial class CT_PivotCacheRecords
    {
        public static CT_PivotCacheRecords Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PivotCacheRecords ctObj = new CT_PivotCacheRecords();
            if (node.Attributes["count"] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes["count"]);
            ctObj.r = new List<CT_PivotCacheRecord>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "extLst")
                {
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "r")
                {
                    ctObj.rField.Add(CT_PivotCacheRecord.Parse(childNode, namespaceManager));
                }
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw)
        {
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write("<pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ");
            sw.Write("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ");
            sw.Write("xmlns:s=\"http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes\" ");
            sw.Write("xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" ");
            sw.Write("mc:Ignorable=\"xr\" ");
            sw.Write("xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" ");
            XmlHelper.WriteAttribute(sw, "count", this.count);
            sw.Write(">");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            if (this.r != null && this.r.Count > 0)
            {
                foreach (CT_PivotCacheRecord o in this.r)
                {
                    o.Write(sw);
                }
            }
            sw.Write("</pivotCacheRecords>");
        }

        public void Save(Stream stream)
        {
            using (StreamWriter sw = new StreamWriter(stream))
            {
                //TODO add namespaceUri
                this.Write(sw);
            }
        }

        private List<CT_PivotCacheRecord> rField;

        private CT_ExtensionList extLstField;

        private uint countField;

        private bool countFieldSpecified;

        [System.Xml.Serialization.XmlArrayAttribute(Order = 0)]
        public List<CT_PivotCacheRecord> r
        {
            get
            {
                return this.rField;
            }
            set
            {
                this.rField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PivotCacheRecord
    {
        private List<object> items;

        public static CT_PivotCacheRecord Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_PivotCacheRecord ctObj = new CT_PivotCacheRecord();
            ctObj.fields = new List<object>();

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "n")
                {
                    ctObj.fields.Add(CT_Number.Parse(childNode, namespaceManager));
                }
                else if (childNode.LocalName == "b")
                {
                    ctObj.fields.Add(CT_Boolean.Parse(childNode, namespaceManager));
                }
                else if (childNode.LocalName == "d")
                {
                    ctObj.fields.Add(CT_DateTime.Parse(childNode, namespaceManager));
                }
                else if (childNode.LocalName == "e")
                {
                    ctObj.fields.Add(CT_Error.Parse(childNode, namespaceManager));
                }
                else if (childNode.LocalName == "m")
                {
                    ctObj.fields.Add(CT_Missing.Parse(childNode, namespaceManager));
                }
                else if (childNode.LocalName == "s")
                {
                    ctObj.fields.Add(CT_String.Parse(childNode, namespaceManager));
                }
                else if (childNode.LocalName == "x")
                {
                    ctObj.fields.Add(CT_Index.Parse(childNode, namespaceManager));
                }
            }

            return ctObj;
        }

        [System.Xml.Serialization.XmlArrayItemAttribute("b", typeof(CT_Boolean), IsNullable = false)]
        [System.Xml.Serialization.XmlArrayItemAttribute("d", typeof(CT_DateTime), IsNullable = false)]
        [System.Xml.Serialization.XmlArrayItemAttribute("e", typeof(CT_Error), IsNullable = false)]
        [System.Xml.Serialization.XmlArrayItemAttribute("m", typeof(CT_Missing), IsNullable = false)]
        [System.Xml.Serialization.XmlArrayItemAttribute("n", typeof(CT_Number), IsNullable = false)]
        [System.Xml.Serialization.XmlArrayItemAttribute("s", typeof(CT_String), IsNullable = false)]
        [System.Xml.Serialization.XmlArrayItemAttribute("x", typeof(CT_Index), IsNullable = false)]
        public List<object> fields
        {
            get
            {
                return this.items;
            }
            set
            {
                this.items = value;
            }
        }

        internal void Write(StreamWriter sw)
        {
            if (this.fields != null && fields.Count > 0)
            {
                sw.Write("<r>");

                foreach (object o in this.fields)
                {
                    if (o is CT_Number)
                        ((CT_Number)o).Write(sw, "n");
                    else if (o is CT_Boolean)
                        ((CT_Boolean)o).Write(sw, "b");
                    else if (o is CT_DateTime)
                        ((CT_DateTime)o).Write(sw, "d");
                    else if (o is CT_Error)
                        ((CT_Error)o).Write(sw, "e");
                    else if (o is CT_Missing)
                        ((CT_Missing)o).Write(sw, "m");
                    else if (o is CT_String)
                        ((CT_String)o).Write(sw, "s");
                    else if (o is CT_Index)
                        ((CT_Index)o).Write(sw, "x");
                }

                sw.Write("</r>");
            }
        }
    }
}
