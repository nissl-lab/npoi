using NPOI.OpenXml4Net.Util;
using System;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class TableDocument
    {
        CT_Table ctTable = null;

        public TableDocument()
        { 
        }
        public TableDocument(CT_Table table)
        {
            this.ctTable = table;
        }

        public static TableDocument Parse(XmlDocument xmldoc, XmlNamespaceManager namespaceMgr)
        {
            CT_Table obj = new CT_Table();
            XmlElement tableNode = xmldoc.DocumentElement;
            obj.id = XmlHelper.ReadUInt(tableNode.Attributes["id"]);
            obj.name = XmlHelper.ReadString(tableNode.Attributes["name"]);
            obj.displayName = XmlHelper.ReadString(tableNode.Attributes["displayName"]);
            if (tableNode.Attributes["tableType"] != null)
            {
                obj.tableType = (ST_TableType)Enum.Parse(typeof(ST_TableType), tableNode.Attributes["tableType"].Value);
            }
            obj.totalsRowCount = XmlHelper.ReadUInt(tableNode.GetAttributeNode("totalsRowCount"));
            obj.totalsRowShown = XmlHelper.ReadBool(tableNode.GetAttributeNode("totalsRowShown"));
            obj.@ref = XmlHelper.ReadString(tableNode.Attributes["ref"]);
            XmlNode autoFilter = xmldoc.SelectSingleNode("//d:autoFilter", namespaceMgr);
            if (autoFilter != null)
            {
                obj.autoFilter = new CT_AutoFilter();
                obj.autoFilter.@ref = XmlHelper.ReadString(autoFilter.Attributes["name"]);
            }
            XmlNodeList tableCols =  xmldoc.SelectNodes("//d:tableColumns/d:tableColumn", namespaceMgr);
            if (tableCols != null)
            {
                obj.tableColumns = new CT_TableColumns();
                obj.tableColumns.count = (uint)tableCols.Count;
                foreach (XmlElement tableColNode in tableCols)
                {
                    CT_TableColumn ctTableCol = new CT_TableColumn();
                    ctTableCol.id = XmlHelper.ReadUInt(tableColNode.GetAttributeNode("id"));
                    ctTableCol.name = XmlHelper.ReadString(tableColNode.GetAttributeNode("name")); 
                    ctTableCol.uniqueName = XmlHelper.ReadString(tableColNode.GetAttributeNode("uniqueName"));
                    ctTableCol.totalsRowCellStyle = XmlHelper.ReadString(tableColNode.GetAttributeNode("totalsRowCellStyle"));
                    ctTableCol.totalsRowDxfId = XmlHelper.ReadUInt(tableColNode.GetAttributeNode("totalsRowDxfId"));
                    ctTableCol.totalsRowLabel = XmlHelper.ReadString(tableColNode.GetAttributeNode("totalsRowDxfId"));
                    ctTableCol.queryTableFieldId = XmlHelper.ReadUInt(tableColNode.GetAttributeNode("queryTableFieldId"));
                    ctTableCol.xmlColumnPr = CT_XmlColumnPr.Parse(tableColNode.SelectSingleNode("d:xmlColumnPr", namespaceMgr), namespaceMgr);
                    //TODO: parse sub element of CT_TableColumn
                    obj.tableColumns.tableColumn.Add(ctTableCol);
                }
            }
            XmlNode tableStyleInfo = xmldoc.SelectSingleNode("//d:tableStyleInfo", namespaceMgr);
            if (tableStyleInfo != null)
            {
                obj.tableStyleInfo = new CT_TableStyleInfo();
                obj.tableStyleInfo.name = XmlHelper.ReadString(tableStyleInfo.Attributes["name"]);
                obj.tableStyleInfo.showFirstColumn = XmlHelper.ReadBool(tableStyleInfo.Attributes["showFirstColumn"]);
                obj.tableStyleInfo.showLastColumn = XmlHelper.ReadBool(tableStyleInfo.Attributes["showLastColumn"]);
                obj.tableStyleInfo.showRowStripes = XmlHelper.ReadBool(tableStyleInfo.Attributes["showRowStripes"]);
                obj.tableStyleInfo.showColumnStripes = XmlHelper.ReadBool(tableStyleInfo.Attributes["showColumnStripes"]);
            }
            return new TableDocument(obj);
        }

        public CT_Table GetTable()
        {
            return ctTable;
        }

        public void SetTable(CT_Table table)
        {
            this.ctTable = table;
        }

        public void Save(Stream stream)
        {
            //serializer.Serialize(stream, ctTable);
            using (StreamWriter sw = new StreamWriter(stream))
            {
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                sw.Write("<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
                XmlHelper.WriteAttribute(sw, "id", this.ctTable.id);
                XmlHelper.WriteAttribute(sw, "name", this.ctTable.name);
                XmlHelper.WriteAttribute(sw, "displayName", this.ctTable.displayName);
                XmlHelper.WriteAttribute(sw, "ref", this.ctTable.@ref);
                XmlHelper.WriteAttribute(sw, "tableType", this.ctTable.tableType.ToString());
                XmlHelper.WriteAttribute(sw, "totalsRowCount", this.ctTable.totalsRowCount);
                XmlHelper.WriteAttribute(sw, "totalsRowShown", this.ctTable.totalsRowShown);
                sw.Write(">");
                if (this.ctTable.tableColumns != null)
                {
                    sw.Write("<tableColumns count=\"{0}\">", this.ctTable.tableColumns.count);
                    if (this.ctTable.tableColumns.tableColumn != null)
                    {
                        foreach (CT_TableColumn ctTableCol in this.ctTable.tableColumns.tableColumn)
                        {
                            sw.Write(string.Format("<tableColumn id=\"{0}\" name=\"{1}\"", ctTableCol.id, ctTableCol.name));
                            XmlHelper.WriteAttribute(sw, "uniqueName", ctTableCol.uniqueName);
                            XmlHelper.WriteAttribute(sw, "totalsRowCellStyle", ctTableCol.totalsRowCellStyle);
                            XmlHelper.WriteAttribute(sw, "totalsRowLabel", ctTableCol.totalsRowLabel);
                            XmlHelper.WriteAttribute(sw, "totalsRowDxfId", ctTableCol.totalsRowDxfId);
                            XmlHelper.WriteAttribute(sw, "queryTableFieldId", ctTableCol.queryTableFieldId);
                            sw.Write(">");
                            if (ctTableCol.xmlColumnPr != null)
                            {
                                ctTableCol.xmlColumnPr.Write(sw);
                            }
                            sw.Write("</tableColumn>");
                        }
                    }
                    sw.Write("</tableColumns>");
                }
                if (this.ctTable.tableStyleInfo != null)
                {
                    sw.Write("<tableStyleInfo");
                    XmlHelper.WriteAttribute(sw, "name", ctTable.tableStyleInfo.name);
                    XmlHelper.WriteAttribute(sw, "showFirstColumn", ctTable.tableStyleInfo.showFirstColumn);
                    XmlHelper.WriteAttribute(sw, "showLastColumn", ctTable.tableStyleInfo.showLastColumn);
                    XmlHelper.WriteAttribute(sw, "showColumnStripes", ctTable.tableStyleInfo.showColumnStripes);
                    XmlHelper.WriteAttribute(sw, "showRowStripes", ctTable.tableStyleInfo.showRowStripes);
                    sw.Write("/>");
                }
                sw.Write("</table>");
            }
        }
    }
}
