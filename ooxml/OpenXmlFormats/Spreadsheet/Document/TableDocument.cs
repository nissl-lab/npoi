using System;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class TableDocument
    {
        static XmlSerializer serializer = new XmlSerializer(typeof(CT_Table));
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
            XmlNode tableNode = xmldoc.DocumentElement;
            obj.id = uint.Parse(tableNode.Attributes["id"].Value);
            obj.name = tableNode.Attributes["name"].Value;
            obj.displayName = tableNode.Attributes["displayName"].Value;
            if (tableNode.Attributes["tableType"] != null)
            {
                obj.tableType = (ST_TableType)Enum.Parse(typeof(ST_TableType), tableNode.Attributes["tableType"].Value);
            }
            if (tableNode.Attributes["totalsRowCount"] != null)
                obj.totalsRowCount = uint.Parse(tableNode.Attributes["totalsRowCount"].Value);
            if(tableNode.Attributes["totalsRowShown"].Value=="1"||tableNode.Attributes["totalsRowShown"].Value.ToLower()=="true")
                obj.totalsRowShown = true;
            obj.@ref = tableNode.Attributes["ref"].Value;
            var autoFilter = xmldoc.SelectSingleNode("//d:autoFilter", namespaceMgr);
            if (autoFilter != null)
            {
                obj.autoFilter = new CT_AutoFilter();
                obj.autoFilter.@ref = autoFilter.Attributes["ref"].Value;
            }
            var tableCols =  xmldoc.SelectNodes("//d:tableColumns/d:tableColumn", namespaceMgr);
            if (tableCols != null)
            {
                obj.tableColumns = new CT_TableColumns();
                obj.tableColumns.count = (uint)tableCols.Count;
                foreach (XmlNode tableColNode in tableCols)
                {
                    CT_TableColumn ctTableCol = new CT_TableColumn();
                    ctTableCol.id = uint.Parse(tableColNode.Attributes["id"].Value);
                    ctTableCol.name = tableColNode.Attributes["name"].Value;
                    if (tableColNode.Attributes["uniqueName"] != null) 
                        ctTableCol.uniqueName = tableColNode.Attributes["uniqueName"].Value;
                    if (tableColNode.Attributes["totalsRowCellStyle"]!=null) 
                        ctTableCol.totalsRowCellStyle = tableColNode.Attributes["totalsRowCellStyle"].Value;
                    if (tableColNode.Attributes["totalsRowDxfId"] != null)
                        ctTableCol.totalsRowDxfId = uint.Parse(tableColNode.Attributes["totalsRowDxfId"].Value);
                    if (tableColNode.Attributes["totalsRowLabel"] != null)
                        ctTableCol.totalsRowLabel = tableColNode.Attributes["totalsRowLabel"].Value;
                    if (tableColNode.Attributes["queryTableFieldId"] != null)
                        ctTableCol.queryTableFieldId = uint.Parse(tableColNode.Attributes["queryTableFieldId"].Value);

                    ctTableCol.xmlColumnPr = CT_XmlColumnPr.Parse(tableColNode.SelectSingleNode("d:xmlColumnPr", namespaceMgr), namespaceMgr);
                    //TODO: parse sub element of CT_TableColumn
                    obj.tableColumns.tableColumn.Add(ctTableCol);
                }
            }
            var tableStyleInfo = xmldoc.SelectSingleNode("//d:tableStyleInfo", namespaceMgr);
            if (tableStyleInfo != null)
            {
                obj.tableStyleInfo = new CT_TableStyleInfo();
                obj.tableStyleInfo.name = tableStyleInfo.Attributes["name"].Value;
                if (tableStyleInfo.Attributes["showFirstColumn"] != null)
                    obj.tableStyleInfo.showFirstColumn = tableStyleInfo.Attributes["showFirstColumn"].Value == "1" ? true : false;
                if (tableStyleInfo.Attributes["showLastColumn"] != null)
                    obj.tableStyleInfo.showLastColumn = tableStyleInfo.Attributes["showLastColumn"].Value == "1" ? true : false;
                if (tableStyleInfo.Attributes["showRowStripes"] != null)
                    obj.tableStyleInfo.showRowStripes = tableStyleInfo.Attributes["showRowStripes"].Value == "1" ? true : false;
                if (tableStyleInfo.Attributes["showColumnStripes"] != null)
                    obj.tableStyleInfo.showColumnStripes = tableStyleInfo.Attributes["showColumnStripes"].Value == "1" ? true : false;
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
                sw.Write(string.Format(" id=\"{0}\" name=\"{1}\" displayName=\"{2}\"", this.ctTable.id, this.ctTable.name, this.ctTable.displayName));
                sw.Write(string.Format(" ref=\"{0}\"", this.ctTable.@ref));
                sw.Write(string.Format(" tableType=\"{0}\"", this.ctTable.tableType));
                if (this.ctTable.totalsRowCount != 0)
                   sw.Write(string.Format(" totalsRowCount=\"{0}\"", this.ctTable.totalsRowCount));
                sw.Write(string.Format(" totalsRowShown=\"{0}\"", this.ctTable.totalsRowShown));
                sw.Write(">");
                sw.Write("<tableColumns count=\"{0}\">", this.ctTable.tableColumns.count);
                foreach (CT_TableColumn ctTableCol in this.ctTable.tableColumns.tableColumn)
                { 
                    sw.Write(string.Format("<tableColumn id=\"{0}\" name=\"{1}\"", ctTableCol.id, ctTableCol.name));
                    if (ctTableCol.uniqueName != null)
                        sw.Write(string.Format(" uniqueName=\"{0}\"", ctTableCol.uniqueName));
                    if(ctTableCol.totalsRowCellStyle!=null)
                        sw.Write(string.Format(" totalsRowCellStyle=\"{0}\"", ctTableCol.totalsRowCellStyle));
                    if (ctTableCol.totalsRowLabel != null)
                        sw.Write(string.Format(" totalsRowLabel=\"{0}\"", ctTableCol.totalsRowLabel));
                    if (ctTableCol.totalsRowDxfId != 0)
                        sw.Write(string.Format(" totalsRowDxfId=\"{0}\"", ctTableCol.totalsRowDxfId));
                    if (ctTableCol.queryTableFieldId != 0)
                        sw.Write(string.Format(" queryTableFieldId=\"{0}\"", ctTableCol.queryTableFieldId));
                    sw.Write(">");
                    if (ctTableCol.xmlColumnPr != null)
                    {
                        ctTableCol.xmlColumnPr.Write(sw);
                    }
                    sw.Write("</tableColumn>");
                }
                sw.Write("</tableColumns>");
                if (this.ctTable.tableStyleInfo != null)
                {
                    sw.Write(string.Format("<tableStyleInfo name=\"{0}\"", ctTable.tableStyleInfo.name));
                    if (ctTable.tableStyleInfo.showColumnStripes)
                        sw.Write(" showColumnStripes=\"1\"");
                    if (ctTable.tableStyleInfo.showFirstColumn)
                        sw.Write(" showFirstColumn=\"1\"");
                    if (ctTable.tableStyleInfo.showLastColumn)
                        sw.Write(" showLastColumn=\"1\"");
                    if (ctTable.tableStyleInfo.showRowStripes)
                        sw.Write(" showRowStripes=\"1\"");
                    sw.Write("/>");
                }
                sw.Write("</table>");
            }
        }
    }
}
