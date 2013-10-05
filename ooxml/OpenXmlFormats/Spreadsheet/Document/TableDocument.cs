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
                foreach (XmlNode tableCol in tableCols)
                {
                    CT_TableColumn tableColObj = new CT_TableColumn();
                    tableColObj.id = uint.Parse(tableCol.Attributes["id"].Value);
                    tableColObj.name = tableCol.Attributes["name"].Value;
                    obj.tableColumns.tableColumn.Add(tableColObj);
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
            serializer.Serialize(stream, ctTable);
        }
    }
}
