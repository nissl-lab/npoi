using NPOI.OpenXml4Net.Util;
using System;
using System.Linq;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough()]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot("table", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public class CT_Table
    {
        public CT_Table()
        {
            //this.extLstField = new CT_ExtensionList();
            //this.tableStyleInfoField = new CT_TableStyleInfo();
            //this.tableColumnsField = new CT_TableColumns();
            //this.sortStateField = new CT_SortState();
            //this.autoFilterField = new CT_AutoFilter();
            this.tableType = ST_TableType.worksheet;
            this.headerRowCount = ((uint)(1));
            this.insertRow = false;
            this.insertRowShift = false;
            this.totalsRowCount = ((uint)(0));
            this.totalsRowShown = true;
            this.published = false;
        }
        public static CT_Table Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Table ctObj = new CT_Table();
            if (node.Attributes[nameof(id)] != null)
                ctObj.id = XmlHelper.ReadUInt(node.Attributes[nameof(id)]);
            ctObj.name = XmlHelper.ReadString(node.Attributes[nameof(name)]);
            ctObj.displayName = XmlHelper.ReadString(node.Attributes[nameof(displayName)]);
            ctObj.comment = XmlHelper.ReadString(node.Attributes[nameof(comment)]);
            ctObj.@ref = XmlHelper.ReadString(node.Attributes[nameof(@ref)]);
            if (node.Attributes[nameof(tableType)] != null)
                ctObj.tableType = (ST_TableType)Enum.Parse(typeof(ST_TableType), node.Attributes[nameof(tableType)].Value);
            if (node.Attributes[nameof(headerRowCount)] != null)
                ctObj.headerRowCount = XmlHelper.ReadUInt(node.Attributes[nameof(headerRowCount)]);
            if (node.Attributes[nameof(insertRow)] != null)
                ctObj.insertRow = XmlHelper.ReadBool(node.Attributes[nameof(insertRow)]);
            if (node.Attributes[nameof(insertRowShift)] != null)
                ctObj.insertRowShift = XmlHelper.ReadBool(node.Attributes[nameof(insertRowShift)]);
            if (node.Attributes[nameof(totalsRowCount)] != null)
                ctObj.totalsRowCount = XmlHelper.ReadUInt(node.Attributes[nameof(totalsRowCount)]);
            if (node.Attributes[nameof(totalsRowShown)] != null)
                ctObj.totalsRowShown = XmlHelper.ReadBool(node.Attributes[nameof(totalsRowShown)]);
            if (node.Attributes[nameof(published)] != null)
                ctObj.published = XmlHelper.ReadBool(node.Attributes[nameof(published)]);
            if (node.Attributes[nameof(headerRowDxfId)] != null)
                ctObj.headerRowDxfId = XmlHelper.ReadUInt(node.Attributes[nameof(headerRowDxfId)]);
            if (node.Attributes[nameof(dataDxfId)] != null)
                ctObj.dataDxfId = XmlHelper.ReadUInt(node.Attributes[nameof(dataDxfId)]);
            if (node.Attributes[nameof(totalsRowDxfId)] != null)
                ctObj.totalsRowDxfId = XmlHelper.ReadUInt(node.Attributes[nameof(totalsRowDxfId)]);
            if (node.Attributes[nameof(headerRowBorderDxfId)] != null)
                ctObj.headerRowBorderDxfId = XmlHelper.ReadUInt(node.Attributes[nameof(headerRowBorderDxfId)]);
            if (node.Attributes[nameof(tableBorderDxfId)] != null)
                ctObj.tableBorderDxfId = XmlHelper.ReadUInt(node.Attributes[nameof(tableBorderDxfId)]);
            if (node.Attributes[nameof(totalsRowBorderDxfId)] != null)
                ctObj.totalsRowBorderDxfId = XmlHelper.ReadUInt(node.Attributes[nameof(totalsRowBorderDxfId)]);
            ctObj.headerRowCellStyle = XmlHelper.ReadString(node.Attributes[nameof(headerRowCellStyle)]);
            ctObj.dataCellStyle = XmlHelper.ReadString(node.Attributes[nameof(dataCellStyle)]);
            ctObj.totalsRowCellStyle = XmlHelper.ReadString(node.Attributes[nameof(totalsRowCellStyle)]);
            if (node.Attributes[nameof(connectionId)] != null)
                ctObj.connectionId = XmlHelper.ReadUInt(node.Attributes[nameof(connectionId)]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(autoFilter))
                    ctObj.autoFilter = CT_AutoFilter.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(sortState))
                    ctObj.sortState = CT_SortState.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(tableColumns))
                    ctObj.tableColumns = CT_TableColumns.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(tableStyleInfo))
                    ctObj.tableStyleInfo = CT_TableStyleInfo.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(extLst))
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw)
        {

            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write("<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
            XmlHelper.WriteAttribute(sw, nameof(id), this.id);
            XmlHelper.WriteAttribute(sw, nameof(name), this.name);
            XmlHelper.WriteAttribute(sw, nameof(displayName), this.displayName);
            XmlHelper.WriteAttribute(sw, nameof(comment), this.comment);
            XmlHelper.WriteAttribute(sw, nameof(@ref), this.@ref);
            XmlHelper.WriteAttribute(sw, nameof(tableType), this.tableType.ToString());
            XmlHelper.WriteAttribute(sw, nameof(headerRowCount), this.headerRowCount);
            XmlHelper.WriteAttribute(sw, nameof(insertRow), this.insertRow);
            XmlHelper.WriteAttribute(sw, nameof(insertRowShift), this.insertRowShift);
            XmlHelper.WriteAttribute(sw, nameof(totalsRowCount), this.totalsRowCount);
            XmlHelper.WriteAttribute(sw, nameof(totalsRowShown), this.totalsRowShown);
            XmlHelper.WriteAttribute(sw, nameof(published), this.published);
            XmlHelper.WriteAttribute(sw, nameof(headerRowDxfId), this.headerRowDxfId);
            XmlHelper.WriteAttribute(sw, nameof(dataDxfId), this.dataDxfId);
            XmlHelper.WriteAttribute(sw, nameof(totalsRowDxfId), this.totalsRowDxfId);
            XmlHelper.WriteAttribute(sw, nameof(headerRowBorderDxfId), this.headerRowBorderDxfId);
            XmlHelper.WriteAttribute(sw, nameof(tableBorderDxfId), this.tableBorderDxfId);
            XmlHelper.WriteAttribute(sw, nameof(totalsRowBorderDxfId), this.totalsRowBorderDxfId);
            XmlHelper.WriteAttribute(sw, nameof(headerRowCellStyle), this.headerRowCellStyle);
            XmlHelper.WriteAttribute(sw, nameof(dataCellStyle), this.dataCellStyle);
            XmlHelper.WriteAttribute(sw, nameof(totalsRowCellStyle), this.totalsRowCellStyle);
            XmlHelper.WriteAttribute(sw, nameof(connectionId), this.connectionId);
            sw.Write(">");
            this.autoFilter?.Write(sw, nameof(autoFilter));
            this.sortState?.Write(sw, nameof(sortState));
            this.tableColumns?.Write(sw, nameof(tableColumns));
            this.tableStyleInfo?.Write(sw, nameof(tableStyleInfo));
            this.extLst?.Write(sw, nameof(extLst));
            sw.Write("</table>");
        }

        [XmlElement]
        public CT_AutoFilter autoFilter { get; set; }
        [XmlElement]
        public CT_SortState sortState { get; set; }
        [XmlElement]
        public CT_TableColumns tableColumns { get; set; }
        [XmlElement]
        public CT_TableStyleInfo tableStyleInfo { get; set; }
        [XmlElement]
        public CT_ExtensionList extLst { get; set; }
        [XmlAttribute]
        public uint id { get; set; }
        [XmlAttribute]
        public string name { get; set; }
        [XmlAttribute]
        public string displayName { get; set; }
        [XmlAttribute]
        public string comment { get; set; }
        [XmlAttribute]
        public string @ref { get; set; }
        [XmlAttribute]
        [DefaultValue(ST_TableType.worksheet)]
        public ST_TableType tableType { get; set; }
        [XmlAttribute]
        [DefaultValue(typeof(uint), "1")]
        public uint headerRowCount { get; set; }
        [XmlAttribute]
        [DefaultValue(false)]
        public bool insertRow { get; set; }
        [XmlAttribute]
        [DefaultValue(false)]
        public bool insertRowShift { get; set; }
        [XmlAttribute]
        [DefaultValue(typeof(uint), "0")]
        public uint totalsRowCount { get; set; }
        [XmlAttribute]
        [DefaultValue(true)]
        public bool totalsRowShown { get; set; }
        [XmlAttribute]
        [DefaultValue(false)]
        public bool published { get; set; }
        [XmlAttribute]
        public uint headerRowDxfId { get; set; }
        [XmlIgnore]
        public bool headerRowDxfIdSpecified { get; set; }
        [XmlAttribute]
        public uint dataDxfId { get; set; }

        [XmlIgnore]
        public bool dataDxfIdSpecified { get; set; }
        [XmlAttribute]
        public uint totalsRowDxfId { get; set; }

        [XmlIgnore]
        public bool totalsRowDxfIdSpecified { get; set; }
        [XmlAttribute]
        public uint headerRowBorderDxfId { get; set; }

        [XmlIgnore]
        public bool headerRowBorderDxfIdSpecified { get; set; }
        [XmlAttribute]
        public uint tableBorderDxfId { get; set; }

        [XmlIgnore]
        public bool tableBorderDxfIdSpecified { get; set; }
        [XmlAttribute]
        public uint totalsRowBorderDxfId { get; set; }

        [XmlIgnore]
        public bool totalsRowBorderDxfIdSpecified { get; set; }
        [XmlAttribute]
        public string headerRowCellStyle { get; set; }
        [XmlAttribute]
        public string dataCellStyle { get; set; }
        [XmlAttribute]
        public string totalsRowCellStyle { get; set; }
        [XmlAttribute]
        public uint connectionId { get; set; }

        [XmlIgnore]
        public bool connectionIdSpecified { get; set; }
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_TableColumns
    {
        public CT_TableColumns()
        {
            //this.tableColumnField = new List<CT_TableColumn>();
        }
        public static CT_TableColumns Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TableColumns ctObj = new CT_TableColumns();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.tableColumn = new List<CT_TableColumn>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(tableColumn))
                    ctObj.tableColumn.Add(CT_TableColumn.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            tableColumn?.ForEach(x => x.Write(sw, nameof(tableColumn)));
            sw.Write($"</{nodeName}>");
        }

        [XmlElement]
        public List<CT_TableColumn> tableColumn { get; set; }
        [XmlAttribute]
        public uint count { get; set; }

        [XmlIgnore]
        public bool countSpecified { get; set; }
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_TableColumn
    {
        private string totalsRowCellStyleField;

        public CT_TableColumn()
        {
            //this.extLstField = new CT_ExtensionList();
            //this.xmlColumnPrField = new CT_XmlColumnPr();
            //this.totalsRowFormulaField = new CT_TableFormula();
            //this.calculatedColumnFormulaField = new CT_TableFormula();
            this.totalsRowFunction = ST_TotalsRowFunction.none;
        }
        public static CT_TableColumn Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TableColumn ctObj = new CT_TableColumn();
            if (node.Attributes[nameof(id)] != null)
                ctObj.id = XmlHelper.ReadUInt(node.Attributes[nameof(id)]);
            ctObj.uniqueName = XmlHelper.ReadString(node.Attributes[nameof(uniqueName)]);
            ctObj.name = XmlHelper.ReadString(node.Attributes[nameof(name)]);
            if (node.Attributes[nameof(totalsRowFunction)] != null)
                ctObj.totalsRowFunction = (ST_TotalsRowFunction)Enum.Parse(typeof(ST_TotalsRowFunction), node.Attributes[nameof(totalsRowFunction)].Value);
            ctObj.totalsRowLabel = XmlHelper.ReadString(node.Attributes[nameof(totalsRowLabel)]);
            if (node.Attributes[nameof(queryTableFieldId)] != null)
                ctObj.queryTableFieldId = XmlHelper.ReadUInt(node.Attributes[nameof(queryTableFieldId)]);
            if (node.Attributes[nameof(headerRowDxfId)] != null)
                ctObj.headerRowDxfId = XmlHelper.ReadUInt(node.Attributes[nameof(headerRowDxfId)]);
            if (node.Attributes[nameof(dataDxfId)] != null)
                ctObj.dataDxfId = XmlHelper.ReadUInt(node.Attributes[nameof(dataDxfId)]);
            if (node.Attributes[nameof(totalsRowDxfId)] != null)
                ctObj.totalsRowDxfId = XmlHelper.ReadUInt(node.Attributes[nameof(totalsRowDxfId)]);
            ctObj.headerRowCellStyle = XmlHelper.ReadString(node.Attributes[nameof(headerRowCellStyle)]);
            ctObj.dataCellStyle = XmlHelper.ReadString(node.Attributes[nameof(dataCellStyle)]);
            ctObj.totalsRowCellStyle = XmlHelper.ReadString(node.Attributes[nameof(totalsRowCellStyle)]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(calculatedColumnFormula))
                    ctObj.calculatedColumnFormula = CT_TableFormula.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(totalsRowFormula))
                    ctObj.totalsRowFormula = CT_TableFormula.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(xmlColumnPr))
                    ctObj.xmlColumnPr = CT_XmlColumnPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(extLst))
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");
            XmlHelper.WriteAttribute(sw, nameof(id), this.id);
            XmlHelper.WriteAttribute(sw, nameof(uniqueName), this.uniqueName);
            XmlHelper.WriteAttribute(sw, nameof(name), this.name);
            XmlHelper.WriteAttribute(sw, nameof(totalsRowFunction), this.totalsRowFunction.ToString());
            XmlHelper.WriteAttribute(sw, nameof(totalsRowLabel), this.totalsRowLabel);
            XmlHelper.WriteAttribute(sw, nameof(queryTableFieldId), this.queryTableFieldId);
            XmlHelper.WriteAttribute(sw, nameof(headerRowDxfId), this.headerRowDxfId);
            XmlHelper.WriteAttribute(sw, nameof(dataDxfId), this.dataDxfId);
            XmlHelper.WriteAttribute(sw, nameof(totalsRowDxfId), this.totalsRowDxfId);
            XmlHelper.WriteAttribute(sw, nameof(headerRowCellStyle), this.headerRowCellStyle);
            XmlHelper.WriteAttribute(sw, nameof(dataCellStyle), this.dataCellStyle);
            XmlHelper.WriteAttribute(sw, nameof(totalsRowCellStyle), this.totalsRowCellStyle);
            sw.Write(">");
            this.calculatedColumnFormula?.Write(sw, nameof(calculatedColumnFormula));
            this.totalsRowFormula?.Write(sw, nameof(totalsRowFormula));
            this.xmlColumnPr?.Write(sw, nameof(xmlColumnPr));
            this.extLst?.Write(sw, nameof(extLst));
            sw.Write($"</{nodeName}>");
        }

        [XmlElement]
        public CT_TableFormula calculatedColumnFormula { get; set; }
        [XmlElement]
        public CT_TableFormula totalsRowFormula { get; set; }
        public CT_XmlColumnPr xmlColumnPr { get; set; }
        [XmlElement]
        public CT_ExtensionList extLst { get; set; }
        [XmlAttribute]
        public uint id { get; set; }
        [XmlAttribute]
        public string uniqueName { get; set; }
        [XmlAttribute]
        public string name { get; set; }
        [XmlAttribute]
        [DefaultValue(ST_TotalsRowFunction.none)]
        public ST_TotalsRowFunction totalsRowFunction { get; set; }
        [XmlAttribute]
        public string totalsRowLabel { get; set; }
        [XmlAttribute]
        public uint queryTableFieldId { get; set; }

        [XmlIgnore]
        public bool queryTableFieldIdSpecified { get; set; }
        [XmlAttribute]
        public uint headerRowDxfId { get; set; }

        [XmlIgnore]
        public bool headerRowDxfIdSpecified { get; set; }
        [XmlAttribute]
        public uint dataDxfId { get; set; }

        [XmlIgnore]
        public bool dataDxfIdSpecified { get; set; }
        [XmlAttribute]
        public uint totalsRowDxfId { get; set; }

        [XmlIgnore]
        public bool totalsRowDxfIdSpecified { get; set; }
        [XmlAttribute]
        public string headerRowCellStyle { get; set; }
        [XmlAttribute]
        public string dataCellStyle { get; set; }
        [XmlAttribute]
        public string totalsRowCellStyle
        {
            get
            {
                return this.totalsRowCellStyleField;
            }
            set
            {
                this.totalsRowCellStyleField = value;
            }
        }
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_TableFormula
    {
        public CT_TableFormula()
        {
            this.array = false;
        }
        [XmlAttribute]
        [DefaultValue(false)]
        public bool array { get; set; }

        [XmlText]
        public string Value { get; set; }

        public static CT_TableFormula Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TableFormula ctObj = new CT_TableFormula();
            if (node.Attributes[nameof(array)] != null)
                ctObj.array = XmlHelper.ReadBool(node.Attributes[nameof(array)]);
            ctObj.Value = node.InnerText;
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");
            XmlHelper.WriteAttribute(sw, nameof(array), this.array);
            sw.Write(">");
            sw.Write(XmlHelper.EncodeXml(this.Value));
            sw.Write($"</{nodeName}>");
        }

    }
}
