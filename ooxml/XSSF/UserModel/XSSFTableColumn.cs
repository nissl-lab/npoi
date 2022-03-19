using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.XSSF.UserModel.Helpers;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.XSSF.UserModel
{
    public class XSSFTableColumn
    {
        private XSSFTable table;
        private CT_TableColumn ctTableColumn;
        private XSSFXmlColumnPr xmlColumnPr;
        internal XSSFTableColumn(XSSFTable table, CT_TableColumn ctTableColumn)
        {
            this.table = table;
            this.ctTableColumn = ctTableColumn;
        }
        /// <summary>
        /// Get the table which contains this column
        /// </summary>
        /// <returns>the table containing this column</returns>
        public XSSFTable GetTable()
        {
            return table;
        }
        public long Id { 
            get {
                return ctTableColumn.id;
            }
            set 
            {
                ctTableColumn.id = (uint)value;
            }
        }
        public string Name
        {
            get { return ctTableColumn.name; }
            set { ctTableColumn.name = value; }
        }
        public XSSFXmlColumnPr GetXmlColumnPr()
        {
            if (xmlColumnPr == null)
            {
                CT_XmlColumnPr ctXmlColumnPr = ctTableColumn.xmlColumnPr;
                if (ctXmlColumnPr != null)
                {
                    xmlColumnPr = new XSSFXmlColumnPr(this, ctXmlColumnPr);
                }
            }
            return xmlColumnPr;
        }
        public int ColumnIndex
        {
            get { return table.FindColumnIndex(this.Name); }
        }
    }
}
