using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.Model;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.OOXML.XSSF.UserModel
{
    public class XSSFTableStyleInfo : ITableStyleInfo
    {
        private CT_TableStyleInfo styleInfo;
        private StylesTable stylesTable;
        private ITableStyle style;
        private bool columnStripes;
        private bool rowStripes;
        private bool firstColumn;
        private bool lastColumn;

        public XSSFTableStyleInfo(StylesTable stylesTable, CT_TableStyleInfo tableStyleInfo)
        {
            this.columnStripes = tableStyleInfo.showColumnStripes;
            this.rowStripes = tableStyleInfo.showRowStripes;
            this.firstColumn = tableStyleInfo.showFirstColumn;
            this.lastColumn = tableStyleInfo.showLastColumn;
            this.style = stylesTable.GetTableStyle(tableStyleInfo.name);
            this.stylesTable = stylesTable;
            this.styleInfo = tableStyleInfo;
        }
        public bool IsShowColumnStripes
        {
            get { return columnStripes; }
            set {
                this.columnStripes = value;
                styleInfo.showColumnStripes=value;
            }
        }

        public bool IsShowRowStripes
        {
            get { return rowStripes; }
            set {
                this.rowStripes = value;
                styleInfo.showRowStripes = value;
            }
        }

        public bool IsShowFirstColumn
        {
            get { return firstColumn; }
            set
            {
                this.firstColumn = value;
                styleInfo.showFirstColumn = value;
            }
        }

        public bool IsShowLastColumn
        {
            get { return lastColumn; }
            set {
                this.lastColumn = value;
                styleInfo.showLastColumn = value;
            }
        }

        public string Name
        {
            get { return style.Name; }
            set {
                styleInfo.name = value;
                style = stylesTable.GetTableStyle(value);
            }
        }

        public ITableStyle Style
        {
            get { return style; }
        }
    }
}
