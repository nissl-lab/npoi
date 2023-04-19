using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.Model;
using System;

namespace NPOI.OOXML.XSSF.UserModel
{
    public class XSSFPivotTableStyleInfo : IPivotTableStyleInfo
    {
        private readonly CT_PivotTableStyle _pivotStyle;
        private readonly StylesTable _stylesTable;

        private bool _showColStripes;
        private bool _showRowStripes;
        private bool _showColHeaders;
        private bool _showRowHeaders;
        private bool _showLastColumn;

        public XSSFPivotTableStyleInfo(StylesTable stylesTable, CT_PivotTableStyle pivotTableStyle)
        {
            _showColStripes = pivotTableStyle.showColStripes;
            _showRowStripes = pivotTableStyle.showRowStripes;
            _showColHeaders = pivotTableStyle.showColHeaders;
            _showRowHeaders = pivotTableStyle.showRowHeaders;
            _showLastColumn = pivotTableStyle.showLastColumn;
            _stylesTable = stylesTable;
            _pivotStyle = pivotTableStyle;
            Style = stylesTable.GetTableStyle(pivotTableStyle.name);
        }

        public bool IsShowColumnStripes 
        {
            get { 
                return _showColStripes; 
            }

            set {
                _showColStripes = value;
                _pivotStyle.showColStripes = value;
            }
        }

        public bool IsShowRowStripes 
        {
            get
            {
                return _showRowStripes;
            }

            set
            {
                _showRowStripes = value;
                _pivotStyle.showRowStripes = value;
            }
        }

        public bool IsShowColumnHeaders
        {
            get
            {
                return _showColHeaders;
            }

            set
            {
                _showColHeaders = value;
                _pivotStyle.showColHeaders = value;
            }
        }

        public bool IsShowRowHeaders
        {
            get
            {
                return _showRowHeaders;
            }

            set
            {
                _showRowHeaders = value;
                _pivotStyle.showRowHeaders = value;
            }
        }

        public bool IsShowLastColumn
        {
            get
            {
                return _showLastColumn;
            }

            set
            {
                _showLastColumn = value;
                _pivotStyle.showLastColumn = value;
            }
        }

        public bool IsShowFirstColumn
        {
            get
            {
                return true;
            }

            set => throw new NotSupportedException();
        }

        public string Name
        {
            get
            {
                return _pivotStyle.name;
            }

            set
            {
                _pivotStyle.name = value;
                Style = _stylesTable.GetTableStyle(value);
            }
        }

        public ITableStyle Style { get; private set; }
    }
}
