using NPOI.OpenXml4Net.DataVirtualization;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Xml;

namespace NPOI.XSSF.UserModel
{
    public class XSSFRowProvider : IItemsProvider<XSSFRow>
    {
        private const int SHEET_DATA_DEPTH = 1;
        private const string SHEET_DATA_NAME = "sheetData";
        private const int ROW_DEPTH = 2;
        private const string ROW_NAME = "row";

        private static POILogger _logger = POILogFactory.GetLogger(typeof(XSSFRowProvider));

        private PackagePart _packagePart;
        private XSSFSheet _sheet;
        private Stream _stream;
        private int _rowCount;
        private XmlNamespaceManager _xmlNamespaceManager;

        private XmlTextReader _reader;
        private int _lastIndex = 0;//已获取队列最后一个元素的索引号

        public static Stopwatch Stopwatch;

        public XSSFRowProvider(PackagePart packagePart, XSSFSheet sheet, int rowCount, XmlNamespaceManager xmlNamespaceManager)
        {
            _packagePart = packagePart;
            _sheet = sheet;
            _rowCount = rowCount;
            _xmlNamespaceManager = xmlNamespaceManager;
            Stopwatch = new Stopwatch();
        }

        public int FetchCount()
        {
            return _rowCount;
        }

        public IList<XSSFRow> FetchRange(int startIndex, int count)
        {
            Stopwatch.Start();
            var rows = GetRows(startIndex, count);
            Stopwatch.Stop();
            return rows;
        }

        private IList<XSSFRow> GetRows(int startIndex, int count)
        {
            if (_reader == null || startIndex != _lastIndex + 1)
            {
                _stream = _packagePart.GetInputStream();
                CloseReader();
                _reader = new XmlTextReader(_stream);
                _reader.DtdProcessing = DtdProcessing.Ignore;
                _lastIndex = -1;
            }

            List<XSSFRow> rows = new List<XSSFRow>();
            bool isOutSheetData = false;
            int endIndex = startIndex + count - 1;
            try
            {
                while (_reader.Read())
                {
                    if (isOutSheetData)
                    {
                        break;
                    }

                    switch (_reader.NodeType)
                    {
                        case XmlNodeType.XmlDeclaration:
                        case XmlNodeType.Text:
                        case XmlNodeType.Whitespace:
                            break;
                        case XmlNodeType.Element:
                            if (IsRowElement())
                            {
                                if (_lastIndex + 1 >= startIndex && _lastIndex + 1 <= endIndex)
                                {
                                    var ct_row = CT_Row.Parse(_reader, _xmlNamespaceManager);
                                    if (!CheckRowIsEmpty(ct_row))
                                    {
                                        rows.Add(TransferRow(ct_row));
                                        _lastIndex++;
                                    }
                                }
                                if (_lastIndex == _rowCount - 1)
                                {
                                    CloseReader();
                                }
                                if (_lastIndex == endIndex)
                                {
                                    return rows;
                                }
                            }
                            break;
                        case XmlNodeType.EndElement:
                            isOutSheetData = IsOutSheetData();
                            break;
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {

            }

            return rows;
        }

        private void CloseReader()
        {
            if (_reader != null && _reader.ReadState != ReadState.Closed)
                _reader.Close();
        }

        private bool IsRowElement()
        {
            return _reader.Depth == ROW_DEPTH && _reader.Name == ROW_NAME;
        }

        private bool IsOutSheetData()
        {
            return _reader.Depth == SHEET_DATA_DEPTH && _reader.Name == SHEET_DATA_NAME;
        }

        private bool CheckRowIsEmpty(CT_Row ct_row)
        {
            var rst = true;
            foreach (var cell in ct_row.c)
            {
                if (!string.IsNullOrEmpty(cell.v))
                {
                    rst = false;
                }
            }

            return rst;
        }

        private XSSFRow TransferRow(CT_Row ct_row)
        {
            return new XSSFRow(ct_row, _sheet);
        }
    }

}