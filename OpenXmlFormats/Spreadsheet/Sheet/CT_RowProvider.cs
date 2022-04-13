using NPOI.OpenXml4Net.DataVirtualization;
using System;
using System.Collections.Generic;
using System.Xml;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class CT_RowProvider : IItemsProvider<CT_Row>
    {
        internal string sheetPath;
        internal XmlNamespaceManager namespaceManager;

        public int FetchCount()
        {
            using (XmlReader reader = XmlReader.Create(sheetPath))
            {
                int count = 0;
                while (reader.Read())
                {
                    //Process only the elements
                    if (reader.NodeType == XmlNodeType.Element && reader.Name == "row")
                    {
                        count++;
                    }
                }
                return count;
            }
        }

        public IList<CT_Row> FetchRange(int startIndex, int count)
        {
            using (XmlReader reader = XmlReader.Create(sheetPath))
            {
                int num = 0;
                var list = new List<CT_Row>();
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        if (reader.Name == "row")
                        {
                            if (num >= startIndex && num < (startIndex + count))
                            {
                                var row = CT_Row.Parse(reader, namespaceManager);
                                list.Add(row);
                            }
                        }
                    }
                }
                return list;
            }
        }
    }
}
