using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using EnumsNET;

namespace NPOI.OOXML.XSSF.UserModel
{
    public class XSSFTableStyle : ITableStyle
    {
        private String name;
        private int index;
        private Dictionary<TableStyleType, IDifferentialStyleProvider> elementMap = new Dictionary<TableStyleType, IDifferentialStyleProvider>();

        public XSSFTableStyle(int index, CT_Dxfs dxfs, CT_TableStyle tableStyle, IIndexedColorMap colorMap)
        {
            this.name = tableStyle.name;
            this.index = index;

            List<CT_Dxf> dxfList = dxfs.dxf;

            foreach (CT_TableStyleElement element in tableStyle.tableStyleElement)
            {
                TableStyleType type = Enums.Parse<TableStyleType>(element.type.GetName());
                IDifferentialStyleProvider dstyle = null;
                if (element.dxfIdSpecified)
                {
                    int idx = (int)element.dxfId;
                    CT_Dxf dxf = dxfList[idx];
                    int stripeSize = 0;
                    if (element.size!=0) stripeSize = (int)element.size;
                    if (dxf != null) dstyle = new XSSFDxfStyleProvider(dxf, stripeSize, colorMap);
                }
                elementMap.Add(type, dstyle);
            }
        }
        public string Name
        {
            get { return name; }
        }

        public int Index
        {
            get { return index; }
        }

        public bool IsBuiltin
        {
            get { return false; }
        }

        public IDifferentialStyleProvider GetStyle(TableStyleType type)
        {
            return elementMap.TryGetValue(type, out IDifferentialStyleProvider value) ? value : null;
        }
    }
}
