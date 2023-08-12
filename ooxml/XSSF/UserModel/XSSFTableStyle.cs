using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;

namespace NPOI.OOXML.XSSF.UserModel
{
    public class XSSFTableStyle : ITableStyle
    {
        private String name;
        private int index;
        private Dictionary<TableStyleType, DifferentialStyleProvider> elementMap = new Dictionary<TableStyleType, DifferentialStyleProvider>();

        public XSSFTableStyle(int index, CT_Dxfs dxfs, CT_TableStyle tableStyle, IIndexedColorMap colorMap)
        {
            this.name = tableStyle.name;
            this.index = index;

            List<CT_Dxf> dxfList = dxfs.dxf;

            foreach (CT_TableStyleElement element in tableStyle.tableStyleElement)
            {
                TableStyleType type = Translate(element.type);
                DifferentialStyleProvider dstyle = null;
                if (element.dxfIdSpecified)
                {
                    int idx = (int)element.dxfId;
                    CT_Dxf dxf;
                    dxf = dxfList[idx];
                    int stripeSize = 0;
                    if (element.size != 0) stripeSize = (int)element.size;
                    if (dxf != null) dstyle = new XSSFDxfStyleProvider(dxf, stripeSize, colorMap);
                }
                elementMap.Add(type, dstyle);
            }
        }

        private static TableStyleType Translate(ST_TableStyleType type) => type switch
        {
            ST_TableStyleType.wholeTable => TableStyleType.wholeTable,
            ST_TableStyleType.headerRow => TableStyleType.headerRow,
            ST_TableStyleType.totalRow => TableStyleType.totalRow,
            ST_TableStyleType.firstColumn => TableStyleType.firstColumn,
            ST_TableStyleType.lastColumn => TableStyleType.lastColumn,
            ST_TableStyleType.firstRowStripe => TableStyleType.firstRowStripe,
            ST_TableStyleType.secondRowStripe => TableStyleType.secondRowStripe,
            ST_TableStyleType.firstColumnStripe => TableStyleType.firstColumnStripe,
            ST_TableStyleType.secondColumnStripe => TableStyleType.secondColumnStripe,
            ST_TableStyleType.firstHeaderCell => TableStyleType.firstHeaderCell,
            ST_TableStyleType.lastHeaderCell => TableStyleType.lastHeaderCell,
            ST_TableStyleType.firstTotalCell => TableStyleType.firstTotalCell,
            ST_TableStyleType.lastTotalCell => TableStyleType.lastTotalCell,
            ST_TableStyleType.firstSubtotalColumn => TableStyleType.firstSubtotalColumn,
            ST_TableStyleType.secondSubtotalColumn => TableStyleType.secondSubtotalColumn,
            ST_TableStyleType.thirdSubtotalColumn => TableStyleType.thirdSubtotalColumn,
            ST_TableStyleType.firstSubtotalRow => TableStyleType.firstSubtotalRow,
            ST_TableStyleType.secondSubtotalRow => TableStyleType.secondSubtotalRow,
            ST_TableStyleType.thirdSubtotalRow => TableStyleType.thirdSubtotalRow,
            ST_TableStyleType.blankRow => TableStyleType.blankRow,
            ST_TableStyleType.firstColumnSubheading => TableStyleType.firstColumnSubheading,
            ST_TableStyleType.secondColumnSubheading => TableStyleType.secondColumnSubheading,
            ST_TableStyleType.thirdColumnSubheading => TableStyleType.thirdColumnSubheading,
            ST_TableStyleType.firstRowSubheading => TableStyleType.firstRowSubheading,
            ST_TableStyleType.secondRowSubheading => TableStyleType.secondRowSubheading,
            ST_TableStyleType.thirdRowSubheading => TableStyleType.thirdRowSubheading,
            ST_TableStyleType.pageFieldLabels => TableStyleType.pageFieldLabels,
            ST_TableStyleType.pageFieldValues => TableStyleType.pageFieldValues,
            _ => throw new ArgumentException()
        };
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

        public DifferentialStyleProvider GetStyle(TableStyleType type)
        {
            return elementMap[type];
        }
    }
}
