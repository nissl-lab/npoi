using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

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

            List<CT_Dxf> dxfList = new List<CT_Dxf>();

            // CT* classes don't handle "mc:AlternateContent" elements, so get the Dxf instances manually
            XmlCursor cur = dxfs.newCursor();
            // sometimes there are namespaces sometimes not.
            String xquery = "declare namespace x='" + XSSFRelation.NS_SPREADSHEETML + "' .//x:dxf | .//dxf";
            cur.selectPath(xquery);
            while (cur.toNextSelection())
            {
                XmlObject obj = cur.getObject();
                String parentName = obj.getDomNode().getParentNode().getNodeName();
                // ignore alternate content choices, we won't know anything about their namespaces
                if (parentName.Equals("mc:Fallback") || parentName.Equals("x:dxfs") || parentName.contentEquals("dxfs"))
                {
                    CTDxf dxf;
                    try
                    {
                        if (obj is CTDxf) {
                            dxf = (CTDxf)obj;
                        } else
                        {
                            dxf = CT_Dxf.Factory.parse(obj.newXMLStreamReader(), new XmlOptions().setDocumentType(CTDxf.type));
                        }
                        if (dxf != null) dxfList.Add(dxf);
                    }
                    catch (XmlException e)
                    {
                        logger.log(POILogger.WARN, "Error parsing XSSFTableStyle", e);
                    }
                }
            }

            foreach (CT_TableStyleElement element in tableStyle.tableStyleElement)
            {
                TableStyleType type = TableStyleType.valueOf(element.getType().toString());
                DifferentialStyleProvider dstyle = null;
                if (element.isSetDxfId())
                {
                    int idx = (int)element.getDxfId();
                    CT_Dxf dxf;
                    dxf = dxfList.get(idx);
                    int stripeSize = 0;
                    if (element.isSetSize()) stripeSize = (int)element.getSize();
                    if (dxf != null) dstyle = new XSSFDxfStyleProvider(dxf, stripeSize, colorMap);
                }
                elementMap.put(type, dstyle);
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

        public DifferentialStyleProvider GetStyle(TableStyleType type)
        {
            return elementMap[type];
        }
    }
}
