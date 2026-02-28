using NPOI.OpenXml4Net.Util;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class MapInfoDocument
    {
        CT_MapInfo mapInfo = null;

        public MapInfoDocument()
        { 
        }
        public MapInfoDocument(CT_MapInfo map)
        {
            this.mapInfo = map;
        }
        public CT_MapInfo GetMapInfo()
        {
            return this.mapInfo;
        }
        public void SetMapInfo(CT_MapInfo map)
        {
            this.mapInfo = map;
        }
        public void SetComments(CT_MapInfo map)
        {
            this.mapInfo = map;
        }
        public string SelectionNamespaces { get; set; }
        public void Save(Stream stream)
        {
            using (StreamWriter sw = new StreamWriter(stream))
            {
                mapInfo.Write(sw, "MapInfo");
            }
        }

        public static MapInfoDocument Parse(XmlDocument xmlDoc, XmlNamespaceManager nameSpaceManager)
        {
            MapInfoDocument doc = new MapInfoDocument();
            doc.mapInfo = CT_MapInfo.Parse(xmlDoc.DocumentElement, nameSpaceManager);
            return doc;
        }
    }
}
