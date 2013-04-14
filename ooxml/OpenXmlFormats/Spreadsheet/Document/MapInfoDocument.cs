using System.IO;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class MapInfoDocument
    {
        static XmlSerializer serializer = new XmlSerializer(typeof(CT_MapInfo));
        CT_MapInfo map = null;

        public MapInfoDocument()
        { 
        }
        public MapInfoDocument(CT_MapInfo map)
        {
            this.map = map;
        }
        public static MapInfoDocument Parse(Stream stream)
        {
            CT_MapInfo obj = (CT_MapInfo)serializer.Deserialize(stream);
            return new MapInfoDocument(obj);
        }
        public CT_MapInfo GetMapInfo()
        {
            return this.map;
        }
        public void SetMapInfo(CT_MapInfo map)
        {
            this.map = map;
        }
        public void SetComments(CT_MapInfo map)
        {
            this.map = map;
        }
        public void Save(Stream stream)
        {
            serializer.Serialize(stream, map);
        }
    }
}
