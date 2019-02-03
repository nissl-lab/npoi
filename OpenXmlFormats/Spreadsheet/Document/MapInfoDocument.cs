using NPOI.OpenXml4Net.Util;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class MapInfoDocument
    {
        CT_MapInfo map = null;

        public MapInfoDocument()
        { 
        }
        public MapInfoDocument(CT_MapInfo map)
        {
            this.map = map;
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
            using (StreamWriter sw = new StreamWriter(stream))
            {
                sw.Write("<MapInfo xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" SelectionNamespaces=\"xmlns:ns1='http://schemas.openxmlformats.org/spreadsheetml/2006/main'\">");
                if (this.map.Schema != null)
                {
                    foreach (CT_Schema ctSchema in this.map.Schema)
                    {
                        sw.Write(string.Format("<Schema ID=\"{0}\"", ctSchema.ID));
                        if (ctSchema.Namespace != null)
                            sw.Write(string.Format(" Namespace=\"{0}\"", ctSchema.Namespace));
                        if (ctSchema.SchemaRef != null)
                            sw.Write(string.Format(" SchemaRef=\"{0}\"", ctSchema.SchemaRef));
                        sw.Write(">");
                        sw.Write(ctSchema.InnerXml);
                        sw.Write("</Schema>");
                    }
                }
                if (this.map.Map != null)
                {
                    foreach (CT_Map ctMap in this.map.Map)
                    {
                        sw.Write(string.Format("<Map ID=\"{0}\"", ctMap.ID));
                        if (ctMap.SchemaID != null)
                            sw.Write(string.Format(" SchemaID=\"{0}\"", ctMap.SchemaID));
                        if (ctMap.RootElement != null)
                            sw.Write(string.Format(" RootElement=\"{0}\"", ctMap.RootElement));
                        if (ctMap.Name != null)
                            sw.Write(string.Format(" Name=\"{0}\"", ctMap.Name));
                        if (ctMap.PreserveFormat)
                            sw.Write(" PreserveFormat=\"true\"");
                        if (ctMap.PreserveSortAFLayout)
                            sw.Write(" PreserveSortAFLayout=\"true\"");
                        if (ctMap.ShowImportExportValidationErrors)
                            sw.Write(" ShowImportExportValidationErrors=\"true\"");
                        if (ctMap.Append)
                            sw.Write(" Append=\"true\"");
                        if (ctMap.AutoFit)
                            sw.Write(" AutoFit=\"true\"");
                        sw.Write(" />");
                    }
                }
                sw.Write("</MapInfo");
            }
        }

        public static MapInfoDocument Parse(System.Xml.XmlDocument xmldoc, System.Xml.XmlNamespaceManager NameSpaceManager)
        {
            MapInfoDocument doc = new MapInfoDocument();
            doc.map = new CT_MapInfo();
            doc.map.Map = new System.Collections.Generic.List<CT_Map>();
            foreach (XmlElement mapNode in xmldoc.SelectNodes("d:MapInfo/d:Map", NameSpaceManager))
            { 
                CT_Map ctMap=new CT_Map();
                ctMap.ID = XmlHelper.ReadUInt(mapNode.GetAttributeNode("ID"));
                ctMap.Name = XmlHelper.ReadString(mapNode.GetAttributeNode("Name"));
                ctMap.RootElement = XmlHelper.ReadString(mapNode.GetAttributeNode("RootElement"));
                ctMap.SchemaID = XmlHelper.ReadString(mapNode.GetAttributeNode("SchemaID"));
                ctMap.ShowImportExportValidationErrors = XmlHelper.ReadBool(mapNode.GetAttributeNode("ShowImportExportValidationErrors"));
                ctMap.PreserveFormat = XmlHelper.ReadBool(mapNode.GetAttributeNode("PreserveFormat"));
                ctMap.PreserveSortAFLayout = XmlHelper.ReadBool(mapNode.GetAttributeNode("PreserveSortAFLayout"));
                ctMap.Append = XmlHelper.ReadBool(mapNode.GetAttributeNode("Append"));
                ctMap.AutoFit = XmlHelper.ReadBool(mapNode.GetAttributeNode("AutoFit"));
                doc.map.Map.Add(ctMap);
            }
            doc.map.Schema = new System.Collections.Generic.List<CT_Schema>();
            foreach (XmlNode schemaNode in xmldoc.SelectNodes("d:MapInfo/d:Schema", NameSpaceManager))
            {
                CT_Schema ctSchema = new CT_Schema();
                ctSchema.ID = schemaNode.Attributes["ID"].Value;
                if (schemaNode.Attributes["Namespace"] != null)
                    ctSchema.Namespace = schemaNode.Attributes["Namespace"].Value;
                if (schemaNode.Attributes["SchemaRef"] != null)
                    ctSchema.SchemaRef = schemaNode.Attributes["SchemaRef"].Value;
                ctSchema.InnerXml = schemaNode.InnerXml;
                doc.map.Schema.Add(ctSchema);
            }
            return doc;
        }
    }
}
