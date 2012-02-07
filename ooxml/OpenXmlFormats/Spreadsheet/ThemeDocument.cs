using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using NPOI.OpenXmlFormats.Dml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class ThemeDocument
    {
        static XmlSerializer serializer = new XmlSerializer(typeof(CT_OfficeStyleSheet));
        public ThemeDocument()
        { 
        
        }
        CT_OfficeStyleSheet stylesheet;
        public ThemeDocument(CT_OfficeStyleSheet stylesheet)
        {
            this.stylesheet = stylesheet;
        }

        public static ThemeDocument Parse(Stream stream)
        {
            CT_OfficeStyleSheet obj = (CT_OfficeStyleSheet)serializer.Deserialize(stream);

            return new ThemeDocument(obj) ;
        }
        public CT_OfficeStyleSheet GetTheme()
        {
            return stylesheet;
        }

        public void Save(Stream stream)
        {
            serializer.Serialize(stream, stylesheet);
        }
    }
}
