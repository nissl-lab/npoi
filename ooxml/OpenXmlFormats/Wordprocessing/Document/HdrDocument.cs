using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Wordprocessing
{
    public class HdrDocument
    {
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Hdr));

        CT_Hdr hdr = null;
        public HdrDocument()
        {
            hdr = new CT_Hdr();
        }
        public static HdrDocument Parse(Stream stream)
        {
            CT_Hdr obj = (CT_Hdr)serializer.Deserialize(stream);

            return new HdrDocument(obj);
        }
        public HdrDocument(CT_Hdr hdr)
        {
            this.hdr = hdr;
        }
        public CT_Hdr Hdr
        {
            get
            {
                return this.hdr;
            }
        }
        public void Save(Stream stream, XmlSerializerNamespaces namespaces)
        {
            serializer.Serialize(stream, hdr, namespaces);
        }

        public void SetHdr(CT_Hdr hdr)
        {
            this.hdr = hdr;
        }
    }
}
