using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Wordprocessing
{
    public class FtrDocument
    {
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Ftr));

        CT_Ftr ftr = null;
        public FtrDocument()
        {
            ftr = new CT_Ftr();
        }
        public static FtrDocument Parse(Stream stream)
        {
            CT_Ftr obj = (CT_Ftr)serializer.Deserialize(stream);

            return new FtrDocument(obj);
        }
        public FtrDocument(CT_Ftr ftr)
        {
            this.ftr = ftr;
        }
        public CT_Ftr Ftr
        {
            get
            {
                return this.ftr;
            }
        }
        public void Save(Stream stream, XmlSerializerNamespaces namespaces)
        {
            serializer.Serialize(stream, ftr, namespaces);
        }

        public void SetFtr(CT_Ftr ftr)
        {
            this.ftr = ftr;
        }
    }
}
