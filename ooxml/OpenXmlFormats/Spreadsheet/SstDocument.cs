using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class SstDocument
    {
        static XmlSerializer serializer = new XmlSerializer(typeof(CT_Sst));

        public SstDocument()
        { 
        
        }
        CT_Sst sst = null;
        public SstDocument(CT_Sst sst)
        {
            this.sst = sst;
        }


        public void AddNewSst()
        {
            this.sst = new CT_Sst();
        }

        public CT_Sst GetSst()
        {
            return this.sst;
        }
        public static SstDocument Parse(Stream in1)
        {
            CT_Sst sst= (CT_Sst)serializer.Deserialize(in1);
            return new SstDocument(sst);
        }
        public void Save(Stream stream)
        {
            serializer.Serialize(stream, sst);
        }

    }
}
