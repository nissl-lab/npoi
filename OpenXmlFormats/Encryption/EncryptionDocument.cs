using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace NPOI.OpenXmlFormats.Encryption
{
    public class EncryptionDocument
    {
        public static EncryptionDocument Parse(XmlDocument xmlDoc, XmlNamespaceManager NameSpaceManager)
        {
            throw new NotImplementedException();
        }
        public CT_Encryption GetEncryption()
        {
            throw new NotImplementedException();
        }
        public void SetEncryption(CT_Encryption encryption)
        {
            throw new NotImplementedException();
        }
        public CT_Encryption AddNewEncryption()
        {
            throw new NotImplementedException();
        }
        public void Save(Stream stream)
        {
            StreamWriter sw = new StreamWriter(stream);
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>");

            sw.Flush();
        }

        public static EncryptionDocument NewInstance()
        {
            throw new NotImplementedException();
        }

        public static EncryptionDocument Parse(string descriptor)
        {
            throw new NotImplementedException();
        }

        public static EncryptionDocument Parse(Stream descriptor)
        {
            throw new NotImplementedException();
        }
    }
}
