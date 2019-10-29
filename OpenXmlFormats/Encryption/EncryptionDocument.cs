using NPOI.OpenXml4Net.OPC;
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
        public static string ENCRYPTION_DEFAULT = "http://schemas.microsoft.com/office/2006/encryption";
        public static string ENCRYPTION_PASSWORD = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";
        public static string ENCRYPTION_CERTIFICATE = "http://schemas.microsoft.com/office/2006/keyEncryptor/certificate";
        static XmlNamespaceManager nsm = null;
        public static XmlNamespaceManager EncryptionNamespaceManager
        {
            get
            {
                if (nsm == null)
                    nsm = CreateEncryptionNSM();
                return nsm;
            }
        }
        internal static XmlNamespaceManager CreateEncryptionNSM()
        {
            //  Create a NamespaceManager to handle the default namespace, 
            //  and create a prefix for the default namespace:
            NameTable nt = new NameTable();
            XmlNamespaceManager ns = new XmlNamespaceManager(nt);
            ns.AddNamespace(string.Empty, EncryptionDocument.ENCRYPTION_DEFAULT);
            ns.AddNamespace("p", EncryptionDocument.ENCRYPTION_PASSWORD);
            ns.AddNamespace("c", EncryptionDocument.ENCRYPTION_CERTIFICATE);

            ns.AddNamespace("xsd", "http://www.w3.org/2001/XMLSchema");
            return ns;
        }
        public EncryptionDocument()
        {
        }
        public EncryptionDocument(CT_Encryption encryption)
        {
            this.ctEncryption = encryption;
        }
        private CT_Encryption ctEncryption;

        public static EncryptionDocument Parse(XmlDocument xmlDoc, XmlNamespaceManager NameSpaceManager)
        {
            CT_Encryption obj = CT_Encryption.Parse(xmlDoc.DocumentElement, NameSpaceManager);
            return new EncryptionDocument(obj);
        }
        public CT_Encryption GetEncryption()
        {
            return this.ctEncryption;
        }
        public void SetEncryption(CT_Encryption encryption)
        {
            this.ctEncryption = encryption;
        }
        public CT_Encryption AddNewEncryption()
        {
            this.ctEncryption = new CT_Encryption();
            return this.ctEncryption;
        }
        public void Save(Stream stream)
        {
            StreamWriter sw = new StreamWriter(stream);
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>");

            sw.Flush();
            throw new NotImplementedException();
        }

        public static EncryptionDocument NewInstance()
        {
            throw new NotImplementedException();
        }

        public static EncryptionDocument Parse(string descriptor)
        {
            throw new NotImplementedException();
        }

        public static EncryptionDocument Parse(XmlDocument xmlDoc)
        {
            return Parse(xmlDoc, EncryptionNamespaceManager);
        }
    }
}
