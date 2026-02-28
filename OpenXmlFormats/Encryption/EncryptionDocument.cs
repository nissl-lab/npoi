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
            if(ctEncryption == null)
                throw new InvalidOperationException("Encryption data not initialized.");

            var keyEncryptors = ctEncryption.keyEncryptors?.keyEncryptor ?? new List<CT_KeyEncryptor>();
            bool hasPass = false, hasCert = false;
            foreach(var ke in keyEncryptors)
            {
                hasPass |= HasPasswordKey(ke);
                hasCert |= HasCertificateKey(ke);
                if(hasPass && hasCert)
                    break;
            }

            var settings = new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false, OmitXmlDeclaration = false };
            using(var xw = XmlWriter.Create(stream, settings))
            {
                xw.WriteStartDocument(standalone: false); // match sample (standalone="no")
                xw.WriteStartElement("encryption", ENCRYPTION_DEFAULT);
                if(hasPass)
                    xw.WriteAttributeString("xmlns", "p", null, ENCRYPTION_PASSWORD);
                if(hasCert)
                    xw.WriteAttributeString("xmlns", "c", null, ENCRYPTION_CERTIFICATE);

                if(ctEncryption.keyData != null)
                    WriteKeyData(xw, ctEncryption.keyData);
                if(ctEncryption.dataIntegrity != null)
                    WriteDataIntegrity(xw, ctEncryption.dataIntegrity);

                if(keyEncryptors.Count > 0)
                {
                    xw.WriteStartElement("keyEncryptors", ENCRYPTION_DEFAULT);
                    foreach(var ke in keyEncryptors)
                    {
                        xw.WriteStartElement("keyEncryptor", ENCRYPTION_DEFAULT);
                        if(ke.uriSpecified)
                        {
                            var uriStr = EnumsNET.Enums.AsString(ke.uri, EnumsNET.EnumFormat.Description) ?? ke.uri.ToString();
                            xw.WriteAttributeString("uri", uriStr);
                        }
                        WriteKeyEncryptorChild(xw, ke); // always emit child element
                        xw.WriteEndElement(); // keyEncryptor
                    }
                    xw.WriteEndElement(); // keyEncryptors
                }

                xw.WriteEndElement(); // encryption
                xw.WriteEndDocument();
            }
        }

        private static void WriteKeyData(XmlWriter xw, CT_KeyData kd)
        {
            xw.WriteStartElement("keyData", ENCRYPTION_DEFAULT);
            xw.WriteAttributeString("blockSize", kd.blockSize.ToString());
            xw.WriteAttributeString("cipherAlgorithm", kd.cipherAlgorithm.ToString());
            xw.WriteAttributeString("cipherChaining", kd.cipherChaining.ToString());
            xw.WriteAttributeString("hashAlgorithm", kd.hashAlgorithm.ToString());
            xw.WriteAttributeString("hashSize", kd.hashSize.ToString());
            xw.WriteAttributeString("keyBits", kd.keyBits.ToString());
            xw.WriteAttributeString("saltSize", kd.saltSize.ToString());
            if(kd.saltValue != null)
                xw.WriteAttributeString("saltValue", Convert.ToBase64String(kd.saltValue));
            xw.WriteEndElement();
        }

        private static void WriteDataIntegrity(XmlWriter xw, CT_DataIntegrity di)
        {
            xw.WriteStartElement("dataIntegrity", ENCRYPTION_DEFAULT);
            if(di.encryptedHmacKey != null)
                xw.WriteAttributeString("encryptedHmacKey", Convert.ToBase64String(di.encryptedHmacKey));
            if(di.encryptedHmacValue != null)
                xw.WriteAttributeString("encryptedHmacValue", Convert.ToBase64String(di.encryptedHmacValue));
            xw.WriteEndElement();
        }

        private static void WriteKeyEncryptorChild(XmlWriter xw, CT_KeyEncryptor ke)
        {
            var child = ke.Item;
            string uri = null;
            if(ke.uriSpecified)
            {
                uri = EnumsNET.Enums.AsString(ke.uri, EnumsNET.EnumFormat.Description);
            }
            bool password = string.Equals(uri, ENCRYPTION_PASSWORD, StringComparison.OrdinalIgnoreCase);
            bool certificate = string.Equals(uri, ENCRYPTION_CERTIFICATE, StringComparison.OrdinalIgnoreCase);
            string prefix = password ? "p" : (certificate ? "c" : null);
            string localName = "encryptedKey"; // both use same local name with different namespace prefix

            if(prefix != null)
            {
                xw.WriteStartElement(prefix, localName, password ? ENCRYPTION_PASSWORD : ENCRYPTION_CERTIFICATE);
            }
            else
            {
                // Fallback to default namespace if uri not specified
                xw.WriteStartElement(localName, ENCRYPTION_DEFAULT);
            }

            if(child != null)
            {
                // Known agile password key encryptor attribute names (order matching sample)
                var attrOrder = new[]
                {
                    "blockSize","cipherAlgorithm","cipherChaining","encryptedKeyValue","encryptedVerifierHashInput","encryptedVerifierHashValue","hashAlgorithm","hashSize","keyBits","saltSize","saltValue","spinCount","encryptedHmacKey","encryptedHmacValue" // include potential future fields
                };

                var props = child.GetType().GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                foreach(var name in attrOrder)
                {
                    var p = props.FirstOrDefault(pr => pr.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                    if(p == null)
                        continue;
                    var val = p.GetValue(child, null);
                    if(val == null)
                        continue;
                    string strVal = null;
                    switch(val)
                    {
                        case byte[] bytes:
                            strVal = Convert.ToBase64String(bytes);
                            break;
                        default:
                            strVal = Convert.ToString(val, System.Globalization.CultureInfo.InvariantCulture);
                            break;
                    }
                    if(!string.IsNullOrEmpty(strVal))
                        xw.WriteAttributeString(p.Name, strVal);
                }
            }

            xw.WriteEndElement(); // encryptedKey
        }

        private static bool HasPasswordKey(CT_KeyEncryptor keyEncryptor)
        {
            if(!keyEncryptor.uriSpecified)
                return false;
            var uriStr = EnumsNET.Enums.AsString(keyEncryptor.uri, EnumsNET.EnumFormat.Description);
            return string.Equals(uriStr, ENCRYPTION_PASSWORD, StringComparison.OrdinalIgnoreCase);
        }
        private static bool HasCertificateKey(CT_KeyEncryptor keyEncryptor)
        {
            if(!keyEncryptor.uriSpecified)
                return false;
            var uriStr = EnumsNET.Enums.AsString(keyEncryptor.uri, EnumsNET.EnumFormat.Description);
            return string.Equals(uriStr, ENCRYPTION_CERTIFICATE, StringComparison.OrdinalIgnoreCase);
        }

        public static EncryptionDocument Parse(string descriptor)
        {
            if(string.IsNullOrEmpty(descriptor))
                throw new ArgumentNullException(nameof(descriptor));
            string trimmed = descriptor.TrimStart();
            var xmlDoc = new XmlDocument { XmlResolver = null };
            bool looksLikeXml = trimmed.StartsWith('<');
            if(!looksLikeXml && File.Exists(descriptor))
            {
                using(var fs = File.OpenRead(descriptor))
                    xmlDoc.Load(fs);
            }
            else
            {
#if NETSTANDARD2_1 || NET5_0_OR_GREATER
                if (descriptor.Contains("<!DOCTYPE", StringComparison.OrdinalIgnoreCase))
#else
                if(descriptor.IndexOf("<!DOCTYPE", StringComparison.OrdinalIgnoreCase) >= 0)
#endif
                    throw new XmlException("DOCTYPE is not allowed in encryption descriptor.");
                xmlDoc.LoadXml(descriptor);
            }
            return Parse(xmlDoc, EncryptionNamespaceManager);
        }

        public static EncryptionDocument Parse(XmlDocument xmlDoc)
        {
            return Parse(xmlDoc, EncryptionNamespaceManager);
        }
    }
}
