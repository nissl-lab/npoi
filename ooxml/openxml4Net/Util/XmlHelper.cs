using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;

namespace NPOI.OpenXml4Net.Util
{
    public static class XmlHelper
    {
        public static int ReadInt(XmlAttribute attr)
        {
            if (attr == null)
                return 0;
            int i;
            if (int.TryParse(attr.Value, out i))
            {
                return i;
            }
            else
            {
                return 0;
            }
        }
        public static long ReadLong(XmlAttribute attr)
        {
            if (attr == null)
                return 0;
            long i;
            if (long.TryParse(attr.Value, out i))
            {
                return i;
            }
            else
            {
                return 0;
            }
        }
        public static int? ReadIntNull(XmlAttribute attr)
        {
            if (attr == null)
                return null;
            int i;
            string s = attr.Value;
            if (s != "" && int.TryParse(s, out i))
            {
                return i;
            }
            else
            {
                return null;
            }
        }
        public static string ReadString(XmlAttribute attr)
        {
            if (attr == null)
                return null;
            return attr.Value;
        }
        public static decimal ReadDecimal(XmlAttribute attr)
        {
            if (attr == null)
                return 0;
            decimal d;
            if (decimal.TryParse(attr.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
            {
                return d;
            }
            else
            {
                return 0;
            }
        }
        public static uint ReadUInt(XmlAttribute attr)
        {
            if (attr == null)
                return 0;

            uint i;
            if (uint.TryParse(attr.Value, out i))
            {
                return i;
            }
            else
            {
                return 0;
            }
        }
        public static ulong ReadULong(XmlAttribute attr)
        {
            if (attr == null)
                return 0;

            ulong i;
            if (ulong.TryParse(attr.Value, out i))
            {
                return i;
            }
            else
            {
                return 0;
            }
        }
        public static bool ReadBool(XmlAttribute attr)
        {
            return ReadBool(attr, false);
        }
        public static double ReadDouble(XmlAttribute attr)
        {
            if (attr == null)
                return 0.0;
            string s = attr.Value;
            if (s == "")
            {
                return 0.0;
            }
            else
            {
                double v;
                if (double.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out v))
                {
                    return v;
                }
                else
                {
                    return 0.0;
                }
            }
        }
        public static double? ReadDoubleNull(XmlAttribute attr)
        {
            if (attr == null)
                return null;
            string s = attr.Value;
            if (s == "")
            {
                return null;
            }
            else
            {
                double v;
                if (double.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out v))
                {
                    return v;
                }
                else
                {
                    return null;
                }
            }
        }

        public static bool ReadBool(XmlAttribute attr, bool blankValue)
        {
            if (attr == null)
                return blankValue;

            string value = attr.Value;
            if (value == "1" || value == "-1" || value.ToLower() == "true" || value.ToLower() == "on")
            {
                return true;
            }
            else if (value == "")
            {
                return blankValue;
            }
            else
            {
                return false;
            }
        }
        public static string EncodeXml(string xml)
        {
            return xml.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;");
        }
        public static void WriteAttribute(StreamWriter sw, string attributeName, bool value)
        {
            WriteAttribute(sw, attributeName, value, true);
        }
        public static void WriteAttribute(StreamWriter sw, string attributeName, bool value, bool writeIfBlank)
        {
            if (value == false && !writeIfBlank)
                return;
            WriteAttribute(sw, attributeName, value ? "1" : "0");
        }
        public static void WriteAttribute(StreamWriter sw, string attributeName, double value)
        {
            WriteAttribute(sw, attributeName, value, false);
        }
        public static void WriteAttribute(StreamWriter sw, string attributeName, double value, bool writeIfBlank)
        {
            if (value == 0.0 && !writeIfBlank)
                return;
            WriteAttribute(sw, attributeName, value == 0.0 ? "0" : value.ToString());
        }
        public static void WriteAttribute(StreamWriter sw, string attributeName, int value, bool writeIfBlank)
        {
            if (value == 0 && !writeIfBlank)
                return;

            WriteAttribute(sw, attributeName, value.ToString());
        }
        public static void WriteAttribute(StreamWriter sw, string attributeName, int value)
        {
            WriteAttribute(sw, attributeName, value, false);
        }
        public static void WriteAttribute(StreamWriter sw, string attributeName, string value)
        {
            WriteAttribute(sw, attributeName, value, false);
        }
        public static void WriteAttribute(StreamWriter sw, string attributeName, string value, bool writeIfBlank)
        {
            if (string.IsNullOrEmpty(value) && !writeIfBlank)
                return;
            sw.Write(string.Format(" {0}=\"{1}\"", attributeName, value == null ? string.Empty : EncodeXml(value)));
        }
        public static void WriteAttribute(StreamWriter sw, string attributeName, byte[] value)
        {
            if (value == null)
                return;

            WriteAttribute(sw, attributeName, BitConverter.ToString(value).Replace("-", ""), false);
        }
        public static void WriteAttribute(StreamWriter sw, string attributeName, uint value)
        {
            WriteAttribute(sw, attributeName, (int)value, false);
        }
        public static void LoadXmlSafe(XmlDocument xmlDoc, Stream stream)
        {
            XmlReaderSettings settings = new XmlReaderSettings();
            //Disable entity parsing (to aviod xmlbombs, External Entity Attacks etc).
            settings.ProhibitDtd = true;

            XmlReader reader = XmlReader.Create(stream, settings);
            xmlDoc.Load(reader);
        }
        public static void LoadXmlSafe(XmlDocument xmlDoc, string xml, Encoding encoding)
        {
            MemoryStream stream = new MemoryStream(encoding.GetBytes(xml));
            LoadXmlSafe(xmlDoc, stream);
        }

        public static byte[] ReadBytes(XmlAttribute attr)
        {
            if (attr == null || string.IsNullOrEmpty(attr.Value))
                return null;

            int NumberChars = attr.Value.Length;
            byte[] bytes = new byte[NumberChars / 2];
            for (int i = 0; i < NumberChars; i += 2)
                bytes[i / 2] = Convert.ToByte(attr.Value.Substring(i, 2), 16);
            return bytes;
        }

        public static sbyte ReadSByte(XmlAttribute attr)
        {
            if (attr == null)
                return 0;

            sbyte i;
            if (sbyte.TryParse(attr.Value, out i))
            {
                return i;
            }
            else
            {
                return 0;
            }
        }

        public static ushort ReadUShort(XmlAttribute attr)
        {
            if (attr == null)
                return 0;

            ushort i;
            if (ushort.TryParse(attr.Value, out i))
            {
                return i;
            }
            else
            {
                return 0;
            }
        }

        public static byte ReadByte(XmlAttribute attr)
        {
            if (attr == null)
                return 0;

            byte i;
            if (byte.TryParse(attr.Value, out i))
            {
                return i;
            }
            else
            {
                return 0;
            }
        }
    }
}
