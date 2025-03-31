﻿using NPOI.Util;
using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text; 
using Cysharp.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXml4Net.Util
{
    public static class XmlHelper
    {
        public static T ReadEnum<T>(XmlAttribute attr)
        {
            if (attr != null)
                return (T)Enum.Parse(typeof(T), attr.Value);
            return default(T);
        }
        public static string GetEnumValue(Enum e)
        {
            // Get the Type of the enum
            Type t = e.GetType();

            // Get the FieldInfo for the member field with the enums name
            FieldInfo info = t.GetField(e.ToString("G"));

            // Check to see if the XmlEnumAttribute is defined on this field
            if (!info.IsDefined(typeof(XmlEnumAttribute), false))
            {
                // If no XmlEnumAttribute then return the string version of the enum.
                return e.ToString("G");
            }

            // Get the XmlEnumAttribute
            object[] o = info.GetCustomAttributes(typeof(XmlEnumAttribute), false);
            XmlEnumAttribute att = (XmlEnumAttribute)o[0];
            return att.Name;
        }

        public static string GetXmlAttrNameFromEnumValue<T>(T pEnumVal)
        {
            // http://stackoverflow.com/q/3047125/194717
            Type type = pEnumVal.GetType();
            FieldInfo info = type.GetField(Enum.GetName(typeof(T), pEnumVal));
            XmlEnumAttribute att = (XmlEnumAttribute)info.GetCustomAttributes(typeof(XmlEnumAttribute), false)[0];
            //If there is an xmlattribute defined, return the name

            return att.Name;
        }
        public static T GetEnumValueFromString<T>(string value)
        {
            // http://stackoverflow.com/a/3073272/194717
            foreach (object o in System.Enum.GetValues(typeof(T)))
            {
                T enumValue = (T)o;
                if (GetXmlAttrNameFromEnumValue<T>(enumValue).Equals(value, StringComparison.OrdinalIgnoreCase))
                {
                    return (T)o;
                }
            }

            throw new ArgumentException("No XmlEnumAttribute code exists for type " + typeof(T).ToString() + " corresponding to value of " + value);
        }
        public static int ReadInt(XmlAttribute attr, int? defaultValue = null)
        {
            if(attr == null)
            {
                if(defaultValue != null)
                {
                    return (int)defaultValue;
                }
                return 0;
            }
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
        public static uint ReadUInt(XmlAttribute attr, uint defaultValue)
        {
            if (attr == null)
                return defaultValue;

            uint i;
            if (uint.TryParse(attr.Value, out i))
            {
                return i;
            }
            else
            {
                return defaultValue;
            }
        }
        public static uint ReadUInt(XmlAttribute attr)
        {
            return ReadUInt(attr, 0);
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
                if (double.TryParse(s, NumberStyles.Number|NumberStyles.Float| NumberStyles.AllowExponent, CultureInfo.InvariantCulture, out v))
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
            if (value == "1" || value == "-1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) || string.Equals(value, "on", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            else if (string.IsNullOrEmpty(value))
            {
                return blankValue;
            }
            else
            {
                return false;
            }
        }
        public static DateTime? ReadDateTime(XmlAttribute attr)
        {
            if (attr == null)
                return null;
            //TODO make this stable.
            return DateTime.Parse(attr.Value);
        }

        public static string ExcelEncodeString(string t)
        {
            StringWriter sw = new StringWriter();
            //poi dose not add prefix _x005f before _x????_ char.
            //if (Regex.IsMatch(t, "(_x[0-9A-F]{4,4}_)"))
            //{
            //    Match match = Regex.Match(t, "(_x[0-9A-F]{4,4}_)");
            //    int indexAdd = 0;
            //    while (match.Success)
            //    {
            //        t = t.Insert(match.Index + indexAdd, "_x005F");
            //        indexAdd += 6;
            //        match = match.NextMatch();
            //    }
            //}
            for (int i = 0; i < t.Length; i++)
            {
                if (t[i] <= 0x1f && t[i] != '\t' && t[i] != '\n' && t[i] != '\r') //Not Tab, CR or LF
                {
                    //[0x00-0x0a]-[\r\n\t]
                    //poi replace those chars with ?
                    sw.Write('?');
                    //sw.Write("_x00{0}_", (t[i] < 0xa ? "0" : "") + ((int)t[i]).ToString("X"));
                }
                else if (t[i] == '\uFFFE')
                {
                    sw.Write('?');
                }
                else
                {
                    sw.Write(t[i]);
                }
            }
            return sw.ToString();
        }
        public static string ExcelDecodeString(string t)
        {
            Match match = Regex.Match(t, "(_x005F|_x[0-9A-F]{4,4}_)");
            if (!match.Success) return t;

            bool useNextValue = false;
            using var ret = ZString.CreateStringBuilder();
            int prevIndex = 0;
            while (match.Success)
            {
                if (prevIndex < match.Index) ret.Append(t.Substring(prevIndex, match.Index - prevIndex));
                if (!useNextValue && match.Value == "_x005F")
                {
                    useNextValue = true;
                }
                else
                {
                    if (useNextValue)
                    {
                        ret.Append(match.Value);
                        useNextValue = false;
                    }
                    else
                    {
                        ret.Append((char)int.Parse(match.Value.Substring(2, 4)));
                    }
                }
                prevIndex = match.Index + match.Length;
                match = match.NextMatch();
            }
            ret.Append(t.Substring(prevIndex, t.Length - prevIndex));
            return ret.ToString();
        }

        private static readonly char[] xmlEncodeCharsToReplace = { '&', '<', '>', '"' };

        public static string EncodeXml(string xml)
        {
            // quick check whether needed
            if (xml.IndexOfAny(xmlEncodeCharsToReplace) == -1)
            {
                return xml;
            }

            return xml
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;");//.Replace("'", "&apos;");
        }

        public static void WriteAttribute(StreamWriter sw, string attributeName, bool value)
        {
            WriteAttribute(sw, attributeName, value, true);
        }
        public static void WriteAttribute(StreamWriter sw, string attributeName, bool value, bool writeIfBlank, bool defaultValue = false)
        {
            if (value == defaultValue && !writeIfBlank)
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
            WriteAttribute(sw, attributeName, value == 0.0 ? "0" : value.ToString(CultureInfo.InvariantCulture));
        }
        public static void WriteAttribute(StreamWriter sw, string attributeName, int value, bool writeIfBlank)
        {
            if (value == 0 && !writeIfBlank)
                return;

            sw.Write(" ");
            sw.Write(attributeName);
            sw.Write("=\"");
            sw.Write(value);
            sw.Write("\"");
        }
        public static void WriteAttribute(StreamWriter sw, string attributeName, int value, int defaultValue)
        {
            if(value == defaultValue)
                return;

            WriteAttribute(sw, attributeName, value.ToString(CultureInfo.InvariantCulture));
        }
        public static void WriteAttribute(StreamWriter sw, string attributeName, int value)
        {
            WriteAttribute(sw, attributeName, value, false);
        }
        public static void WriteAttribute(StreamWriter sw, string attributeName, uint value, bool writeIfBlank)
        {
            if (value == 0 && !writeIfBlank)
                return;

            sw.Write(" ");
            sw.Write(attributeName);
            sw.Write("=\"");
            sw.Write(value);
            sw.Write("\"");
        }
        public static void WriteAttribute(StreamWriter sw, string attributeName, uint value)
        {
            WriteAttribute(sw, attributeName, value, false);
        }
        public static void WriteAttribute(StreamWriter sw, string attributeName, string value)
        {
            WriteAttribute(sw, attributeName, value, false);
        }
        public static void WriteAttribute(StreamWriter sw, string attributeName, string value, bool writeIfBlank)
        {
            WriteAttribute(sw, attributeName, value, writeIfBlank, string.Empty);
        }
        public static void WriteAttribute(StreamWriter sw, string attributeName, string value, bool writeIfBlank, string defaultValue)
        {
            if ((string.IsNullOrEmpty(value) || defaultValue.Equals(value)) && !writeIfBlank)
                return;

            sw.Write(" ");
            sw.Write(attributeName);
            sw.Write("=\"");
            sw.Write(value == null ? string.Empty : EncodeXml(value));
            sw.Write("\"");
        }
        public static void WriteAttribute(StreamWriter sw, string attributeName, byte[] value)
        {
            if (value == null)
                return;

            WriteAttribute(sw, attributeName, BitConverter.ToString(value).Replace("-", ""), false);
        }

        public static void WriteAttribute(StreamWriter sw, string attributeName, uint value, uint defaultValue, bool writeIfBlank = false)
        {
            if(value != defaultValue)
                WriteAttribute(sw, attributeName, value, true);
            else if(writeIfBlank)
                WriteAttribute(sw, attributeName, value, writeIfBlank);
        }

        public static void WriteAttribute(StreamWriter sw, string attributeName, DateTime? value)
        {
            if (value == null)
                return;
            WriteAttribute(sw, attributeName, value.Value.ToString("yyyy-MM-ddTHH:mm:ss"), false);
        }

        public static void LoadXmlSafe(XmlDocument xmlDoc, Stream stream)
        {
            XmlReaderSettings settings = new XmlReaderSettings();
            //Disable entity parsing (to aviod xmlbombs, External Entity Attacks etc).
            settings.XmlResolver = null;
            settings.ProhibitDtd = true;
            //settings.MaxCharactersFromEntities = 4096;
            //settings.ValidationType = ValidationType.DTD;
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
