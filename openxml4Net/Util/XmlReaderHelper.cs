using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXml4Net.Util
{
    public static class XmlReaderHelper
    {
        public static int ReadInt(string attr)
        {
            if (attr == null)
                return 0;
            int i;
            if (int.TryParse(attr, out i))
            {
                return i;
            }
            else
            {
                return 0;
            }
        }
        public static long ReadLong(string attr)
        {
            if (attr == null)
                return 0;
            long i;
            if (long.TryParse(attr, out i))
            {
                return i;
            }
            else
            {
                return 0;
            }
        }
        public static int? ReadIntNull(string attr)
        {
            if (attr == null)
                return null;
            int i;
            string s = attr;
            if (s != "" && int.TryParse(s, out i))
            {
                return i;
            }
            else
            {
                return null;
            }
        }
        public static string ReadString(string attr)
        {
            return attr;
        }
        public static decimal ReadDecimal(string attr)
        {
            if (attr == null)
                return 0;
            decimal d;
            if (decimal.TryParse(attr, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
            {
                return d;
            }
            else
            {
                return 0;
            }
        }
        public static uint ReadUInt(string attr, uint defaultValue)
        {
            uint i;
            if (uint.TryParse(attr, out i))
            {
                return i;
            }
            else
            {
                return defaultValue;
            }
        }
        public static uint ReadUInt(string attr)
        {
            return ReadUInt(attr, 0);
        }
        public static ulong ReadULong(string attr)
        {
            if (attr == null)
                return 0;

            ulong i;
            if (ulong.TryParse(attr, out i))
            {
                return i;
            }
            else
            {
                return 0;
            }
        }
        public static bool ReadBool(string attr)
        {
            return ReadBool(attr, false);
        }
        public static double ReadDouble(string attr)
        {
            if (attr == null)
                return 0.0;
            string s = attr;
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
        public static double? ReadDoubleNull(string attr)
        {
            if (attr == null)
                return null;
            string s = attr;
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

        public static bool ReadBool(string attr, bool blankValue)
        {
            if (attr == null)
                return blankValue;

            string value = attr;
            if (value == "1" || value == "-1" || value.ToLower() == "true" || value.ToLower() == "on")
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
        public static DateTime? ReadDateTime(string attr)
        {
            if (attr == null)
                return null;
            //TODO make this stable.
            return DateTime.Parse(attr);
        }

        public static string EncodeXml(string xml)
        {
            return xml.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");//.Replace("'", "&apos;");
        }

        public static byte[] ReadBytes(string attr)
        {
            if (attr == null || string.IsNullOrEmpty(attr))
                return null;

            int NumberChars = attr.Length;
            byte[] bytes = new byte[NumberChars / 2];
            for (int i = 0; i < NumberChars; i += 2)
                bytes[i / 2] = Convert.ToByte(attr.Substring(i, 2), 16);
            return bytes;
        }

        public static sbyte ReadSByte(string attr)
        {
            if (attr == null)
                return 0;

            sbyte i;
            if (sbyte.TryParse(attr, out i))
            {
                return i;
            }
            else
            {
                return 0;
            }
        }

        public static ushort ReadUShort(string attr)
        {
            if (attr == null)
                return 0;

            ushort i;
            if (ushort.TryParse(attr, out i))
            {
                return i;
            }
            else
            {
                return 0;
            }
        }

        public static byte ReadByte(string attr)
        {
            if (attr == null)
                return 0;

            byte i;
            if (byte.TryParse(attr, out i))
            {
                return i;
            }
            else
            {
                return 0;
            }
        }

        /// <summary>
        /// 删除SheetData节点下第二个及其之后所有Row节点
        /// </summary>
        /// <param name="sourceStream"></param>
        /// <param name="dataRowCount"></param>
        /// <returns></returns>
        public static Stream RemoveSheetData(Stream sourceStream, out int dataRowCount)
        {
            XmlTextReader reader = new XmlTextReader(sourceStream);
            reader.DtdProcessing = DtdProcessing.Ignore;
            dataRowCount = 0;

            int depth = 1;
            string removeElementName = "sheetData";
            string rowElementName = "row";

            MemoryStream ms = new MemoryStream();
            var writer = new XmlTextWriter(ms, reader.Encoding);
            try
            {
                bool isInSheetDataNode = false;
                while (reader.Read())
                {
                    //go in sheetData node
                    if (!isInSheetDataNode && reader.NodeType == XmlNodeType.Element && reader.Depth == depth && reader.Name == removeElementName)
                    {
                        isInSheetDataNode = true;
                    }

                    //go in sheetData node
                    if (isInSheetDataNode && reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth && reader.Name == removeElementName)
                    {
                        isInSheetDataNode = false;
                    }

                    //count row if is in the sheetdata node 
                    if (isInSheetDataNode && reader.NodeType == XmlNodeType.Element && reader.Name == rowElementName)
                    {
                        dataRowCount++;
                    }

                    if (isInSheetDataNode && dataRowCount > 1)
                    {
                        continue;
                    }

                    switch (reader.NodeType)
                    {
                        case XmlNodeType.XmlDeclaration:
                            writer.WriteStartDocument(true);
                            break;
                        case XmlNodeType.Whitespace:
                            writer.WriteWhitespace(reader.Value);
                            break;
                        case XmlNodeType.Element:
                            if (reader.Depth == 0)
                            {
                                WriteRootElement(reader, writer);
                            }
                            else
                            {
                                WriteElement(reader, writer);
                            }
                            break;
                        case XmlNodeType.Text:
                            writer.WriteValue(reader.Value);
                            break;
                        case XmlNodeType.EndElement:
                            writer.WriteEndElement();
                            break;
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                if (writer != null && writer.WriteState != WriteState.Closed)
                {
                    writer.Flush();
                    //writer.Close();
                }

                if (reader != null && reader.ReadState != ReadState.Closed)
                    reader.Close();
            }

            ms.Position = 0;
            return ms;
        }

        private static void WriteElement(XmlTextReader reader, XmlTextWriter writer)
        {
            writer.WriteStartElement(reader.Name);
            var isEmptyElement = reader.IsEmptyElement;
            if (reader.MoveToFirstAttribute())
            {
                do
                {
                    writer.WriteStartAttribute(reader.Name);
                    writer.WriteString(reader.Value);
                    writer.WriteEndAttribute();
                }
                while (reader.MoveToNextAttribute());
            }
            if (isEmptyElement)
            {
                writer.WriteEndElement();
            }
        }

        private static void WriteRootElement(XmlTextReader reader, XmlTextWriter writer)
        {
            string rootName = reader.Name;
            string xmlns = string.Empty;
            var rootAttributes = new List<string[]>();
            if (reader.MoveToFirstAttribute())
            {
                do
                {
                    if (reader.Name == "xmlns")
                    {
                        xmlns = reader.Value;
                    }
                    else
                    {
                        rootAttributes.Add(new string[] { reader.Name, reader.Value });
                    }
                }
                while (reader.MoveToNextAttribute());
            }

            writer.WriteStartElement(rootName, xmlns);
            foreach (var attr in rootAttributes)
            {
                writer.WriteStartAttribute(attr[0]);
                writer.WriteString(attr[1]);
                writer.WriteEndAttribute();
            }
        }

        private static void WriteNodeToXmlFile(XmlTextReader reader, string nodeName, int depth, string filePath)
        {
            var fileName = string.Format(filePath + "_{0}.xml", nodeName);
            if (File.Exists(fileName))
            {
                try
                {
                    File.Delete(fileName);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            using (var fs = new FileStream(fileName, FileMode.OpenOrCreate))
            {
                try
                {
                    StreamWriter sw = new StreamWriter(fs);
                    sw.Write(reader.ReadOuterXml());
                    sw.Close();
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }
    }
}
