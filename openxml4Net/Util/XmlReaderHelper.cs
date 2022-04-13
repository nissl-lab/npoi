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
        /// xml拆分
        /// </summary>
        /// <param name="path">大文件路径</param>
        /// <param name="nodeCount">小文件中节点数</param>
        public static void SplitXml(string path, int nodeCount)
        {
            XmlTextReader reader = new XmlTextReader(path);
            reader.DtdProcessing = DtdProcessing.Ignore;
            XmlWriter writer = null;
            string rootName = string.Empty;
            string filePath = path.Substring(0, path.LastIndexOf("."));
            try
            {
                List<string[]> rootAttributes = new List<string[]>();
                int count = 0;
                while (reader.Read())
                {
                    switch (reader.NodeType)
                    {
                        case XmlNodeType.Whitespace:
                            if (writer != null && writer.WriteState != WriteState.Closed)
                            {
                                writer.WriteWhitespace(reader.Value);
                            }

                            break;
                        case XmlNodeType.Element:
                            if (reader.Depth == 0) rootName = reader.Name;

                            if (reader.Name == rootName) // root
                            {
                                // read root Attributes
                                if (reader.HasAttributes)
                                {
                                    rootAttributes = new List<string[]>();
                                    for (int i = 0; i < reader.AttributeCount; i++)
                                    {
                                        reader.MoveToAttribute(i);
                                        rootAttributes.Add(new string[] { reader.Name, reader.Value });
                                    }
                                    reader.MoveToElement();
                                }
                            }
                            else
                            {
                                if (reader.Depth == 1 && count % nodeCount == 0)
                                {
                                    writer = XmlWriter.Create(string.Format(filePath + ".part{0}.xml", count / nodeCount + 1));

                                    writer.WriteStartDocument(); // <?xml version="1.0" encoding="utf-8"?>
                                    writer.WriteWhitespace(Environment.NewLine);

                                    // write root Start Element
                                    writer.WriteStartElement(rootName);
                                    // write root Attributes
                                    foreach (var attribute in rootAttributes)
                                    {
                                        writer.WriteStartAttribute(attribute[0]);
                                        writer.WriteString(attribute[1]);
                                        writer.WriteEndAttribute();
                                    }
                                    writer.WriteWhitespace(Environment.NewLine);
                                }

                                if (reader.IsEmptyElement) // empty element, <{0} />
                                {
                                    writer.WriteRaw(string.Format("<{0} />", reader.Name));
                                }
                                else
                                {
                                    // writer Start Element
                                    writer.WriteStartElement(reader.Name);
                                    // writer Element Attributes
                                    if (reader.HasAttributes)
                                    {
                                        for (int i = 0; i < reader.AttributeCount; i++)
                                        {
                                            reader.MoveToAttribute(i);
                                            writer.WriteStartAttribute(reader.Name);
                                            writer.WriteString(reader.Value);
                                            writer.WriteEndAttribute();
                                        }
                                        reader.MoveToElement();
                                    }
                                }
                            }

                            break;
                        case XmlNodeType.Text:
                            writer.WriteValue(reader.Value);

                            break;
                        case XmlNodeType.EndElement:
                            if (reader.Depth == 1)
                            {
                                writer.WriteEndElement();
                                count++;

                                // write root end element
                                if (count > 0 && count % nodeCount == 0)
                                {
                                    writer.WriteWhitespace(Environment.NewLine);
                                    writer.WriteEndElement();
                                    writer.Close();
                                }
                            }
                            else
                            {
                                if (reader.Name != rootName)
                                    writer.WriteEndElement();
                            }

                            // write root end element
                            if (reader.Depth == 0 && writer.WriteState != WriteState.Closed)
                            {
                                writer.WriteWhitespace(Environment.NewLine);
                                writer.WriteEndElement();
                                writer.Close();
                            }

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
                    writer.Close();

                if (reader != null && reader.ReadState != ReadState.Closed)
                    reader.Close();
            }
        }

        public static long[] CountPosition(string sourcePath, int depth, string elementName)
        {
            long[] rst = new long[2];
            long num = 0;
            var sourceStream = File.Open(sourcePath, FileMode.Open);
            XmlTextReader reader = new XmlTextReader(sourceStream);
            bool isInRemoveElement = false;
            try
            {
                while (reader.Read())
                {
                    if (isInRemoveElement && reader.NodeType != XmlNodeType.EndElement && reader.Depth <= depth)
                    {
                        isInRemoveElement = false;
                        rst[1] = num;
                        break;
                    }
                    if (isInRemoveElement)
                    {
                        continue;
                    }

                    num += reader.Value.Length;

                    switch (reader.NodeType)
                    {
                        case XmlNodeType.Element:
                            if (reader.Depth == depth && reader.Name == elementName)
                            {
                                isInRemoveElement = true;
                                rst[0] = num;
                            }
                            break;
                        default:
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
                if (reader != null && reader.ReadState != ReadState.Closed)
                    reader.Close();
            }
            return rst;
        }

        /// <summary>
        /// Remove element and save to a new file
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <param name="newPath"></param>
        /// <param name="depth"></param>
        /// <param name="elementName"></param>
        public static void RemoveXmlElement(string sourcePath, string newPath, int depth, string elementName)
        {
            var sourceStream = File.Open(sourcePath, FileMode.Open);

            //var virtualStream = new VirtualStream(sourceStream, 8192, 67145142);
            var virtualStream = new VirtualStream(sourceStream, 1172, 33760);

            XmlTextReader reader = new XmlTextReader(virtualStream);

            reader.DtdProcessing = DtdProcessing.Ignore;
            XmlTextWriter writer = null;
            bool isInRemoveElement = false;
            try
            {
                while (reader.Read())
                {
                    if (isInRemoveElement && reader.NodeType != XmlNodeType.EndElement && reader.Depth <= depth)
                    {
                        isInRemoveElement = false;
                    }
                    if (isInRemoveElement)
                    {
                        continue;
                    }

                    switch (reader.NodeType)
                    {
                        case XmlNodeType.XmlDeclaration:
                            writer = new XmlTextWriter(File.Open(newPath, FileMode.OpenOrCreate), reader.Encoding);
                            writer.WriteStartDocument(true);

                            break;
                        case XmlNodeType.Whitespace:
                            if (writer != null && writer.WriteState != WriteState.Closed)
                            {
                                writer.WriteWhitespace(reader.Value);
                            }
                            break;
                        case XmlNodeType.Element:

                            if (reader.Depth == 0)
                            {
                                WriteRootElement(reader, writer);
                                break;
                            }

                            if (reader.Depth == depth && reader.Name == elementName)
                            {
                                isInRemoveElement = true;
                                break;
                            }

                            WriteElement(reader, writer);
                            break;
                        case XmlNodeType.Text:
                            writer.WriteValue(reader.Value);
                            break;
                        case XmlNodeType.EndElement:
                            writer.WriteEndElement();

                            if (reader.Depth == 0 && writer.WriteState != WriteState.Closed)
                            {
                                writer.Close();
                            }

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
                    writer.Close();

                if (reader != null && reader.ReadState != ReadState.Closed)
                    reader.Close();
            }
        }

        public static XmlTextReader GetTextReaderXmlElement(string sourcePath, int depth, string elementName)
        {
            var sourceStream = File.Open(sourcePath, FileMode.Open);
            var virtualStream = new VirtualStream(sourceStream, 1172, 33760);

            XmlTextReader reader = new XmlTextReader(virtualStream);
            try
            {
                while (reader.Read())
                {
                    switch (reader.NodeType)
                    {
                        case XmlNodeType.XmlDeclaration:

                            break;
                        case XmlNodeType.Whitespace:

                            break;
                        case XmlNodeType.Element:

                            break;
                        case XmlNodeType.Text:
                            break;
                        case XmlNodeType.EndElement:

                            break;
                        default:

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
                if (reader != null && reader.ReadState != ReadState.Closed)
                    reader.Close();
            }
            return reader;
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
