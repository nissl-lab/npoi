/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
namespace NPOI.XSSF.EventUserModel
{
    using NPOI.OpenXml4Net.OPC;
    using NPOI.XSSF.UserModel;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// <para>
    /// </para>
    /// <para>
    /// This is a lightweight way to process the Shared Strings
    /// table. Most of the text cells will reference something
    /// from in here.
    /// </para>
    /// <para>
    /// Note that each SI entry can have multiple T elements, if the
    /// string is made up of bits with different formatting.
    /// </para>
    /// <para>
    /// Example input:
    /// <code>
    /// &lt;?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
    /// &lt;sst xmlns="http://schemas.Openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
    ///   &lt;si>
    ///     &lt;r>
    ///       &lt;rPr>
    ///         &lt;b />
    ///         &lt;sz val="11" />
    ///         &lt;color theme="1" />
    ///         &lt;rFont val="Calibri" />
    ///         &lt;family val="2" />
    ///         &lt;scheme val="minor" />
    ///       &lt;/rPr>
    ///       &lt;t>This:&lt;/t>
    ///     &lt;/r>
    ///     &lt;r>
    ///       &lt;rPr>
    ///         &lt;sz val="11" />
    ///         &lt;color theme="1" />
    ///         &lt;rFont val="Calibri" />
    ///         &lt;family val="2" />
    ///         &lt;scheme val="minor" />
    ///       &lt;/rPr>
    ///       &lt;t xml:space="preserve">Causes Problems&lt;/t>
    ///     &lt;/r>
    ///   &lt;/si>
    ///   &lt;si>
    ///     &lt;t>This does not&lt;/t>
    ///   &lt;/si>
    /// &lt;/sst>
    /// </code>
    /// </para>
    /// </summary>
    public class ReadOnlySharedStringsTable
    {

        private bool includePhoneticRuns;
        /// <summary>
        /// An integer representing the total count of strings in the workbook. This count does not
        /// include any numbers, it counts only the total of text strings in the workbook.
        /// </summary>
        private int count;

        /// <summary>
        /// An integer representing the total count of unique strings in the Shared String Table.
        /// A string is unique even if it is a copy of another string, but has different formatting applied
        /// at the character level.
        /// </summary>
        private int uniqueCount;

        /// <summary>
        /// The shared strings table.
        /// </summary>
        private List<String> strings;

        /// <summary>
        /// Map of phonetic strings (if they exist) indexed
        /// with the integer matching the index in strings
        /// </summary>
        private Dictionary<int, String> phoneticStrings;

        /// <summary>
        /// Calls <see cref="ReadOnlySharedStringsTable(OPCPackage, bool)" /> with
        /// a value of <c>true</c> for including phonetic runs
        /// </summary>
        /// <param name="pkg">The <see cref="OPCPackage"/> to use as basis for the shared-strings table.</param>
        /// <exception cref="IOException"> If reading the data from the package fails.</exception>
        public ReadOnlySharedStringsTable(OPCPackage pkg)
            : this(pkg, true)
        {
        }

        /// <summary>
        /// </summary>
        /// <param name="pkg">The <see cref="OPCPackage"/> to use as basis for the shared-strings table.</param>
        /// <param name="includePhoneticRuns">whether or not to concatenate phoneticRuns onto the shared string</param>
        /// <exception cref="IOException">IOException If reading the data from the package fails.</exception>
        public ReadOnlySharedStringsTable(OPCPackage pkg, bool includePhoneticRuns)
        {
            this.includePhoneticRuns = includePhoneticRuns;
            List<PackagePart> parts =
                    pkg.GetPartsByContentType(XSSFRelation.SHARED_STRINGS.ContentType);

            // Some workbooks have no shared strings table.
            if (parts.Count > 0)
            {
                PackagePart sstPart = parts[0];
                ReadFrom(sstPart.GetInputStream());
            }
        }

        /// <summary>
        /// <para>
        /// Like POIXMLDocumentPart constructor
        /// </para>
        /// <para>
        /// Calls <see cref="ReadOnlySharedStringsTable(PackagePart, bool)" />, with a
        /// value of <c>true</c> to include phonetic runs.
        /// </para>
        /// </summary>
        /// @since POI 3.14-Beta1
        public ReadOnlySharedStringsTable(PackagePart part)
            : this(part, true)
        {
        }

        /// <summary>
        /// Like POIXMLDocumentPart constructor
        /// </summary>
        /// @since POI 3.14-Beta3
        public ReadOnlySharedStringsTable(PackagePart part, bool includePhoneticRuns)
        {
            this.includePhoneticRuns = includePhoneticRuns;
            ReadFrom(part.GetInputStream());
        }

        /// <summary>
        /// Read this shared strings table from an XML file.
        /// </summary>
        /// <param name="is1">The input stream containing the XML document.</param>
        /// <exception cref="IOException"> if an error occurs while reading.</exception>
        public void ReadFrom(Stream is1)
        {
            // test if the file is empty, otherwise parse it
            if(is1.Length > 0)
            {
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.DtdProcessing = DtdProcessing.Ignore;
                var reader = XmlReader.Create(is1, settings);
                while(reader.Read())
                {
                    if(reader.NodeType == XmlNodeType.Element)
                    {
                        //begin element
                        StartElement(reader);
                    }
                    else if(reader.NodeType == XmlNodeType.EndElement)
                    {
                        EndElement(reader);
                    }
                    else if(reader.NodeType == XmlNodeType.Text ||
                        reader.NodeType == XmlNodeType.SignificantWhitespace ||
                        reader.NodeType == XmlNodeType.Whitespace)
                    {
                        TextNode(reader);
                    }
                }
            }
        }

        /// <summary>
        /// Return an integer representing the total count of strings in the workbook. This count does not
        /// include any numbers, it counts only the total of text strings in the workbook.
        /// </summary>
        /// <return>the total count of strings in the workbook</return>
        public int Count => count;

        /// <summary>
        /// Returns an integer representing the total count of unique strings in the Shared String Table.
        /// A string is unique even if it is a copy of another string, but has different formatting applied
        /// at the character level.
        /// </summary>
        /// <return>the total count of unique strings in the workbook</return>
        public int UniqueCount => uniqueCount;

        /// <summary>
        /// Return the string at a given index.
        /// Formatting is ignored.
        /// </summary>
        /// <param name="idx">index of item to return.</param>
        /// <return>the item at the specified position in this Shared String table.</return>
        public String GetEntryAt(int idx)
        {
            return strings[idx];
        }

        public List<String> Items => strings;


        //// ContentHandler methods ////

        private StringBuilder characters;
        private bool tIsOpen;
        private bool inRPh;

        public void TextNode(XmlReader reader)
        {
            if(tIsOpen)
            {
                if(inRPh && includePhoneticRuns)
                {
                    characters.Append(reader.Value);
                }
                else if(!inRPh)
                {
                    characters.Append(reader.Value);
                }
            }
        }

        public void StartElement(XmlReader reader)
        {
            string uri = reader.NamespaceURI;
            string localName = reader.LocalName;
            //string name = reader.Name;
            if(uri != null && !uri.Equals(XSSFRelation.NS_SPREADSHEETML))
            {
                return;
            }
            if("sst".Equals(localName))
            {
                String count = reader.GetAttribute("count");
                if(count != null)
                    this.count = Int32.Parse(count);
                String uniqueCount = reader.GetAttribute("uniqueCount");
                if(uniqueCount != null)
                    this.uniqueCount = Int32.Parse(uniqueCount);

                this.strings = new List<String>(this.uniqueCount);
                this.phoneticStrings = new Dictionary<int, String>();
                characters = new StringBuilder();
            }
            else if("si".Equals(localName))
            {
                characters.Length = 0;
            }
            else if("t".Equals(localName))
            {
                tIsOpen = true;
            }
            else if("rPh".Equals(localName))
            {
                inRPh = true;
                //append space...this assumes that rPh always comes After regular <t>
                if(includePhoneticRuns && characters.Length > 0)
                {
                    characters.Append(" ");
                }
            }
        }

        public void EndElement(XmlReader reader)
        {
            string uri = reader.NamespaceURI;
            string localName = reader.LocalName;
            //string name = reader.Name;
            if(uri != null && !uri.Equals(XSSFRelation.NS_SPREADSHEETML))
            {
                return;
            }

            if("si".Equals(localName))
            {
                strings.Add(characters.ToString());
            }
            else if("t".Equals(localName))
            {
                tIsOpen = false;
            }
            else if("rPh".Equals(localName))
            {
                inRPh = false;
            }
        }
    }
}

