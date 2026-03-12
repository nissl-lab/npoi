/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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


namespace NPOI.XSSF.Model
{
    using System.Buffers;
    using System.Collections.Generic;
    using System.Globalization;
    using OpenXmlFormats.Spreadsheet;
    using System;
    using System.IO;
    using NPOI.OpenXml4Net.OPC;
    using System.Xml;
    using System.Text;

    /**
     * Table of strings shared across all sheets in a workbook.
     * <p>
     * A workbook may contain thousands of cells Containing string (non-numeric) data. Furthermore this data is very
     * likely to be repeated across many rows or columns. The goal of implementing a single string table that is shared
     * across the workbook is to improve performance in opening and saving the file by only Reading and writing the
     * repetitive information once.
     * </p>
     * <p>
     * Consider for example a workbook summarizing information for cities within various countries. There may be a
     * column for the name of the country, a column for the name of each city in that country, and a column
     * Containing the data for each city. In this case the country name is repetitive, being duplicated in many cells.
     * In many cases the repetition is extensive, and a tremendous savings is realized by making use of a shared string
     * table when saving the workbook. When displaying text in the spreadsheet, the cell table will just contain an
     * index into the string table as the value of a cell, instead of the full string.
     * </p>
     * <p>
     * The shared string table Contains all the necessary information for displaying the string: the text, formatting
     * properties, and phonetic properties (for East Asian languages).
     * </p>
     *
     * @author Nick Birch
     * @author Yegor Kozlov
     */
    public class SharedStringsTable : POIXMLDocumentPart
    {

        /**
         *  Array of individual string items in the Shared String table.
         */
        private readonly List<CT_Rst> strings = new List<CT_Rst>();

        /**
         *  Maps strings and their indexes in the <code>strings</code> arrays.
         *  Built lazily on the first call to AddEntry().
         */
        private readonly Dictionary<String, int> stmap = new Dictionary<String, int>();

        /**
         * An integer representing the total count of strings in the workbook. This count does not
         * include any numbers, it counts only the total of text strings in the workbook.
         */
        private int count;

        /**
         * An integer representing the total count of unique strings in the Shared String Table.
         * A string is unique even if it is a copy of another string, but has different formatting applied
         * at the character level.
         */
        private int uniqueCount;

        private SstDocument _sstDoc;

        // Lazy-loading state
        private PackagePart _loadPart;
        private bool _loaded;
        private bool _dirty;
        private bool _stmapBuilt;

        public SharedStringsTable()
            : base()
        {
            _sstDoc = new SstDocument();
            _sstDoc.AddNewSst();
            _loaded = true;
            _stmapBuilt = true;
        }

        public SharedStringsTable(PackagePart part)
            : base(part)
        {
            // Defer full parsing until first access (lazy loading).
            // However, perform an early security scan to detect DOCTYPE entity
            // expansion attacks (e.g. "billion laughs") in the SST, matching
            // the behaviour of ConvertStreamToXml / LoadXmlSafe.
            _loadPart = part;
            ValidateSstSecurity(part);
        }

        [Obsolete("deprecated in POI 3.14, scheduled for removal in POI 3.16")]
        public SharedStringsTable(PackagePart part, PackageRelationship rel)
             : this(part)
        {
        }

        /// <summary>
        /// Performs a lightweight security scan of the SST part to reject files that
        /// contain DOCTYPE entity declarations (e.g. "billion laughs" / XML bomb attacks).
        /// This is a fast, forward-only scan that does not build any object model.
        /// </summary>
        private static void ValidateSstSecurity(PackagePart part)
        {
            Stream stream = null;
            try
            {
                stream = part.GetInputStream();
                if (stream == null || (stream.CanSeek && stream.Length == 0))
                    return;

                var settings = new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                    IgnoreWhitespace = true,
                    IgnoreComments = true,
                    IgnoreProcessingInstructions = true,
                };
                using var reader = XmlReader.Create(stream, settings);
                // Reading up to the first Element node is sufficient:
                // DtdProcessing.Prohibit throws when <!DOCTYPE is encountered.
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element)
                        break;
                }
            }
            catch (XmlException e)
            {
                throw new IOException("unable to parse shared strings table", e);
            }
            finally
            {
                stream?.Dispose();
            }
        }


        private void EnsureLoaded()
        {
            if (_loaded) return;
            _loaded = true;
            if (_loadPart != null)
            {
                Stream stream = _loadPart.GetInputStream();
                try
                {
                    if (stream == null || (stream.CanSeek && stream.Length == 0))
                    {
                        _sstDoc = new SstDocument();
                        _sstDoc.AddNewSst();
                    }
                    else
                    {
                        try
                        {
                            ReadFromStream(stream);
                        }
                        catch (XmlException e)
                        {
                            throw new IOException("unable to parse shared strings table", e);
                        }
                    }
                }
                finally
                {
                    stream?.Dispose();
                }
            }
            else
            {
                _sstDoc = new SstDocument();
                _sstDoc.AddNewSst();
            }
        }

        /// <summary>
        /// Ensures stmap is consistent with the strings list. Called lazily before AddEntry.
        /// </summary>
        private void EnsureStmapBuilt()
        {
            if (_stmapBuilt) return;
            _stmapBuilt = true;
            for (int i = 0; i < strings.Count; i++)
            {
                string key = GetKey(strings[i]);
                if (key != null && !stmap.ContainsKey(key))
                    stmap[key] = i;
            }
        }

        /// <summary>
        /// Streaming XmlReader-based parser for sharedStrings.xml. Replaces the
        /// DOM-based ConvertStreamToXml + SstDocument.Parse approach to reduce
        /// allocations. Uses ArrayPool&lt;char&gt; for text buffering.
        /// </summary>
        private const int TextReadBufferSize = 1024;

        private void ReadFromStream(Stream stream)
        {
            _sstDoc = new SstDocument();
            _sstDoc.AddNewSst();
            CT_Sst sst = _sstDoc.GetSst();

            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                IgnoreWhitespace = false,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
            };

            char[] readBuf = ArrayPool<char>.Shared.Rent(TextReadBufferSize);
            try
            {
                using var reader = XmlReader.Create(stream, settings);

                CT_Rst currentSi = null;
                CT_RElt currentR = null;
                CT_RPrElt currentRPr = null;
                CT_PhoneticRun currentRPh = null;
                StringBuilder textBuf = null;
                bool inSiT = false, inRT = false, inRPhT = false;

                while (reader.Read())
                {
                    switch (reader.NodeType)
                    {
                        case XmlNodeType.Element:
                        {
                            string localName = reader.LocalName;
                            bool isEmpty = reader.IsEmptyElement;

                            // rPr children are all self-closing attribute-only elements
                            if (currentRPr != null && localName != "rPr")
                            {
                                ParseRPrChild(reader, localName, currentRPr);
                                break;
                            }

                            switch (localName)
                            {
                                case "sst":
                                {
                                    string cStr = reader.GetAttribute("count");
                                    if (cStr != null && int.TryParse(cStr, NumberStyles.None, CultureInfo.InvariantCulture, out int c))
                                        count = c;
                                    string ucStr = reader.GetAttribute("uniqueCount");
                                    if (ucStr != null && int.TryParse(ucStr, NumberStyles.None, CultureInfo.InvariantCulture, out int uc))
                                        uniqueCount = uc;
                                    break;
                                }
                                case "si":
                                    currentSi = new CT_Rst();
                                    break;
                                case "t":
                                    if (isEmpty)
                                    {
                                        if (currentR != null) currentR.t = string.Empty;
                                        else if (currentRPh != null) currentRPh.t = string.Empty;
                                        else if (currentSi != null) currentSi.t = string.Empty;
                                    }
                                    else
                                    {
                                        textBuf = new StringBuilder();
                                        if (currentR != null) inRT = true;
                                        else if (currentRPh != null) inRPhT = true;
                                        else if (currentSi != null) inSiT = true;
                                    }
                                    break;
                                case "r":
                                    currentR = new CT_RElt();
                                    if (currentSi != null)
                                    {
                                        if (currentSi.r == null) currentSi.r = new List<CT_RElt>();
                                        currentSi.r.Add(currentR);
                                    }
                                    if (isEmpty) currentR = null;
                                    break;
                                case "rPr":
                                    currentRPr = new CT_RPrElt();
                                    if (currentR != null) currentR.rPr = currentRPr;
                                    if (isEmpty) currentRPr = null;
                                    break;
                                case "rPh":
                                {
                                    currentRPh = new CT_PhoneticRun();
                                    string sbStr = reader.GetAttribute("sb");
                                    string ebStr = reader.GetAttribute("eb");
                                    if (sbStr != null && uint.TryParse(sbStr, out uint sbVal)) currentRPh.sb = sbVal;
                                    if (ebStr != null && uint.TryParse(ebStr, out uint ebVal)) currentRPh.eb = ebVal;
                                    if (currentSi != null)
                                    {
                                        if (currentSi.rPh == null) currentSi.rPh = new List<CT_PhoneticRun>();
                                        currentSi.rPh.Add(currentRPh);
                                    }
                                    if (isEmpty) currentRPh = null;
                                    break;
                                }
                                case "phoneticPr":
                                    if (currentSi != null)
                                        currentSi.phoneticPr = ParsePhoneticPrAttributes(reader);
                                    break;
                            }
                            break;
                        }
                        case XmlNodeType.Text:
                        case XmlNodeType.SignificantWhitespace:
                        case XmlNodeType.Whitespace:
                        {
                            if ((inSiT || inRT || inRPhT) && textBuf != null)
                            {
                                int charsRead;
                                while ((charsRead = reader.ReadValueChunk(readBuf, 0, readBuf.Length)) > 0)
                                    textBuf.Append(readBuf, 0, charsRead);
                            }
                            break;
                        }
                        case XmlNodeType.EndElement:
                        {
                            switch (reader.LocalName)
                            {
                                case "si":
                                    if (currentSi != null)
                                    {
                                        sst.si.Add(currentSi);
                                        strings.Add(currentSi);
                                    }
                                    currentSi = null;
                                    break;
                                case "t":
                                {
                                    string text = textBuf?.ToString() ?? string.Empty;
                                    textBuf = null;
                                    if (inSiT && currentSi != null) { currentSi.t = text; inSiT = false; }
                                    else if (inRT && currentR != null) { currentR.t = text; inRT = false; }
                                    else if (inRPhT && currentRPh != null) { currentRPh.t = text; inRPhT = false; }
                                    break;
                                }
                                case "r":
                                    currentR = null;
                                    break;
                                case "rPr":
                                    currentRPr = null;
                                    break;
                                case "rPh":
                                    currentRPh = null;
                                    break;
                            }
                            break;
                        }
                    }
                }
            }
            finally
            {
                ArrayPool<char>.Shared.Return(readBuf);
            }
        }

        private static CT_PhoneticPr ParsePhoneticPrAttributes(XmlReader reader)
        {
            var pr = new CT_PhoneticPr();
            string fontId = reader.GetAttribute("fontId");
            if (fontId != null && uint.TryParse(fontId, out uint fid)) pr.fontId = fid;
            string type = reader.GetAttribute("type");
            if (type != null && Enum.TryParse(type, out ST_PhoneticType pt)) pr.type = pt;
            string alignment = reader.GetAttribute("alignment");
            if (alignment != null && Enum.TryParse(alignment, out ST_PhoneticAlignment pa)) pr.alignment = pa;
            return pr;
        }

        private static void ParseRPrChild(XmlReader reader, string localName, CT_RPrElt rPr)
        {
            switch (localName)
            {
                case "sz":
                {
                    string val = reader.GetAttribute("val");
                    if (val != null && double.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out double sz))
                        rPr.sz = new CT_FontSize { val = sz };
                    break;
                }
                case "color":
                    rPr.color = ParseColorAttributes(reader);
                    break;
                case "rFont":
                {
                    string val = reader.GetAttribute("val");
                    if (val != null) rPr.rFont = new CT_FontName { val = val };
                    break;
                }
                case "family":
                {
                    string val = reader.GetAttribute("val");
                    if (val != null && int.TryParse(val, out int fam))
                        rPr.family = new CT_IntProperty { val = fam };
                    break;
                }
                case "charset":
                {
                    string val = reader.GetAttribute("val");
                    if (val != null && int.TryParse(val, out int cs))
                        rPr.charset = new CT_IntProperty { val = cs };
                    break;
                }
                case "b":
                    rPr.b = ParseBoolProp(reader);
                    break;
                case "i":
                    rPr.i = ParseBoolProp(reader);
                    break;
                case "strike":
                    rPr.strike = ParseBoolProp(reader);
                    break;
                case "outline":
                    rPr.outline = ParseBoolProp(reader);
                    break;
                case "shadow":
                    rPr.shadow = ParseBoolProp(reader);
                    break;
                case "condense":
                    rPr.condense = ParseBoolProp(reader);
                    break;
                case "extend":
                    rPr.extend = ParseBoolProp(reader);
                    break;
                case "u":
                {
                    string val = reader.GetAttribute("val") ?? "single";
                    if (Enum.TryParse(val, out ST_UnderlineValues uv))
                        rPr.u = new CT_UnderlineProperty { val = uv };
                    break;
                }
                case "vertAlign":
                {
                    string val = reader.GetAttribute("val");
                    if (val != null && Enum.TryParse(val, out ST_VerticalAlignRun va))
                        rPr.vertAlign = new CT_VerticalAlignFontProperty { val = va };
                    break;
                }
                case "scheme":
                {
                    string val = reader.GetAttribute("val");
                    if (val != null && Enum.TryParse(val, out ST_FontScheme fs))
                        rPr.scheme = new CT_FontScheme { val = fs };
                    break;
                }
            }
        }

        private static CT_BooleanProperty ParseBoolProp(XmlReader reader)
        {
            string val = reader.GetAttribute("val");
            // Default for boolean property is true when attribute is absent
            bool v = val == null || (val != "0" && !val.Equals("false", StringComparison.OrdinalIgnoreCase));
            return new CT_BooleanProperty { val = v };
        }

        private static CT_Color ParseColorAttributes(XmlReader reader)
        {
            var color = new CT_Color();
            string auto = reader.GetAttribute("auto");
            if (auto != null)
                color.auto = auto != "0" && !auto.Equals("false", StringComparison.OrdinalIgnoreCase);
            string indexed = reader.GetAttribute("indexed");
            if (indexed != null && uint.TryParse(indexed, out uint idx)) color.indexed = idx;
            string rgb = reader.GetAttribute("rgb");
            if (rgb != null && rgb.Length >= 2)
            {
                byte[] bytes = new byte[rgb.Length / 2];
                for (int i = 0; i < bytes.Length; i++)
                    bytes[i] = Convert.ToByte(rgb.Substring(i * 2, 2), 16);
                color.rgb = bytes;
            }
            string theme = reader.GetAttribute("theme");
            if (theme != null && uint.TryParse(theme, out uint th)) color.theme = th;
            string tint = reader.GetAttribute("tint");
            if (tint != null && double.TryParse(tint, NumberStyles.Any, CultureInfo.InvariantCulture, out double t)) color.tint = t;
            return color;
        }

        /// <summary>
        /// Read shared strings from a stream. Kept for backward compatibility; internally
        /// delegates to the streaming parser.
        /// </summary>
        public void ReadFrom(Stream is1)
        {
            try
            {
                ReadFromStream(is1);
                _loaded = true;
                _stmapBuilt = false;
            }
            catch (XmlException e)
            {
                throw new IOException("unable to parse shared strings table", e);
            }
        }

        private static String GetKey(CT_Rst st)
        {
            return st.XmlText;
        }

        /**
         * Return a string item by index
         *
         * @param idx index of item to return.
         * @return the item at the specified position in this Shared String table.
         */
        public CT_Rst GetEntryAt(int idx)
        {
            EnsureLoaded();
            return strings[idx];
        }

        /**
         * Return an integer representing the total count of strings in the workbook. This count does not
         * include any numbers, it counts only the total of text strings in the workbook.
         *
         * @return the total count of strings in the workbook
         */
        public int Count
        {
            get
            {
                EnsureLoaded();
                return count;
            }
        }

        /**
         * Returns an integer representing the total count of unique strings in the Shared String Table.
         * A string is unique even if it is a copy of another string, but has different formatting applied
         * at the character level.
         *
         * @return the total count of unique strings in the workbook
         */
        public int UniqueCount
        {
            get
            {
                EnsureLoaded();
                return uniqueCount;
            }
        }

        /**
         * Add an entry to this Shared String table (a new value is appended to the end).
         *
         * <p>
         * If the Shared String table already Contains this <code>CT_Rst</code> bean, its index is returned.
         * Otherwise a new entry is added.
         * </p>
         *
         * @param st the entry to add
         * @return index the index of Added entry
         */
        public int AddEntry(CT_Rst st)
        {
            EnsureLoaded();
            EnsureStmapBuilt();
            String s = GetKey(st);
            count++;
            if (stmap.TryGetValue(s, out int entry))
            {
                return entry;
            }

            uniqueCount++;
            //create a CT_Rst bean attached to this SstDocument and copy the argument CT_Rst into it
            CT_Rst newSt = new CT_Rst();
            _sstDoc.GetSst().si.Add(newSt);
            newSt.Set(st);
            int idx = strings.Count;
            stmap[s] = idx;
            strings.Add(newSt);
            _dirty = true;
            return idx;
        }

        /**
         * Provide low-level access to the underlying array of CT_Rst beans
         *
         * @return array of CT_Rst beans
         */
        public IList<CT_Rst> Items
        {
            get
            {
                EnsureLoaded();
                return strings.AsReadOnly();
            }
        }

        /**
         * Write this table out as XML.
         *
         * @param out The stream to write to.
         * @throws IOException if an error occurs while writing.
         */
        public void WriteTo(Stream out1)
        {
            EnsureLoaded();
            CT_Sst sst = _sstDoc.GetSst();
            sst.count = count;
            sst.uniqueCount = uniqueCount;
            _sstDoc.Save(out1);
        }

        /// <summary>
        /// Returns true if the SST has been parsed from its backing part.
        /// Used in tests to verify lazy-load behaviour.
        /// </summary>
        internal bool IsLoaded => _loaded;

        /// <summary>
        /// Prepares the part for commit. No-op when SST has not been modified,
        /// preserving the original part bytes.
        /// </summary>
        protected internal override void PrepareForCommit()
        {
            if (_dirty)
                base.PrepareForCommit();
        }

        /// <summary>
        /// Commits the SST to the package. No-op when SST has not been modified,
        /// so the original sharedStrings.xml bytes are preserved without parsing.
        /// </summary>
        protected internal override void Commit()
        {
            if (!_dirty) return;
            PackagePart part = GetPackagePart();
            Stream out1 = part.GetOutputStream();
            WriteTo(out1);
            out1.Close();
        }
    }
}






