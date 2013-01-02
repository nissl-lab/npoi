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
namespace NPOI.HPSF.Extractor
{
    using System;
    using System.Text;
    using System.IO;
    using System.Collections;

    using NPOI;
    using NPOI.HPSF;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;
    using System.Globalization;

    /// <summary>
    /// Extracts all of the HPSF properties, both
    /// build in and custom, returning them in
    /// textual form.
    /// </summary>
    public class HPSFPropertiesExtractor : POITextExtractor
    {
        public HPSFPropertiesExtractor(POITextExtractor mainExtractor)
            : base(mainExtractor)
        {

        }
        public HPSFPropertiesExtractor(POIDocument doc)
            : base(doc)
        {

        }
        public HPSFPropertiesExtractor(POIFSFileSystem fs)
            : base(new PropertiesOnlyDocument(fs))
        {

        }
        public HPSFPropertiesExtractor(NPOIFSFileSystem fs)
            : base(new PropertiesOnlyDocument(fs))
        {
            
        }
        /// <summary>
        /// Gets the document summary information text.
        /// </summary>
        /// <value>The document summary information text.</value>
        public String DocumentSummaryInformationText
        {
            get
            {
                DocumentSummaryInformation dsi = document.DocumentSummaryInformation;
                StringBuilder text = new StringBuilder();

                // Normal properties
                text.Append(GetPropertiesText(dsi));

                // Now custom ones
                CustomProperties cps = dsi == null ? null : dsi.CustomProperties;
                
                if (cps != null)
                {
                    IEnumerator keys = cps.NameSet().GetEnumerator();
                    while (keys.MoveNext())
                    {
                        String key = keys.Current.ToString();
                        String val = GetPropertyValueText(cps[key]);
                        text.Append(key + " = " + val + "\n");
                    }
                }
                // All done
                return text.ToString();
            }
        }
        /// <summary>
        /// Gets the summary information text.
        /// </summary>
        /// <value>The summary information text.</value>
        public String SummaryInformationText
        {
            get
            {
                SummaryInformation si = document.SummaryInformation;

                // Just normal properties
                return GetPropertiesText(si);
            }
        }

        /// <summary>
        /// Gets the properties text.
        /// </summary>
        /// <param name="ps">The ps.</param>
        /// <returns></returns>
        private static String GetPropertiesText(SpecialPropertySet ps)
        {
            if (ps == null)
            {
                // Not defined, oh well
                return "";
            }

            StringBuilder text = new StringBuilder();

            Wellknown.PropertyIDMap idMap = ps.PropertySetIDMap;
            Property[] props = ps.Properties;
            for (int i = 0; i < props.Length; i++)
            {
                String type = props[i].ID.ToString(CultureInfo.InvariantCulture);
                Object typeObj = idMap.Get(props[i].ID);
                if (typeObj != null)
                {
                    type = typeObj.ToString();
                }

                String val = GetPropertyValueText(props[i].Value);
                text.Append(type + " = " + val + "\n");
            }

            return text.ToString();
        }
        /// <summary>
        /// Gets the property value text.
        /// </summary>
        /// <param name="val">The val.</param>
        /// <returns></returns>
        private static String GetPropertyValueText(Object val)
        {
            if (val == null)
            {
                return "(not set)";
            }
            if (val is byte[])
            {
                byte[] b = (byte[])val;
                if (b.Length == 0)
                {
                    return "";
                }
                if (b.Length == 1)
                {
                    return b[0].ToString(CultureInfo.InvariantCulture);
                }
                if (b.Length == 2)
                {
                    return LittleEndian.GetUShort(b).ToString(CultureInfo.InvariantCulture);
                }
                if (b.Length == 4)
                {
                    return LittleEndian.GetUInt(b).ToString(CultureInfo.InvariantCulture);
                }
                // Maybe it's a string? who knows!
                return b.ToString();
            }

            return val.ToString();
        }

        /// <summary>
        /// Return the text of all the properties defined in
        /// the document.
        /// </summary>
        /// <value>All the text from the document.</value>
        public override String Text
        {
            get
            {
                return SummaryInformationText + DocumentSummaryInformationText;
            }
        }

        /// <summary>
        /// Returns another text extractor, which is able to
        /// output the textual content of the document
        /// metadata / properties, such as author and title.
        /// </summary>
        /// <value>The metadata text extractor.</value>
        public override POITextExtractor MetadataTextExtractor
        {
            get
            {
                throw new InvalidOperationException("You already have the Metadata Text Extractor, not recursing!");
            }
        }

        /// <summary>
        /// So we can get at the properties of any
        /// random OLE2 document.
        /// </summary>
        private class PropertiesOnlyDocument : POIDocument
        {
            public PropertiesOnlyDocument(NPOIFSFileSystem fs)
                : base(fs.Root)
            {
                
            }
            public PropertiesOnlyDocument(POIFSFileSystem fs)
                : base(fs)
            {

            }

            public override void Write(Stream out1)
            {
                throw new InvalidOperationException("Unable to write, only for properties!");
            }
        }
    }
}