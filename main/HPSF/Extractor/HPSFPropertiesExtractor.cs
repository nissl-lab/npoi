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
            : base(new HPSFPropertiesOnlyDocument(fs))
        {

        }
        public HPSFPropertiesExtractor(NPOIFSFileSystem fs)
            : base(new HPSFPropertiesOnlyDocument(fs))
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
                if (document == null)
                {  // event based extractor does not have a document
                    return "";
                }
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
                        String val = HelperPropertySet.GetPropertyValueText(cps[key]);
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
                if (document == null)
                {  // event based extractor does not have a document
                    return "";
                }
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

                String val = HelperPropertySet.GetPropertyValueText(props[i].Value);
                text.Append(type + " = " + val + "\n");
            }

            return text.ToString();
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

        public abstract class HelperPropertySet : SpecialPropertySet
        {
            public HelperPropertySet()
                : base(null)
            {
            }
            public static String GetPropertyValueText(Object val)
            {
                if (val == null)
                {
                    return "(not set)";
                }
                return SpecialPropertySet.GetPropertyStringValue(val);
            }
        }
    }
}