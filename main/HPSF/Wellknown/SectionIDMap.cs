/* ====================================================================
   Licensed To the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file To You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed To in writing, software
   distributed under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

/* ================================================================
 * About NPOI
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * Author's Blog: tonyqus.wordpress.com.cn (wp.tonyqus.cn)
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/

namespace NPOI.HPSF.Wellknown
{
    using System;
    using System.Text;
    using System.Collections;

    /// <summary>
    /// Maps section format IDs To {@link PropertyIDMap}s. It Is
    /// initialized with two well-known section format IDs: those of the
    /// <c>\005SummaryInformation</c> stream and the
    /// <c>\005DocumentSummaryInformation</c> stream.
    /// If you have a section format ID you can use it as a key To query
    /// this map.  If you Get a {@link PropertyIDMap} returned your section
    /// is well-known and you can query the {@link PropertyIDMap} for PID
    /// strings. If you Get back <c>null</c> you are on your own.
    /// This {@link java.util.Map} expects the byte arrays of section format IDs
    /// as keys. A key maps To a {@link PropertyIDMap} describing the
    /// property IDs in sections with the specified section format ID.
    /// @author Rainer Klute (klute@rainer-klute.de)
    /// @since 2002-02-09
    /// </summary>
    public class SectionIDMap : Hashtable
    {

        /**
         * The SummaryInformation's section's format ID.
         */
        public static readonly byte[] SUMMARY_INFORMATION_ID = new byte[]
        {
            (byte) 0xF2, (byte) 0x9F, (byte) 0x85, (byte) 0xE0,
            (byte) 0x4F, (byte) 0xF9, (byte) 0x10, (byte) 0x68,
            (byte) 0xAB, (byte) 0x91, (byte) 0x08, (byte) 0x00,
            (byte) 0x2B, (byte) 0x27, (byte) 0xB3, (byte) 0xD9
        };

        /**
         * The DocumentSummaryInformation's first and second sections' format
         * ID.
         */
        public static readonly byte[] DOCUMENT_SUMMARY_INFORMATION_ID1 =
        {
                (byte) 0xD5, (byte) 0xCD, (byte) 0xD5, (byte) 0x02,
                (byte) 0x2E, (byte) 0x9C, (byte) 0x10, (byte) 0x1B,
                (byte) 0x93, (byte) 0x97, (byte) 0x08, (byte) 0x00,
                (byte) 0x2B, (byte) 0x2C, (byte) 0xF9, (byte) 0xAE
            };
        public static readonly byte[] DOCUMENT_SUMMARY_INFORMATION_ID2 =
            {
                (byte) 0xD5, (byte) 0xCD, (byte) 0xD5, (byte) 0x05,
                (byte) 0x2E, (byte) 0x9C, (byte) 0x10, (byte) 0x1B,
                (byte) 0x93, (byte) 0x97, (byte) 0x08, (byte) 0x00,
                (byte) 0x2B, (byte) 0x2C, (byte) 0xF9, (byte) 0xAE
            };

        /**
         * A property without a known name is described by this string. 
         */
        public const string UNDEFINED = "[undefined]";

        /**
         * The default section ID map. It maps section format IDs To
         * {@link PropertyIDMap}s.
         */
        private static SectionIDMap defaultMap;



        /// <summary>
        /// Returns the singleton instance of the default {@link
        /// SectionIDMap}.
        /// </summary>
        /// <returns>The instance value</returns>
        public static SectionIDMap GetInstance()
        {
            if (defaultMap == null)
            {
                SectionIDMap m = new SectionIDMap();
                m.Put(SUMMARY_INFORMATION_ID,
                      PropertyIDMap.SummaryInformationProperties);
                m.Put(DOCUMENT_SUMMARY_INFORMATION_ID1,
                      PropertyIDMap.DocumentSummaryInformationProperties);
                defaultMap = m;
            }
            return defaultMap;
        }



        /// <summary>
        /// Returns the property ID string that is associated with a
        /// given property ID in a section format ID's namespace.
        /// </summary>
        /// <param name="sectionFormatID">Each section format ID has its own name
        /// space of property ID strings and thus must be specified.</param>
        /// <param name="pid">The property ID</param>
        /// <returns>The well-known property ID string associated with the
        /// property ID pid in the name space spanned by sectionFormatID If the pid
        /// sectionFormatID combination is not well-known, the
        /// string "[undefined]" is returned.
        /// </returns>
        public static String GetPIDString(byte[] sectionFormatID,
                                          long pid)
        {
            PropertyIDMap m = GetInstance().Get(sectionFormatID);
            if (m == null)
                return UNDEFINED;
            else
            {
                String s = (String)m.Get(pid);
                if (s == null)
                    return UNDEFINED;
                return s;
            }
        }



        /// <summary>
        /// Returns the {@link PropertyIDMap} for a given section format
        /// ID.
        /// </summary>
        /// <param name="sectionFormatID">The section format ID.</param>
        /// <returns>the property ID map</returns>
        public PropertyIDMap Get(byte[] sectionFormatID)
        {
            return (PropertyIDMap)this[Encoding.UTF8.GetString(sectionFormatID)];
        }



        /// <summary>
        /// Returns the {@link PropertyIDMap} for a given section format
        /// ID.
        /// </summary>
        /// <param name="sectionFormatID">A section format ID as a 
        /// <c>byte[]</c></param>
        /// <returns>the property ID map</returns>
        public Object Get(Object sectionFormatID)
        {
            return Get((byte[])sectionFormatID);
        }



        /// <summary>
        /// Associates a section format ID with a {@link
        /// PropertyIDMap}.
        /// </summary>
        /// <param name="sectionFormatID">the section format ID</param>
        /// <param name="propertyIDMap">The property ID map.</param>
        /// <returns></returns>
        public Object Put(byte[] sectionFormatID,
                          PropertyIDMap propertyIDMap)
        {
            return this[sectionFormatID] = propertyIDMap;
        }



        /// <summary>
        /// Puts the specified key.
        /// </summary>
        /// <param name="key">This parameter remains undocumented since the method Is
        /// deprecated.</param>
        /// <param name="value">This parameter remains undocumented since the method Is
        /// deprecated.</param>
        /// <returns>The return value remains undocumented since the method Is
        /// deprecated.</returns>
        public Object Put(Object key, Object value)
        {
            return Put((byte[])key, (PropertyIDMap)value);
        }

    }
}