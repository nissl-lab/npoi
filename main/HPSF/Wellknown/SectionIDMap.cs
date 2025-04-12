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
    using System.Threading;
    using System.Collections.Generic;

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
        private static ThreadLocal<Dictionary<ClassID,PropertyIDMap>> defaultMap =
        new ThreadLocal<Dictionary<ClassID,PropertyIDMap>>();
        /**
         * The SummaryInformation's section's format ID.
         */
        public static ClassID SUMMARY_INFORMATION_ID =
            new ClassID("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}");

        /**
         * The DocumentSummaryInformation's first and second sections' format
         * ID.
         */
        private static readonly ClassID DOC_SUMMARY_INFORMATION =
            new ClassID("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}");    
        private static ClassID USER_DEFINED_PROPERTIES =
            new ClassID("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
    
        public static ClassID[] DOCUMENT_SUMMARY_INFORMATION_ID = {
            DOC_SUMMARY_INFORMATION, USER_DEFINED_PROPERTIES
        };

        /**
         * A property without a known name is described by this string. 
         */
        public const string UNDEFINED = "[undefined]";

        /// <summary>
        /// Returns the singleton instance of the default {@link
        /// SectionIDMap}.
        /// </summary>
        /// <returns>The instance value</returns>
        public static SectionIDMap GetInstance()
        {
            Dictionary<ClassID,PropertyIDMap> m = defaultMap.Value;
            if(m == null)
            {
                m = new Dictionary<ClassID, PropertyIDMap> {
                    {
                        SUMMARY_INFORMATION_ID,
                        PropertyIDMap.SummaryInformationProperties
                    },
                    {
                        DOCUMENT_SUMMARY_INFORMATION_ID[0],
                        PropertyIDMap.DocumentSummaryInformationProperties
                    }
                };
                defaultMap.Value = m;
            }
            return new SectionIDMap();
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
            if(m == null)
                return UNDEFINED;
            else
            {
                String s = (String)m.Get(pid);
                if(s == null)
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
            return (PropertyIDMap) this[Encoding.UTF8.GetString(sectionFormatID)];
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
    }
}