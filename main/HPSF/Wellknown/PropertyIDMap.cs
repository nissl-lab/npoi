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
    using System.Collections;
    using System.Collections.Generic;


    /// <summary>
    /// This is a dictionary which maps property ID values To property
    /// ID strings.
    /// The methods {@link #GetSummaryInformationProperties} and {@link
    /// #GetDocumentSummaryInformationProperties} return singleton {@link
    /// PropertyIDMap}s. An application that wants To extend these maps
    /// should treat them as unmodifiable, copy them and modifiy the
    /// copies.
    /// @author Rainer Klute 
    /// <a href="mailto:klute@rainer-klute.de">&lt;klute@rainer-klute.de&gt;</a>
    /// @since 2002-02-09
    /// </summary>
    public class PropertyIDMap : IDictionary<long, string>
    {

        /*
         * The following definitions are for property IDs in the first
         * (and only) section of the Summary Information property Set.
         */

        /// <summary>
        /// ID of the property that denotes the document's title */
        /// </summary>
        public const int PID_TITLE = 2;

        /// <summary>
        /// ID of the property that denotes the document's subject */
        /// </summary>
        public const int PID_SUBJECT = 3;

        /// <summary>
        /// ID of the property that denotes the document's author */
        /// </summary>
        public const int PID_AUTHOR = 4;

        /// <summary>
        /// ID of the property that denotes the document's keywords */
        /// </summary>
        public const int PID_KEYWORDS = 5;

        /// <summary>
        /// ID of the property that denotes the document's comments */
        /// </summary>
        public const int PID_COMMENTS = 6;

        /// <summary>
        /// ID of the property that denotes the document's template */
        /// </summary>
        public const int PID_TEMPLATE = 7;

        /// <summary>
        /// ID of the property that denotes the document's last author */
        /// </summary>
        public const int PID_LASTAUTHOR = 8;

        /// <summary>
        /// ID of the property that denotes the document's revision number */
        /// </summary>
        public const int PID_REVNUMBER = 9;

        /// <summary>
        /// ID of the property that denotes the document's edit time */
        /// </summary>
        public const int PID_EDITTIME = 10;

        /// <summary>
        /// ID of the property that denotes the date and time the document was
        /// last printed */
        /// </summary>
        public const int PID_LASTPRINTED = 11;

        /// <summary>
        /// ID of the property that denotes the date and time the document was
        /// created. */
        /// </summary>
        public const int PID_CREATE_DTM = 12;

        /// <summary>
        /// ID of the property that denotes the date and time the document was
        /// saved */
        /// </summary>
        public const int PID_LASTSAVE_DTM = 13;

        /// <summary>
        /// ID of the property that denotes the number of pages in the
        /// document */
        /// </summary>
        public const int PID_PAGECOUNT = 14;

        /// <summary>
        /// ID of the property that denotes the number of words in the
        /// document */
        /// </summary>
        public const int PID_WORDCOUNT = 15;

        /// <summary>
        /// ID of the property that denotes the number of characters in the
        /// document */
        /// </summary>
        public const int PID_CHARCOUNT = 16;

        /// <summary>
        /// ID of the property that denotes the document's thumbnail */
        /// </summary>
        public const int PID_THUMBNAIL = 17;

        /// <summary>
        /// ID of the property that denotes the application that created the
        /// document */
        /// </summary>
        public const int PID_APPNAME = 18;

        /// <summary>
        /// <para>
        /// ID of the property that denotes whether read/write access to the
        /// document is allowed or whether is should be opened as read-only. It can
        /// have the following values:
        /// </para>
        /// <para>
        /// <list type="table">
        /// <listheader><term>Value</term><term>Description</term><term>0</term><term>No restriction</term><term>2</term><term>Read-only recommended</term><term>4</term><description>Read-only enforced</description></listheader>
        /// <item><term>
        ///    </term><term>Value</term><term>
        ///    </term><term>Description</term><description>
        ///   </description></item>
        /// <item><term>
        ///    </term><term>0</term><term>
        ///    </term><term>No restriction</term><description>
        ///   </description></item>
        /// <item><term>
        ///    </term><term>2</term><term>
        ///    </term><term>Read-only recommended</term><description>
        ///   </description></item>
        /// <item><term>
        ///    </term><term>4</term><term>
        ///    </term><term>Read-only enforced</term><description>
        ///   </description></item>
        /// </para>
        /// </summary>
        public const int PID_SECURITY = 19;



        /*
         * The following definitions are for property IDs in the first
         * section of the Document Summary Information property Set.
         */

        /// <summary>
        /// The entry is a dictionary.
        /// </summary>
        public const int PID_DICTIONARY = 0;

        /// <summary>
        /// The entry denotes a code page.
        /// </summary>
        public const  int PID_CODEPAGE = 1;

        /// <summary>
        /// The entry is a string denoting the category the file belongs
        /// to, e.g. review, memo, etc. This is useful to find documents of
        /// same type.
        /// </summary>
        public const int PID_CATEGORY = 2;

        /// <summary>
        /// Target format for power point presentation, e.g. 35mm,
        /// printer, video etc.
        /// </summary>
        public const int PID_PRESFORMAT = 3;

        /// <summary>
        /// Number of bytes.
        /// </summary>
        public const int PID_BYTECOUNT = 4;

        /// <summary>
        /// Number of lines.
        /// </summary>
        public const int PID_LINECOUNT = 5;

        /// <summary>
        /// Number of paragraphs.
        /// </summary>
        public const int PID_PARCOUNT = 6;

        /// <summary>
        /// Number of slides in a power point presentation.
        /// </summary>
        public const int PID_SLIDECOUNT = 7;

        /// <summary>
        /// Number of slides with notes.
        /// </summary>
        public const int PID_NOTECOUNT = 8;

        /// <summary>
        /// Number of hidden slides.
        /// </summary>
        public const int PID_HIDDENCOUNT = 9;

        /// <summary>
        /// Number of multimedia clips, e.g. sound or video.
        /// </summary>
        public const int PID_MMCLIPCOUNT = 10;

        /// <summary>
        /// This entry is Set to -1 when scaling of the thumbnail is
        /// desired. Otherwise the thumbnail should be cropped.
        /// </summary>
        public const int PID_SCALE = 11;

        /// <summary>
        /// This entry denotes an internally used property. It is a
        /// vector of variants consisting of pairs of a string (VT_LPSTR)
        /// and a number (VT_I4). The string is a heading name, and the
        /// number tells how many document parts are under that
        /// heading.
        /// </summary>
        public const int PID_HEADINGPAIR = 12;

        /// <summary>
        /// This entry contains the names of document parts (word: names
        /// of the documents in the master document, excel: sheet names,
        /// power point: slide titles, binder: document names).
        /// </summary>
        public const int PID_DOCPARTS = 13;

        /// <summary>
        /// This entry contains the name of the project manager.
        /// </summary>
        public const int PID_MANAGER = 14;

        /// <summary>
        /// This entry contains the company name.
        /// </summary>
        public const int PID_COMPANY = 15;

        /// <summary>
        /// If this entry is -1 the links are dirty and should be
        /// re-evaluated.
        /// </summary>
        public const int PID_LINKSDIRTY = 0x10;
    
        /// <summary>
        /// The entry specifies an estimate of the number of characters
        ///  in the document, including whitespace, as an integer
        /// </summary>
        public const int PID_CCHWITHSPACES = 0x11;
    
        // 0x12 Unused
        // 0x13 GKPIDDSI_SHAREDDOC - Must be False
        // 0x14 GKPIDDSI_LINKBASE - Must not be written
        // 0x15 GKPIDDSI_HLINKS - Must not be written

        /// <summary>
        /// This entry contains a bool which marks if the User Defined
        ///  Property Set has been updated outside of the Application, if so the
        ///  hyperlinks should be updated on document load.
        /// </summary>
        public const int PID_HYPERLINKSCHANGED = 0x16;
    
        /// <summary>
        /// This entry contains the version of the Application which wrote the
        ///  Property Set, stored with the two high order bytes having the major
        ///  version number, and the two low order bytes the minor version number.
        /// </summary>
        public const int PID_VERSION = 0x17;
    
        /// <summary>
        /// This entry contains the VBA digital signature for the VBA project
        ///  embedded in the document.
        /// </summary>
        public const int PID_DIGSIG = 0x18;
    
        // 0x19 Unused
    
        /// <summary>
        /// This entry contains a string of the content type of the file.
        /// </summary>
        public const int PID_CONTENTTYPE = 0x1A;
    
        /// <summary>
        /// This entry contains a string of the document status.
        /// </summary>
        public const int PID_CONTENTSTATUS = 0x1B;
    
        /// <summary>
        /// This entry contains a string of the document language, but
        ///  normally should be empty.
        /// </summary>
        public const int PID_LANGUAGE = 0x1C;
    
        /// <summary>
        /// This entry contains a string of the document version, but
        ///  normally should be empty
        /// </summary>
        public const int PID_DOCVERSION = 0x1D;
    
        /// <summary>
        /// The highest well-known property ID. Applications are free to use
        ///  higher values for custom purposes. (This value is based on Office 12,
        ///  earlier versions of Office had lower values)
        /// </summary>
        public const int PID_MAX = 0x1F;

        /// <summary>
        /// The Locale property, if present, MUST have the property identifier 0x80000000,
        /// MUST NOT have a property name, and MUST have type VT_UI4 (0x0013).
        /// If present, its value MUST be a valid language code identifier as specified in [MS-LCID].
        /// Its value is selected in an implementation-specific manner.
        /// </summary>
        public const  int PID_LOCALE = unchecked((int) 0x80000000);


        /// <summary>
        /// <para>
        /// The Behavior property, if present, MUST have the property identifier 0x80000003,
        /// MUST NOT have a property name, and MUST have type VT_UI4 (0x0013).
        /// A version 0 property Set, indicated by the value 0x0000 for the Version field of
        /// the PropertySetStream packet, MUST NOT have a Behavior property.
        /// If the Behavior property is present, it MUST have one of the following values.
        /// </para>
        /// <para>
        /// <list type="bullet">
        /// <item><description>0x00000000 = Property names are case-insensitive (default)</description></item>
        /// <item><description>0x00000001 = Property names are case-sensitive.</description></item>
        /// </list>
        /// </para>
        /// </summary>
        public const int PID_BEHAVIOUR = unchecked((int)0x80000003);

        /// <summary>
        /// Contains the summary information property ID values and
        /// associated strings. See the overall HPSF documentation for
        /// details!
        /// </summary>
        private static PropertyIDMap summaryInformationProperties;
        private static Object[][] summaryInformationIdValues = [
            [(long)PID_TITLE, "PID_TITLE"],
            [(long)PID_SUBJECT, "PID_SUBJECT"],
            [(long)PID_AUTHOR, "PID_AUTHOR"],
            [(long)PID_KEYWORDS, "PID_KEYWORDS"],
            [(long)PID_COMMENTS, "PID_COMMENTS"],
            [(long)PID_TEMPLATE, "PID_TEMPLATE"],
            [(long)PID_LASTAUTHOR, "PID_LASTAUTHOR"],
            [(long)PID_REVNUMBER, "PID_REVNUMBER"],
            [(long)PID_EDITTIME, "PID_EDITTIME"],
            [(long)PID_LASTPRINTED, "PID_LASTPRINTED"],
            [(long)PID_CREATE_DTM, "PID_CREATE_DTM"],
            [(long)PID_LASTSAVE_DTM, "PID_LASTSAVE_DTM"],
            [(long)PID_PAGECOUNT, "PID_PAGECOUNT"],
            [(long)PID_WORDCOUNT, "PID_WORDCOUNT"],
            [(long)PID_CHARCOUNT, "PID_CHARCOUNT"],
            [(long)PID_THUMBNAIL, "PID_THUMBNAIL"],
            [(long)PID_APPNAME, "PID_APPNAME"],
            [(long)PID_SECURITY, "PID_SECURITY"],
        ];
        /// <summary>
        /// Contains the summary information property ID values and
        /// associated strings. See the overall HPSF documentation for
        /// details!
        /// </summary>
        private static PropertyIDMap documentSummaryInformationProperties;
        private static Object[][] documentSummaryInformationIdValues = [
            [(long)PID_DICTIONARY, "PID_DICTIONARY" ],
            [(long)PID_CODEPAGE, "PID_CODEPAGE" ],
            [(long)PID_CATEGORY, "PID_CATEGORY" ],
            [(long)PID_PRESFORMAT, "PID_PRESFORMAT" ],
            [(long)PID_BYTECOUNT, "PID_BYTECOUNT" ],
            [(long)PID_LINECOUNT, "PID_LINECOUNT" ],
            [(long)PID_PARCOUNT, "PID_PARCOUNT" ],
            [(long)PID_SLIDECOUNT, "PID_SLIDECOUNT" ],
            [(long)PID_NOTECOUNT, "PID_NOTECOUNT" ],
            [(long)PID_HIDDENCOUNT, "PID_HIDDENCOUNT" ],
            [(long)PID_MMCLIPCOUNT, "PID_MMCLIPCOUNT" ],
            [(long)PID_SCALE, "PID_SCALE" ],
            [(long)PID_HEADINGPAIR, "PID_HEADINGPAIR" ],
            [(long)PID_DOCPARTS, "PID_DOCPARTS" ],
            [(long)PID_MANAGER, "PID_MANAGER" ],
            [(long)PID_COMPANY, "PID_COMPANY" ],
            [(long)PID_LINKSDIRTY, "PID_LINKSDIRTY" ],
        ];

        /// <summary>
        /// Contains the fallback property ID values and associated strings.
        /// This is only used for lookups and not for initializing a property Set
        /// </summary>
        private static PropertyIDMap fallbackProperties;
        private static Object[][] fallbackIdValues = [
            [(long)PID_DICTIONARY, "PID_DICTIONARY" ],
            [(long)PID_CODEPAGE, "PID_CODEPAGE" ],
            [(long)PID_CATEGORY, "PID_CATEGORY" ],
            [(long)PID_PRESFORMAT, "PID_PRESFORMAT" ],
            [(long)PID_BYTECOUNT, "PID_BYTECOUNT" ],
            [(long)PID_LINECOUNT, "PID_LINECOUNT" ],
            [(long)PID_PARCOUNT, "PID_PARCOUNT" ],
            [(long)PID_SLIDECOUNT, "PID_SLIDECOUNT" ],
            [(long)PID_NOTECOUNT, "PID_NOTECOUNT" ],
            [(long)PID_HIDDENCOUNT, "PID_HIDDENCOUNT" ],
            [(long)PID_MMCLIPCOUNT, "PID_MMCLIPCOUNT" ],
            [(long)PID_SCALE, "PID_SCALE" ],
            [(long)PID_HEADINGPAIR, "PID_HEADINGPAIR" ],
            [(long)PID_DOCPARTS, "PID_DOCPARTS" ],
            [(long)PID_MANAGER, "PID_MANAGER" ],
            [(long)PID_COMPANY, "PID_COMPANY" ],
            [(long)PID_LINKSDIRTY, "PID_LINKSDIRTY" ],
            [(long)PID_CCHWITHSPACES, "PID_CCHWITHSPACES" ],
      // 0x12 Unused
      // 0x13 GKPIDDSI_SHAREDDOC - Must be False
      // 0x14 GKPIDDSI_LINKBASE - Must not be written
      // 0x15 GKPIDDSI_HLINKS - Must not be written
            [(long)PID_HYPERLINKSCHANGED, "PID_HYPERLINKSCHANGED" ],
            [(long)PID_VERSION, "PID_VERSION" ],
            [(long)PID_DIGSIG, "PID_DIGSIG" ],
      // 0x19 Unused
            [(long)PID_CONTENTTYPE, "PID_CONTENTTYPE" ],
            [(long)PID_CONTENTSTATUS, "PID_CONTENTSTATUS" ],
            [(long)PID_LANGUAGE, "PID_LANGUAGE" ],
            [(long)PID_DOCVERSION, "PID_DOCVERSION" ],
            [(long)PID_MAX, "PID_MAX" ],
            [(long)PID_LOCALE, "PID_LOCALE" ],
            [(long)PID_BEHAVIOUR, "PID_BEHAVIOUR" ],
        ];

        private readonly Dictionary<long, string> idMap;

        /// <summary>
        /// Creates a <see cref="PropertyIDMap"/> backed by another map.
        /// </summary>
        /// <param name="map">The instance to be created is backed by this map.</param>
        private PropertyIDMap(object[][] idValues)
        {
            Dictionary<long,String> m = new Dictionary<long,String>(idValues.Length);
            foreach (object[] idValue in idValues)
            {
                m.Add((long)idValue[0], (String)idValue[1]);
            }
            idMap = m;
        }

        /// <summary>
        /// Puts a ID string for an ID into the {@link
        /// PropertyIDMap}.
        /// </summary>
        /// <param name="id">The ID string.</param>
        /// <param name="idString">The id string.</param>
        /// <returns>As specified by the {@link java.util.Map} interface, this method
        /// returns the previous value associated with the specified id</returns>
        public Object Put(long id, String idString)
        {
            return idMap[id]=idString;
        }



        /// <summary>
        /// Gets the ID string for an ID from the {@link
        /// PropertyIDMap}.
        /// </summary>
        /// <param name="id">The ID.</param>
        /// <returns>The ID string associated with id</returns>
        public Object Get(long id)
        {
            if(idMap.TryGetValue(id, out var value))
                return value;
            return null;
        }

        public void Add(long key, string value)
        {
            idMap.Add(key, value);
        }

        public bool ContainsKey(long key)
        {
            return idMap.ContainsKey(key);
        }

        public bool Remove(long key)
        {
            return idMap.Remove(key);
        }

        public bool TryGetValue(long key, out string value)
        {
            return idMap.TryGetValue(key, out value);
        }

        public void Add(KeyValuePair<long, string> item)
        {
            idMap.Add(item.Key, item.Value);
        }

        public void Clear()
        {
            idMap.Clear();
        }

        public bool Contains(KeyValuePair<long, string> item)
        {
            throw new NotImplementedException();
        }

        public void CopyTo(KeyValuePair<long, string>[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public bool Remove(KeyValuePair<long, string> item)
        {
            return idMap.Remove(item.Key);
        }

        public IEnumerator<KeyValuePair<long, string>> GetEnumerator()
        {
            return idMap.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }



        /// <summary>
        /// Gets the Summary Information properties singleton
        /// </summary>
        /// <returns></returns>
        public static PropertyIDMap SummaryInformationProperties
        {
            get
            {
                lock(summaryInformationIdValues)
                {
                    if (summaryInformationProperties == null)
                    {
                        summaryInformationProperties = new PropertyIDMap(summaryInformationIdValues);
                    }
                }

                return summaryInformationProperties;
            }
        }



        /// <summary>
        /// Gets the Document Summary Information properties
        /// singleton.
        /// </summary>
        /// <returns>The Document Summary Information properties singleton.</returns>
        public static PropertyIDMap DocumentSummaryInformationProperties
        {
            get
            {
                lock(documentSummaryInformationIdValues)
                {
                    if (documentSummaryInformationProperties == null)
                    {
                        documentSummaryInformationProperties = new PropertyIDMap(documentSummaryInformationIdValues);
                    }
                }
                
                return documentSummaryInformationProperties;
            }
        }

        /// <summary>
        /// Returns a property map, which is only used as a fallback, i.e. if available, the correct map
        /// for <see cref="DocumentSummaryInformation"/> or <see cref="SummaryInformation"/> should be used.
        /// </summary>
        public static PropertyIDMap FallbackProperties
        {
            get
            {
                lock(fallbackIdValues)
                {
                    if (fallbackProperties == null)
                    {
                        fallbackProperties = new PropertyIDMap(fallbackIdValues);
                    }
                }
                
                return fallbackProperties;
            }
            
        }

        public ICollection<long> Keys => idMap.Keys;

        public ICollection<string> Values => idMap.Values;

        public int Count => idMap.Count;

        public bool IsReadOnly => false;

        public string this[long key] { get => idMap.TryGetValue(key, out string value) ? value : null; set => idMap[key] = value; }
    }
}