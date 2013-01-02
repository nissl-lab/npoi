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
    public class PropertyIDMap : Hashtable
    {

        /*
         * The following definitions are for property IDs in the first
         * (and only) section of the Summary Information property Set.
         */

        /** ID of the property that denotes the document's title */
        public const int PID_TITLE = 2;

        /** ID of the property that denotes the document's subject */
        public const int PID_SUBJECT = 3;

        /** ID of the property that denotes the document's author */
        public const int PID_AUTHOR = 4;

        /** ID of the property that denotes the document's keywords */
        public const int PID_KEYWORDS = 5;

        /** ID of the property that denotes the document's comments */
        public const int PID_COMMENTS = 6;

        /** ID of the property that denotes the document's template */
        public const int PID_TEMPLATE = 7;

        /** ID of the property that denotes the document's last author */
        public const int PID_LASTAUTHOR = 8;

        /** ID of the property that denotes the document's revision number */
        public const int PID_REVNUMBER = 9;

        /** ID of the property that denotes the document's edit time */
        public const int PID_EDITTIME = 10;

        /** ID of the property that denotes the DateTime and time the document was
         * last printed */
        public const int PID_LASTPRINTED = 11;

        /** ID of the property that denotes the DateTime and time the document was
         * Created. */
        public const int PID_Create_DTM = 12;

        /** ID of the property that denotes the DateTime and time the document was
         * saved */
        public const int PID_LASTSAVE_DTM = 13;

        /** ID of the property that denotes the number of pages in the
         * document */
        public const int PID_PAGECOUNT = 14;

        /** ID of the property that denotes the number of words in the
         * document */
        public const int PID_WORDCOUNT = 15;

        /** ID of the property that denotes the number of characters in the
         * document */
        public const int PID_CHARCOUNT = 16;

        /** ID of the property that denotes the document's thumbnail */
        public const int PID_THUMBNAIL = 17;

        /** ID of the property that denotes the application that Created the
         * document */
        public const int PID_APPNAME = 18;

        /** ID of the property that denotes whether Read/Write access To the
         * document is allowed or whether is should be opened as Read-only. It can
         * have the following values:
         * 
         * <table>
         *  <tbody>
         *   <tr>
         *    <th>Value</th>
         *    <th>Description</th>
         *   </tr>
         *   <tr>
         *    <th>0</th>
         *    <th>No restriction</th>
         *   </tr>
         *   <tr>
         *    <th>2</th>
         *    <th>Read-only recommended</th>
         *   </tr>
         *   <tr>
         *    <th>4</th>
         *    <th>Read-only enforced</th>
         *   </tr>
         *  </tbody>
         * </table>
         */
        public const int PID_SECURITY = 19;



        /*
         * The following definitions are for property IDs in the first
         * section of the Document Summary Information property Set.
         */

        /** 
         * The entry is a dictionary.
         */
        public const int PID_DICTIONARY = 0;

        /**
         * The entry denotes a code page.
         */
        public const int PID_CODEPAGE = 1;

        /** 
         * The entry is a string denoting the category the file belongs
         * To, e.g. review, memo, etc. This is useful To Find documents of
         * same type.
         */
        public const int PID_CATEGORY = 2;

        /** 
         * TarGet format for power point presentation, e.g. 35mm,
         * printer, video etc.
         */
        public const int PID_PRESFORMAT = 3;

        /** 
         * Number of bytes.
         */
        public const int PID_BYTECOUNT = 4;

        /** 
         * Number of lines.
         */
        public const int PID_LINECOUNT = 5;

        /** 
         * Number of paragraphs.
         */
        public const int PID_PARCOUNT = 6;

        /** 
         * Number of slides in a power point presentation.
         */
        public const int PID_SLIDECOUNT = 7;

        /** 
         * Number of slides with notes.
         */
        public const int PID_NOTECOUNT = 8;

        /** 
         * Number of hidden slides.
         */
        public const int PID_HIDDENCOUNT = 9;

        /** 
         * Number of multimedia clips, e.g. sound or video.
         */
        public const int PID_MMCLIPCOUNT = 10;

        /** 
         * This entry is Set To -1 when scaling of the thumbnail Is
         * desired. Otherwise the thumbnail should be cropped.
         */
        public const int PID_SCALE = 11;

        /** 
         * This entry denotes an internally used property. It is a
         * vector of variants consisting of pairs of a string (VT_LPSTR)
         * and a number (VT_I4). The string is a heading name, and the
         * number tells how many document parts are under that
         * heading.
         */
        public const int PID_HEADINGPAIR = 12;

        /** 
         * This entry Contains the names of document parts (word: names
         * of the documents in the master document, excel: sheet names,
         * power point: slide titles, binder: document names).
         */
        public const int PID_DOCPARTS = 13;

        /** 
         * This entry Contains the name of the project manager.
         */
        public const int PID_MANAGER = 14;

        /** 
         * This entry Contains the company name.
         */
        public const int PID_COMPANY = 15;

        /** 
         * If this entry is -1 the links are dirty and should be
         * re-evaluated.
         */
        public const int PID_LINKSDIRTY = 16;

        /**
         * The highest well-known property ID. Applications are free To use higher values for custom purposes.
         */
        public const int PID_MAX = PID_LINKSDIRTY;



        /**
         * Contains the summary information property ID values and
         * associated strings. See the overall HPSF documentation for
         * details!
         */
        private static PropertyIDMap summaryInformationProperties;

        /**
         * Contains the summary information property ID values and
         * associated strings. See the overall HPSF documentation for
         * details!
         */
        private static PropertyIDMap documentSummaryInformationProperties;



        /// <summary>
        /// Initializes a new instance of the <see cref="PropertyIDMap"/> class.
        /// </summary>
        /// <param name="initialCapacity">initialCapacity The initial capacity as defined for
        /// {@link HashMap}</param>
        /// <param name="loadFactor">The load factor as defined for {@link HashMap}</param>
        public PropertyIDMap(int initialCapacity, float loadFactor):base(initialCapacity, loadFactor)
        {
            
        }



        /// <summary>
        /// Initializes a new instance of the <see cref="PropertyIDMap"/> class.
        /// </summary>
        /// <param name="map">The instance To be Created is backed by this map.</param>
        public PropertyIDMap(IDictionary map):base(map)
        {
            
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
            return this[id]=idString;
        }



        /// <summary>
        /// Gets the ID string for an ID from the {@link
        /// PropertyIDMap}.
        /// </summary>
        /// <param name="id">The ID.</param>
        /// <returns>The ID string associated with id</returns>
        public Object Get(long id)
        {
            return this[id];
        }



        /// <summary>
        /// Gets the Summary Information properties singleton
        /// </summary>
        /// <returns></returns>
        public static PropertyIDMap SummaryInformationProperties
        {
            get
            {
                if (summaryInformationProperties == null)
                {
                    PropertyIDMap m = new PropertyIDMap(18, (float)1.0);
                    m.Put(PID_TITLE, "PID_TITLE");
                    m.Put(PID_SUBJECT, "PID_SUBJECT");
                    m.Put(PID_AUTHOR, "PID_AUTHOR");
                    m.Put(PID_KEYWORDS, "PID_KEYWORDS");
                    m.Put(PID_COMMENTS, "PID_COMMENTS");
                    m.Put(PID_TEMPLATE, "PID_TEMPLATE");
                    m.Put(PID_LASTAUTHOR, "PID_LASTAUTHOR");
                    m.Put(PID_REVNUMBER, "PID_REVNUMBER");
                    m.Put(PID_EDITTIME, "PID_EDITTIME");
                    m.Put(PID_LASTPRINTED, "PID_LASTPRINTED");
                    m.Put(PID_Create_DTM, "PID_Create_DTM");
                    m.Put(PID_LASTSAVE_DTM, "PID_LASTSAVE_DTM");
                    m.Put(PID_PAGECOUNT, "PID_PAGECOUNT");
                    m.Put(PID_WORDCOUNT, "PID_WORDCOUNT");
                    m.Put(PID_CHARCOUNT, "PID_CHARCOUNT");
                    m.Put(PID_THUMBNAIL, "PID_THUMBNAIL");
                    m.Put(PID_APPNAME, "PID_APPNAME");
                    m.Put(PID_SECURITY, "PID_SECURITY");
                    summaryInformationProperties = m;
                    //new PropertyIDMap(m);
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
                if (documentSummaryInformationProperties == null)
                {
                    PropertyIDMap m = new PropertyIDMap(17, (float)1.0);
                    m.Put(PID_DICTIONARY, "PID_DICTIONARY");
                    m.Put(PID_CODEPAGE, "PID_CODEPAGE");
                    m.Put(PID_CATEGORY, "PID_CATEGORY");
                    m.Put(PID_PRESFORMAT, "PID_PRESFORMAT");
                    m.Put(PID_BYTECOUNT, "PID_BYTECOUNT");
                    m.Put(PID_LINECOUNT, "PID_LINECOUNT");
                    m.Put(PID_PARCOUNT, "PID_PARCOUNT");
                    m.Put(PID_SLIDECOUNT, "PID_SLIDECOUNT");
                    m.Put(PID_NOTECOUNT, "PID_NOTECOUNT");
                    m.Put(PID_HIDDENCOUNT, "PID_HIDDENCOUNT");
                    m.Put(PID_MMCLIPCOUNT, "PID_MMCLIPCOUNT");
                    m.Put(PID_SCALE, "PID_SCALE");
                    m.Put(PID_HEADINGPAIR, "PID_HEADINGPAIR");
                    m.Put(PID_DOCPARTS, "PID_DOCPARTS");
                    m.Put(PID_MANAGER, "PID_MANAGER");
                    m.Put(PID_COMPANY, "PID_COMPANY");
                    m.Put(PID_LINKSDIRTY, "PID_LINKSDIRTY");
                    documentSummaryInformationProperties = m;
                    //new PropertyIDMap(m);
                }
                return documentSummaryInformationProperties;
            }
        }
    }
}