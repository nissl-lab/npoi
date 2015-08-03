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

namespace NPOI.HPSF
{
    using System.IO;
    using NPOI.HPSF.Wellknown;
    using System;
using NPOI.POIFS.FileSystem;

    /// <summary>
    /// Factory class To Create instances of {@link SummaryInformation},
    /// {@link DocumentSummaryInformation} and {@link PropertySet}.
    /// @author Rainer Klute 
    /// <a href="mailto:klute@rainer-klute.de">&lt;klute@rainer-klute.de&gt;</a>
    /// @since 2002-02-09
    /// </summary>
    public class PropertySetFactory
    {
        /**
         * <p>Creates the most specific {@link PropertySet} from an entry
         *  in the specified POIFS Directory. This is preferrably a {@link
         * DocumentSummaryInformation} or a {@link SummaryInformation}. If
         * the specified entry does not contain a property Set stream, an 
         * exception is thrown. If no entry is found with the given name,
         * an exception is thrown.</p>
         *
         * @param dir The directory to find the PropertySet in
         * @param name The name of the entry Containing the PropertySet
         * @return The Created {@link PropertySet}.
         * @if there is no entry with that name
         * @if the stream does not
         * contain a property Set.
         * @if some I/O problem occurs.
         * @exception EncoderFallbackException if the specified codepage is not
         * supported.
         */
        public static PropertySet Create(DirectoryEntry dir, String name)
        {
            Stream inp = null;
            try
            {
                DocumentEntry entry = (DocumentEntry)dir.GetEntry(name);
                inp = new DocumentInputStream(entry);
                try
                {
                    return Create(inp);
                }
                catch (MarkUnsupportedException) { return null; }
            }
            finally
            {
                if (inp != null) inp.Close();
            }
        }

        /// <summary>
        /// Creates the most specific {@link PropertySet} from an {@link
        /// InputStream}. This is preferrably a {@link
        /// DocumentSummaryInformation} or a {@link SummaryInformation}. If
        /// the specified {@link InputStream} does not contain a property
        /// Set stream, an exception is thrown and the {@link InputStream}
        /// is repositioned at its beginning.
        /// </summary>
        /// <param name="stream">Contains the property set stream's data.</param>
        /// <returns>The Created {@link PropertySet}.</returns>
        public static PropertySet Create(Stream stream)
        {
            PropertySet ps = new PropertySet(stream);
            try
            {
                if (ps.IsSummaryInformation)
                    return new SummaryInformation(ps);
                else if (ps.IsDocumentSummaryInformation)
                    return new DocumentSummaryInformation(ps);
                else
                    return ps;
            }
            catch (UnexpectedPropertySetTypeException ex)
            {
                /* This exception will never be throws because we alReady checked
                 * explicitly for this case above. */
                throw new InvalidOperationException(ex.Message, ex);
            }
        }



        /// <summary>
        /// Creates a new summary information
        /// </summary>
        /// <returns>the new summary information.</returns>
        public static SummaryInformation CreateSummaryInformation()
        {
            MutablePropertySet ps = new MutablePropertySet();
            MutableSection s = (MutableSection)ps.FirstSection;
            s.SetFormatID(SectionIDMap.SUMMARY_INFORMATION_ID);
            try
            {
                return new SummaryInformation(ps);
            }
            catch (UnexpectedPropertySetTypeException ex)
            {
                /* This should never happen. */
                throw new HPSFRuntimeException(ex);
            }
        }



        /// <summary>
        /// Creates a new document summary information.
        /// </summary>
        /// <returns>the new document summary information.</returns>
        public static DocumentSummaryInformation CreateDocumentSummaryInformation()
        {
            MutablePropertySet ps = new MutablePropertySet();
            MutableSection s = (MutableSection)ps.FirstSection;
            s.SetFormatID(SectionIDMap.DOCUMENT_SUMMARY_INFORMATION_ID1);
            try
            {
                return new DocumentSummaryInformation(ps);
            }
            catch (UnexpectedPropertySetTypeException ex)
            {
                /* This should never happen. */
                throw new HPSFRuntimeException(ex);
            }
        }

    }
}