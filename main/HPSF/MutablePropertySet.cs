/* ====================================================================
   Licensed To the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file To You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed To in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
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

using System.Collections.Generic;

namespace NPOI.HPSF
{
    using System;
    using System.IO;
    using System.Collections;
    using NPOI.Util;
    using NPOI.POIFS.FileSystem;


    /// <summary>
    /// Adds writing support To the {@link PropertySet} class.
    /// Please be aware that this class' functionality will be merged into the
    /// {@link PropertySet} class at a later time, so the API will Change.
    /// @author Rainer Klute 
    /// <a href="mailto:klute@rainer-klute.de">&lt;klute@rainer-klute.de&gt;</a>
    /// @since 2003-02-19
    /// </summary>
    [Serializable]
    public class MutablePropertySet : PropertySet
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="MutablePropertySet"/> class.
        /// Its primary task is To initialize the immutable field with their proper
        /// values. It also Sets fields that might Change To reasonable defaults.
        /// </summary>
        public MutablePropertySet()
        {
            /* Initialize the "byteOrder" field. */
            byteOrder = LittleEndian.GetUShort(BYTE_ORDER_ASSERTION);

            /* Initialize the "format" field. */
            format = LittleEndian.GetUShort(FORMAT_ASSERTION);

            /* Initialize "osVersion" field as if the property has been Created on
             * a Win32 platform, whether this is the case or not. */
            osVersion = (OS_WIN32 << 16) | 0x0A04;

            /* Initailize the "classID" field. */
            classID = new ClassID();

            /* Initialize the sections. Since property Set must have at least
             * one section it is Added right here. */
            sections = new List<Section>();
            sections.Add(new MutableSection());
        }



        /// <summary>
        /// Initializes a new instance of the <see cref="MutablePropertySet"/> class.
        /// All nested elements, i.e.<c>Section</c>s and <c>Property</c> instances, will be their
        /// mutable counterparts in the new <c>MutablePropertySet</c>.
        /// </summary>
        /// <param name="ps">The property Set To copy</param>
        public MutablePropertySet(PropertySet ps)
        {
            byteOrder = ps.ByteOrder;
            format = ps.Format;
            osVersion = ps.OSVersion;
            ClassID=ps.ClassID;
            ClearSections();
            if (sections == null)
                sections = new List<Section>();
            foreach (Section section in ps.Sections)
            {
                MutableSection s = new MutableSection(section);
                AddSection(s);
            }
        }



        /**
         * The Length of the property Set stream header.
         */
        private int OFFSET_HEADER =
            BYTE_ORDER_ASSERTION.Length + /* Byte order    */
            FORMAT_ASSERTION.Length +     /* Format        */
            LittleEndianConsts.INT_SIZE + /* OS version    */
            ClassID.LENGTH +              /* Class ID      */
            LittleEndianConsts.INT_SIZE;  /* Section count */



        /// <summary>
        /// Gets or sets the "byteOrder" property.
        /// </summary>
        /// <value>the byteOrder value To Set</value>
        public override int ByteOrder
        {
            get { return this.byteOrder; }
            set { this.byteOrder = value; }
        }



        /// <summary>
        /// Gets or sets the "format" property.
        /// </summary>
        /// <value>the format value To Set</value>
        public override int Format
        {
            set { this.format = value; }
            get { return this.format; }
        }



        /// <summary>
        /// Gets or sets the "osVersion" property
        /// </summary>
        /// <value>the osVersion value To Set.</value>
        public override int OSVersion
        {
            set { this.osVersion = value; }
            get { return this.osVersion; }
        }



        /// <summary>
        /// Gets or sets the property Set stream's low-level "class ID"
        /// </summary>
        /// <value>The property Set stream's low-level "class ID" field.</value>
        public override ClassID ClassID
        {
            set { this.classID = value; }
            get { return this.classID; }
        }



        /// <summary>
        /// Removes all sections from this property Set.
        /// </summary>
        public virtual void ClearSections()
        {
            sections=null;
        }



        /// <summary>
        /// Adds a section To this property Set.
        /// </summary>
        /// <param name="section">section The {@link Section} To Add. It will be Appended
        /// after any sections that are alReady present in the property Set
        /// and thus become the last section.</param>
        public virtual void AddSection(Section section)
        {
            if (sections == null)
                sections = new List<Section>();
            sections.Add(section);
        }



        /// <summary>
        /// Writes the property Set To an output stream.
        /// </summary>
        /// <param name="out1">the output stream To Write the section To</param>
        public virtual void Write(Stream out1)
        {
            /* Write the number of sections in this property Set stream. */
            int nrSections = sections.Count;
            int length = 0;

            /* Write the property Set's header. */
            length += TypeWriter.WriteToStream(out1, (short)ByteOrder);
            length += TypeWriter.WriteToStream(out1, (short)Format);
            length += TypeWriter.WriteToStream(out1, OSVersion);
            length += TypeWriter.WriteToStream(out1, ClassID);
            length += TypeWriter.WriteToStream(out1, nrSections);
            int offset = OFFSET_HEADER;

            /* Write the section list, i.e. the references To the sections. Each
             * entry in the section list consist of the section's class ID and the
             * section's offset relative To the beginning of the stream. */
            offset += nrSections * (ClassID.Length + LittleEndianConsts.INT_SIZE);
            int sectionsBegin = offset;
            for (IEnumerator i = sections.GetEnumerator(); i.MoveNext(); )
            {
                MutableSection s = (MutableSection)i.Current;
                ClassID formatID = s.FormatID;
                if (formatID == null)
                    throw new NoFormatIDException();
                length += TypeWriter.WriteToStream(out1, s.FormatID);
                length += TypeWriter.WriteUIntToStream(out1, (uint)offset);

                offset += s.Size;

            }

            /* Write the sections themselves. */
            offset = sectionsBegin;
            for (IEnumerator i = sections.GetEnumerator(); i.MoveNext(); )
            {
                MutableSection s = (MutableSection)i.Current;
                offset += s.Write(out1);
            }

            /* Indicate that we're done */
            out1.Close();
        }

        /// <summary>
        /// Returns the contents of this property set stream as an input stream.
        /// The latter can be used for example to write the property set into a POIFS
        /// document. The input stream represents a snapshot of the property set.
        /// If the latter is modified while the input stream is still being
        /// read, the modifications will not be reflected in the input stream but in
        /// the {@link MutablePropertySet} only.
        /// </summary>
        /// <returns>the contents of this property set stream</returns>
        public virtual Stream ToInputStream()
        {
            using (MemoryStream psStream = new MemoryStream())
            {
                try
                {
                    Write(psStream);
                    psStream.Flush();
                }
                finally
                {
                    psStream.Close();
                }
                byte[] streamData = psStream.ToArray();
                return new MemoryStream(streamData);
            }
        }

        /// <summary>
        /// Writes a property Set To a document in a POI filesystem directory
        /// </summary>
        /// <param name="dir">The directory in the POI filesystem To Write the document To.</param>
        /// <param name="name">The document's name. If there is alReady a document with the
        /// same name in the directory the latter will be overwritten.</param>
        public virtual void Write(DirectoryEntry dir, String name)
        {
            /* If there is alReady an entry with the same name, Remove it. */
            try
            {
                Entry e = dir.GetEntry(name);
                e.Delete();
            }
            catch (FileNotFoundException)
            {
                /* Entry not found, no need To Remove it. */
            }
            /* Create the new entry. */
            dir.CreateDocument(name, ToInputStream());
        }

    }
}