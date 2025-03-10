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

namespace NPOI.HPSF

{


    using NPOI;
    using NPOI.HPSF.Wellknown;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.IO;
    using System.Text;

    /// <summary>
    /// <para>
    /// Represents a property Set in the Horrible Property Set Format
    /// (HPSF). These are usually metadata of a Microsoft Office
    /// </para>
    /// <para>
    /// document.
    /// </para>
    /// <para>
    /// An application that wants to access these metadata should create
    /// an instance of this class or one of its subclasses by calling the
    /// factory method {@link PropertySetFactory#create} and then retrieve
    /// </para>
    /// <para>
    /// the information its needs by calling appropriate methods.
    /// </para>
    /// <para>
    /// {@link PropertySetFactory#create} does its work by calling one
    /// of the constructors {@link PropertySet#PropertySet(InputStream)} or
    /// {@link PropertySet#PropertySet(byte[])}. If the constructor's
    /// argument is not in the Horrible Property Set Format, i.e. not a
    /// property Set stream, or if any other error occurs, an appropriate
    /// </para>
    /// <para>
    /// exception is thrown.
    /// </para>
    /// <para>
    /// A <see cref="PropertySet"/> has a list of <see cref="Section"/>s, and each
    /// <see cref="Section"/> has a {@link Property} array. Use {@link
    /// #getSections} to retrieve the <see cref="Section"/>s, then call {@link
    /// Section#getProperties} for each <see cref="Section"/> to Get hold of the
    /// </para>
    /// <para>
    /// {@link Property} arrays.
    /// </para>
    /// <para>
    /// Since the vast majority of <see cref="PropertySet"/>s contains only a single
    /// <see cref="Section"/>, the convenience method <see cref="getProperties"/> returns
    /// the properties of a <see cref="PropertySet"/>'s <see cref="Section"/> (throwing a
    /// {@link NoSingleSectionException} if the <see cref="PropertySet"/> contains
    /// more (or less) than exactly one <see cref="Section"/>).
    /// </para>
    /// </summary>
    public class PropertySet
    {
        /// <summary>
        /// If the OS version field holds this value the property Set stream was
        /// created on a 16-bit Windows system.
        /// </summary>
        public static int OS_WIN16 = 0x0000;

        /// <summary>
        /// If the OS version field holds this value the property Set stream was
        /// created on a Macintosh system.
        /// </summary>
        public static int OS_MACINTOSH = 0x0001;

        /// <summary>
        /// If the OS version field holds this value the property Set stream was
        /// created on a 32-bit Windows system.
        /// </summary>
        public static int OS_WIN32 = 0x0002;

        /// <summary>
        /// The "byteOrder" field must equal this value.
        /// </summary>
        private static int BYTE_ORDER_ASSERTION = 0xFFFE;

        /// <summary>
        /// The "format" field must equal this value.
        /// </summary>
        private static int FORMAT_ASSERTION = 0x0000;

        /// <summary>
        /// The length of the property Set stream header.
        /// </summary>
        private static int OFFSET_HEADER =
            LittleEndianConsts.SHORT_SIZE + /* Byte order    */
            LittleEndianConsts.SHORT_SIZE + /* Format        */
            LittleEndianConsts.INT_SIZE +   /* OS version    */
            ClassID.LENGTH +                /* Class ID      */
            LittleEndianConsts.INT_SIZE;    /* Section count */


        /// <summary>
        /// Specifies this <see cref="PropertySet"/>'s byte order. See the
        /// HPFS documentation for details!
        /// </summary>
        private int byteOrder;

        /// <summary>
        /// Specifies this <see cref="PropertySet"/>'s format. See the HPFS
        /// documentation for details!
        /// </summary>
        private int format;

        /// <summary>
        /// Specifies the version of the operating system that created this
        /// <see cref="PropertySet"/>. See the HPFS documentation for details!
        /// </summary>
        private int osVersion;

        /// <summary>
        /// Specifies this <see cref="PropertySet"/>'s "classID" field. See
        /// the HPFS documentation for details!
        /// </summary>
        private ClassID classID;

        /// <summary>
        /// The sections in this <see cref="PropertySet"/>.
        /// </summary>
        private List<Section> sections = new List<Section>();


        /// <summary>
        /// Constructs a <c>PropertySet</c> instance. Its
        /// primary task is to initialize the field with their proper values.
        /// It also Sets fields that might change to reasonable defaults.
        /// </summary>
        public PropertySet()
        {
            /* Initialize the "byteOrder" field. */
            byteOrder = BYTE_ORDER_ASSERTION;

            /* Initialize the "format" field. */
            format = FORMAT_ASSERTION;

            /* Initialize "osVersion" field as if the property has been created on
             * a Win32 platform, whether this is the case or not. */
            osVersion = (OS_WIN32 << 16) | 0x0A04;

            /* Initialize the "classID" field. */
            classID = new ClassID();

            /* Initialize the sections. Since property Set must have at least
             * one section it is added right here. */
            AddSection(new MutableSection());
        }



        /// <summary>
        /// <para>
        /// Creates a <see cref="PropertySet"/> instance from an <see cref="InputStream"/>
        /// in the Horrible Property Set Format.
        /// </para>
        /// <para>
        /// The constructor reads the first few bytes from the stream
        /// and determines whether it is really a property Set stream. If
        /// it is, it parses the rest of the stream. If it is not, it
        /// resets the stream to its beginning in order to let other
        /// components mess around with the data and throws an
        /// exception.
        /// </para>
        /// </summary>
        /// <param name="stream">Holds the data making out the property Set
        /// stream.
        /// </param>
        /// <exception cref="MarkUnsupportedException">
        /// if the stream does not support the <see cref="InputStream.MarkSupported"/> method.
        /// </exception>
        /// <exception cref="IOException">
        /// if the <see cref="InputStream"/> cannot be accessed as needed.
        /// </exception>
        /// <exception cref="NoPropertySetStreamException">
        /// if the input stream does not contain a property Set.
        /// </exception>
        /// <exception cref="UnsupportedEncodingException">
        /// if a character encoding is not supported.
        /// </exception>
        public PropertySet(InputStream stream)
        {
            if(!IsPropertySetStream(stream))
            {
                throw new NoPropertySetStreamException();
            }

            byte[] buffer = IOUtils.ToByteArray(stream);
            Init(buffer, 0, buffer.Length);
        }



        /// <summary>
        /// Creates a <see cref="PropertySet"/> instance from a byte array that
        /// represents a stream in the Horrible Property Set Format.
        /// </summary>
        /// <param name="stream">The byte array holding the stream data.</param>
        /// <param name="offset">The offset in <c>stream</c> where the stream
        /// data begin. If the stream data begin with the first byte in the
        /// array, the <c>offset</c> is 0.
        /// </param>
        /// <param name="length">The length of the stream data.</param>
        /// <exception cref="NoPropertySetStreamException">if the byte array is not a
        /// property Set stream.
        /// </exception>
        public PropertySet(byte[] stream, int offset, int length)
        {

            if(!IsPropertySetStream(stream, offset, length))
            {
                throw new NoPropertySetStreamException();
            }
            Init(stream, offset, length);
        }

        /// <summary>
        /// Creates a <see cref="PropertySet"/> instance from a byte array
        /// that represents a stream in the Horrible Property Set Format.
        /// </summary>
        /// <param name="stream">The byte array holding the stream data. The
        /// complete byte array contents is the stream data.
        /// </param>
        /// <exception cref="NoPropertySetStreamException">if the byte array is not a
        /// property Set stream.
        /// </exception>
        public PropertySet(byte[] stream)
            : this(stream, 0, stream.Length)
        {
        }

        /// <summary>
        /// Constructs a <c>PropertySet</c> by doing a deep copy of
        /// an existing <c>PropertySet</c>. All nested elements, i.e.
        /// <c>Section</c>s and <c>Property</c> instances, will be their
        /// counterparts in the new <c>PropertySet</c>.
        /// </summary>
        /// <param name="ps">The property Set to copy</param>
        public PropertySet(PropertySet ps)
        {
            ByteOrder = ps.ByteOrder;
            Format = ps.Format;
            OSVersion = ps.OSVersion;
            ClassID = ps.ClassID;
            foreach(Section section in ps.Sections)
            {
                sections.Add(new MutableSection(section));
            }
        }


        /// <summary>
        /// get or set the property Set stream's low-level "byte order" field.
        /// </summary>
        public int ByteOrder
        {
            get { return byteOrder; }
            set { byteOrder = value; }
        }

        /// <summary>
        /// get or set the property Set stream's low-level "format" field. It is always <c>0x0000</c>.
        /// </summary>
        public int Format
        {
            get { return format; }
            set { format = value; }
        }


        /// <summary>
        /// get or set the property Set stream's low-level "OS version" field.
        /// </summary>
        public int OSVersion
        {
            get { return osVersion; }
            set { osVersion = value; }
        }

        /// <summary>
        /// get or set the property Set stream's low-level "class ID" field.
        /// </summary>
        public ClassID ClassID
        {
            get { return classID; }
            set { classID = value; }
        }

        /// <summary>
        /// </summary>
        /// <return>number of <see cref="Section"/>s in the property Set.</return>
        public int SectionCount
        {
            get
            {
                return sections.Count;
            }

        }

        /// <summary>
        /// </summary>
        /// <return>unmodifiable list of <see cref="Section"/>s in the property Set.</return>
        public List<Section> Sections
        {
            get
            {
                return sections;
                //return sections.AsReadOnly();
            }
        }



        /// <summary>
        /// Adds a section to this property Set.
        /// </summary>
        /// <param name="section">The <see cref="Section"/> to add. It will be appended
        /// After any sections that are already present in the property Set
        /// and thus become the last section.
        /// </param>
        public void AddSection(Section section)
        {
            sections.Add(section);
        }

        /// <summary>
        /// Removes all sections from this property Set.
        /// </summary>
        public void ClearSections()
        {
            sections.Clear();
        }

        /// <summary>
        /// The id to name mapping of the properties in this Set.
        /// </summary>
        /// <return>id to name mapping of the properties in this Set or <c>null</c> if not applicable</return>
        public virtual PropertyIDMap PropertySetIDMap
        {
            get { return null; }
        }


        /// <summary>
        /// Checks whether an {@link InputStream} is in the Horrible
        /// Property Set Format.
        /// </summary>
        /// <param name="stream">The {@link InputStream} to check. In order to
        /// perform the check, the method reads the first bytes from the
        /// stream. After reading, the stream is reset to the position it
        /// had before reading. The {@link InputStream} must support the
        /// {@link InputStream#mark} method.
        /// </param>
        /// <return>true} if the stream is a property Set
        /// stream, else <c>false</c>.
        /// </return>
        /// <exception cref="MarkUnsupportedException">if the {@link InputStream}
        /// does not support the {@link InputStream#mark} method.
        /// </exception>
        /// <exception cref="IOException">if an I/O error occurs</exception>
        public static bool IsPropertySetStream(InputStream stream)
        {

            /*
             * Read at most this many bytes.
             */
            int BUFFER_SIZE = 50;

            /*
             * Read a couple of bytes from the stream.
             */
            try
            {
                byte[] buffer = IOUtils.PeekFirstNBytes(stream, BUFFER_SIZE);
                bool isPropertySetStream = IsPropertySetStream(buffer, 0, buffer.Length);
                return isPropertySetStream;
            }
            catch(EmptyFileException e)
            {
                return false;
            }
        }



        /// <summary>
        /// Checks whether a byte array is in the Horrible Property Set Format.
        /// </summary>
        /// <param name="src">The byte array to check.</param>
        /// <param name="offset">The offset in the byte array.</param>
        /// <param name="length">The significant number of bytes in the byte
        /// array. Only this number of bytes will be checked.
        /// </param>
        /// <return>true} if the byte array is a property Set
        /// stream, <c>false</c> if not.
        /// </return>
        public static bool IsPropertySetStream(byte[] src, int offset, int length)
        {
            /* FIXME (3): Ensure that at most "length" bytes are read. */

            /*
             * Read the header fields of the stream. They must always be
             * there.
             */
            int o = offset;
            int byteOrder = LittleEndian.GetUShort(src, o);
            o += LittleEndianConsts.SHORT_SIZE;
            if(byteOrder != BYTE_ORDER_ASSERTION)
            {
                return false;
            }
            int format = LittleEndian.GetUShort(src, o);
            o += LittleEndianConsts.SHORT_SIZE;
            if(format != FORMAT_ASSERTION)
            {
                return false;
            }
            // final long osVersion = LittleEndian.GetUInt(src, offset);
            o += LittleEndianConsts.INT_SIZE;
            // final ClassID classID = new ClassID(src, offset);
            o += ClassID.LENGTH;
            long sectionCount = LittleEndian.GetUInt(src, o);
            return (sectionCount >= 0);
        }



        /// <summary>
        /// Initializes this <see cref="PropertySet"/> instance from a byte
        /// array. The method assumes that it has been checked already that
        /// the byte array indeed represents a property Set stream. It does
        /// no more checks on its own.
        /// </summary>
        /// <param name="src">Byte array containing the property Set stream</param>
        /// <param name="offset">The property Set stream starts at this offset
        /// from the beginning of <c>src</c>
        /// </param>
        /// <param name="length">Length of the property Set stream.</param>
        /// <exception cref="UnsupportedEncodingException">if HPSF does not (yet) support the
        /// property Set's character encoding.
        /// </exception>
        private void Init(byte[] src, int offset, int length)
        {

            /* FIXME (3): Ensure that at most "length" bytes are read. */

            /*
             * Read the stream's header fields.
             */
            int o = offset;
            byteOrder = LittleEndian.GetUShort(src, o);
            o += LittleEndianConsts.SHORT_SIZE;
            format = LittleEndian.GetUShort(src, o);
            o += LittleEndianConsts.SHORT_SIZE;
            osVersion = (int) LittleEndian.GetUInt(src, o);
            o += LittleEndianConsts.INT_SIZE;
            classID = new ClassID(src, o);
            o += ClassID.LENGTH;
            int sectionCount = LittleEndian.GetInt(src, o);
            o += LittleEndianConsts.INT_SIZE;
            if(sectionCount < 0)
            {
                throw new HPSFRuntimeException("Section count " + sectionCount + " is negative.");
            }

            /*
             * Read the sections, which are following the header. They
             * start with an array of section descriptions. Each one
             * consists of a format ID telling what the section contains
             * and an offset telling how many bytes from the start of the
             * stream the section begins.
             * 
             * Most property Sets have only one section. The Document
             * Summary Information stream has 2. Everything else is a rare
             * exception and is no longer fostered by Microsoft.
             */

            /*
             * Loop over the section descriptor array. Each descriptor
             * consists of a ClassID and a DWord, and we have to increment
             * "offset" accordingly.
             */
            for(int i = 0; i < sectionCount; i++)
            {
                Section s = new MutableSection(src, o);
                o += ClassID.LENGTH + LittleEndianConsts.INT_SIZE;
                sections.Add(s);
            }
        }

        /// <summary>
        /// Writes the property Set to an output stream.
        /// </summary>
        /// <param name="out">the output stream to write the section to</param>
        /// <exception cref="IOException">if an error when writing to the output stream
        /// occurs
        /// </exception>
        /// <exception cref="WritingNotSupportedException">if HPSF does not yet support
        /// writing a property's variant type.
        /// </exception>
        public void Write(Stream out1)
        {

            /* Write the number of sections in this property Set stream. */
            int nrSections = SectionCount;

            /* Write the property Set's header. */
            TypeWriter.WriteToStream(out1, (short) ByteOrder);
            TypeWriter.WriteToStream(out1, (short) Format);
            TypeWriter.WriteToStream(out1, OSVersion);
            TypeWriter.WriteToStream(out1, ClassID);
            TypeWriter.WriteToStream(out1, nrSections);
            int offset = OFFSET_HEADER;

            /* Write the section list, i.e. the references to the sections. Each
             * entry in the section list consist of the section's class ID and the
             * section's offset relative to the beginning of the stream. */
            offset += nrSections * (ClassID.LENGTH + LittleEndianConsts.INT_SIZE);
            int sectionsBegin = offset;
            foreach(Section section in Sections)
            {
                ClassID formatID = section.FormatID;
                if(formatID == null)
                {
                    throw new NoFormatIDException();
                }
                TypeWriter.WriteToStream(out1, section.FormatID);
                TypeWriter.WriteUIntToStream(out1, (uint) offset);
                try
                {
                    offset += section.Size;
                }
                catch(HPSFRuntimeException ex)
                {
                    Exception cause = ex.InnerException;
                    if(cause is ArgumentException)
                    {
                        throw new IllegalPropertySetDataException(cause);
                    }
                    throw;
                }
            }

            /* Write the sections themselves. */
            offset = sectionsBegin;
            foreach(Section section in Sections)
            {
                offset += section.Write(out1);
            }

            /* Indicate that we're done */
            out1.Close();
        }

        /// <summary>
        /// Writes a property Set to a document in a POI filesystem directory.
        /// </summary>
        /// <param name="dir">The directory in the POI filesystem to write the document to.</param>
        /// <param name="name">The document's name. If there is already a document with the
        /// same name in the directory the latter will be overwritten.
        /// </param>
        /// 
        /// <exception cref="WritingNotSupportedException">if the filesystem doesn't support writing</exception>
        /// <exception cref="IOException">if the old entry can't be deleted or the new entry be written</exception>
        public void Write(DirectoryEntry dir, String name)
        {
            /* If there is already an entry with the same name, remove it. */
            if(dir.HasEntry(name))
            {
                Entry e = dir.GetEntry(name);
                e.Delete();
            }

            /* Create the new entry. */
            dir.CreateDocument(name, ToInputStream());
        }

        /// <summary>
        /// Returns the contents of this property Set stream as an input stream.
        /// The latter can be used for example to write the property Set into a POIFS
        /// document. The input stream represents a snapshot of the property Set.
        /// If the latter is modified while the input stream is still being
        /// read, the modifications will not be reflected in the input stream but in
        /// the <see cref="MutablePropertySet"/> only.
        /// </summary>
        /// <return>contents of this property Set stream</return>
        /// 
        /// <exception cref="WritingNotSupportedException">if HPSF does not yet support writing
        /// of a property's variant type.
        /// </exception>
        /// <exception cref="IOException">if an I/O exception occurs.</exception>
        public InputStream ToInputStream()
        {
            MemoryStream psStream = new MemoryStream();
            try
            {
                Write(psStream);
            }
            finally
            {
                psStream.Close();
            }
            byte[] streamData = psStream.ToArray();
            return new ByteArrayInputStream(streamData);
        }

        /// <summary>
        /// Fetches the property with the given ID, then does its
        /// best to return it as a String
        /// </summary>
        /// <param name="propertyId">the property id</param>
        /// 
        /// <return>property as a String, or null if unavailable</return>
        protected String GetPropertyStringValue(int propertyId)
        {
            Object propertyValue = GetProperty(propertyId);
            return GetPropertyStringValue(propertyValue);
        }

        /// <summary>
        /// Return the string representation of a property value
        /// </summary>
        /// <param name="propertyValue">the property value</param>
        /// 
        /// <return>property value as a String, or null if unavailable</return>
        public static String GetPropertyStringValue(Object propertyValue)
        {
            // Normal cases
            if(propertyValue == null)
            {
                return null;
            }
            if(propertyValue is String value)
            {
                return value;
            }

            // Do our best with some edge cases
            if(propertyValue is byte[] data)
            {
                switch(data.Length)
                {
                    case 0:
                        return "";
                    case 1:
                        return data[0].ToString();
                    case 2:
                        return LittleEndian.GetUShort(data).ToString();
                    case 4:
                        return LittleEndian.GetUInt(data).ToString();
                    default:
                        // Maybe it's a string? who knows!
                        return Encoding.ASCII.GetString(data);
                }
            }
            return propertyValue.ToString();
        }

        /// <summary>
        /// Checks whether this <see cref="PropertySet"/> represents a Summary Information.
        /// </summary>
        /// <return>true if this <see cref="PropertySet"/>
        /// represents a Summary Information, else <c>false</c>.
        /// </return>
        public bool IsSummaryInformation => matchesSummary(SectionIDMap.SUMMARY_INFORMATION_ID);

        /// <summary>
        /// Checks whether this <see cref="PropertySet"/> is a Document Summary Information.
        /// </summary>
        /// <return>true if this <see cref="PropertySet"/>
        /// represents a Document Summary Information, else <c>false</c>.
        /// </return>
        public bool IsDocumentSummaryInformation => matchesSummary(SectionIDMap.DOCUMENT_SUMMARY_INFORMATION_ID[0]);


        private bool matchesSummary(byte[] summaryBytes)
        {
            return !(sections.Count == 0) &&
                Arrays.Equals(FirstSection.FormatID.Bytes, summaryBytes);
        }



        /// <summary>
        /// Convenience method returning the {@link Property} array contained in this
        /// property Set. It is a shortcut for Getting he <see cref="PropertySet"/>'s
        /// <see cref="Section"/>s list and then Getting the <see cref="Property"/> array from the
        /// first <see cref="Section"/>.
        /// </summary>
        /// <return>properties of the only <see cref="Section"/>
        /// </return>
        /// <exception cref="NoSingleSectionException">if the <see cref="PropertySet"/> has
        /// more or less than one <see cref="Section"/>.
        /// </exception>
        public Property[] Properties => FirstSection.Properties;


        /// <summary>
        /// Convenience method returning the value of the property with the specified ID.
        /// If the property is not available, <c>null</c> is returned and a subsequent
        /// call to <see cref="WasNull"/> will return <c>true</c>.
        /// </summary>
        /// <param name="id">The property ID</param>
        /// <return>property value</return>
        /// <exception cref="NoSingleSectionException">if the <see cref="PropertySet"/> has
        /// more or less than one <see cref="Section"/>.
        /// </exception>
        protected Object GetProperty(int id)
        {
            return FirstSection.GetProperty(id);
        }



        /// <summary>
        /// Convenience method returning the value of a bool property with the
        /// specified ID. If the property is not available, <c>false</c> is returned.
        /// A subsequent call to <see cref="WasNull"/> will return <c>true</c> to let the
        /// caller distinguish that case from a real property value of <c>false</c>.
        /// </summary>
        /// <param name="id">The property ID</param>
        /// <return>property value</return>
        /// <exception cref="NoSingleSectionException">if the <see cref="PropertySet"/> has
        /// more or less than one <see cref="Section"/>.
        /// </exception>
        protected bool GetPropertyBooleanValue(int id)
        {

            return FirstSection.GetPropertyBooleanValue(id);
        }



        /// <summary>
        /// Convenience method returning the value of the numeric
        /// property with the specified ID. If the property is not
        /// available, 0 is returned. A subsequent call to <see cref="WasNull"/>
        /// will return <c>true</c> to let the caller distinguish
        /// that case from a real property value of 0.
        /// </summary>
        /// <param name="id">The property ID</param>
        /// <return>propertyIntValue value</return>
        /// <exception cref="NoSingleSectionException">if the <see cref="PropertySet"/> has
        /// more or less than one <see cref="Section"/>.
        /// </exception>
        protected int GetPropertyIntValue(int id)
        {

            return FirstSection.GetPropertyIntValue(id);
        }



        /// <summary>
        /// Checks whether the property which the last call to <see cref="GetPropertyIntValue"/> 
        /// or <see cref="Property"/> tried to access
        /// was available or not. This information might be important for
        /// callers of <see cref="GetPropertyIntValue"/> since the latter
        /// returns 0 if the property does not exist. Using <see cref="WasNull"/>,
        ///  the caller can distiguish this case from a property's real value of 0.
        /// </summary>
        /// <return>true if the last call to <see cref="GetPropertyIntValue"/>
        /// or <see cref="Property"/> tried to access a
        /// property that was not available, else <c>false</c>.
        /// </return>
        /// <exception cref="NoSingleSectionException">if the <see cref="PropertySet"/> has
        /// more than one <see cref="Section"/>.
        /// </exception>
        public bool WasNull => FirstSection.WasNull();


        /// <summary>
        /// Gets the <see cref="PropertySet"/>'s first section.
        /// </summary>
        /// <return><see cref="PropertySet"/>'s first section.</return>
        public Section FirstSection
        {
            get
            {
                if(sections.Count==0)
                {
                    throw new MissingSectionException("Property Set does not contain any sections.");
                }
                return sections[0];
            }

        }



        /// <summary>
        /// If the <see cref="PropertySet"/> has only a single section this method returns it.
        /// </summary>
        /// <return>singleSection value</return>
        public Section SingleSection
        {
            get
            {
                int sectionCount = SectionCount;
                if(sectionCount != 1)
                {
                    throw new NoSingleSectionException("Property Set contains " + sectionCount + " sections.");
                }
                return sections[0];
            }
        }



        /// <summary>
        /// Returns <c>true</c> if the <c>PropertySet</c> is equal
        /// to the specified parameter, else <c>false</c>.
        /// </summary>
        /// <param name="o">the object to compare this <c>PropertySet</c> with</param>
        /// 
        /// <return>true if the objects are equal, <c>false</c>
        /// if not
        /// </return>
        public override bool Equals(object o)
        {
            if(o == null || o is not PropertySet ps)
            {
                return false;
            }

            int byteOrder1 = ps.ByteOrder;
            int byteOrder2 = ByteOrder;
            ClassID classID1 = ps.ClassID;
            ClassID classID2 = ClassID;
            int format1 = ps.Format;
            int format2 = Format;
            int osVersion1 = ps.OSVersion;
            int osVersion2 = OSVersion;
            int sectionCount1 = ps.SectionCount;
            int sectionCount2 = SectionCount;
            if(byteOrder1 != byteOrder2 ||
                !classID1.Equals(classID2) ||
                format1 != format2 ||
                osVersion1 != osVersion2 ||
                sectionCount1 != sectionCount2)
            {
                return false;
            }

            /* Compare the sections: */
            return Util.AreEqual(Sections, ps.Sections);
        }

        /// <summary>
        /// </summary>
        /// @see Object#hashCode()
        public override int GetHashCode()
        {
            throw new NotImplementedException("FIXME: Not yet implemented.");
        }



        /// <summary>
        /// </summary>
        /// @see Object#toString()
        public override String ToString()
        {
            StringBuilder b = new StringBuilder();
            int sectionCount = SectionCount;
            b.Append(GetType().Name);
            b.Append('[');
            b.Append("byteOrder: ");
            b.Append(ByteOrder);
            b.Append(", classID: ");
            b.Append(ClassID);
            b.Append(", format: ");
            b.Append(Format);
            b.Append(", OSVersion: ");
            b.Append(OSVersion);
            b.Append(", sectionCount: ");
            b.Append(sectionCount);
            b.Append(", sections: [\n");
            foreach(Section section in Sections)
            {
                b.Append(section);
            }
            b.Append(']');
            b.Append(']');
            return b.ToString();
        }


        protected void Remove1stProperty(long id)
        {
            FirstSection.RemoveProperty(id);
        }

        protected void Set1stProperty(long id, String value)
        {
            FirstSection.SetProperty((int) id, value);
        }

        protected void Set1stProperty(long id, int value)
        {
            FirstSection.SetProperty((int) id, value);
        }

        protected void Set1stProperty(long id, bool value)
        {
            FirstSection.SetProperty((int) id, value);
        }

        protected void Set1stProperty(long id, byte[] value)
        {
            FirstSection.SetProperty((int) id, value);
        }
    }
}

