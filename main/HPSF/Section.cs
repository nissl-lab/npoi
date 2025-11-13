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

    //using TreeBidiMap;
    using NPOI.HPSF.Wellknown;
    using NPOI.Util;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;
    using System.IO;
    using System.Text;
    using Cysharp.Text;

    /// <summary>
    /// Represents a section in a {@link PropertySet}.
    /// </summary>
    public class Section
    {

        /// <summary>
        /// <para>
        /// Maps property IDs to section-private PID strings. These
        /// strings can be found in the property with ID 0.
        /// </para>
        /// <para>
        /// Note: the type should be Dictionary&lt;long, String&gt; .
        /// </para></summary>
        //private Dictionary<long, String> dictionary;

        private Dictionary<long, String> dictionary;

        /// <summary>
        /// The section's format ID, <see cref="FormatID"/>.
        /// </summary>
        private ClassID formatID;

        /// <summary>
        /// Contains the bytes making out the section. This byte array is
        /// established when the section's size is calculated and can be reused
        /// later. If the array is empty, the section was modified and the bytes need to be regenerated.
        /// </summary>
        private ByteArrayOutputStream sectionBytes = new ByteArrayOutputStream();

        /// <summary>
        /// The offset of the section in the stream.
        /// </summary>
        private long _offset;

        /// <summary>
        /// This section's properties.
        /// </summary>
        private IDictionary<long, Property> properties = new Dictionary<long, Property>();

        /// <summary>
        /// This member is <c>true</c> if the last call to {@link
        /// #getPropertyIntValue} or <see cref="Property"/> tried to access a
        /// property that was not available, else <c>false</c>.
        /// </summary>
        private bool wasNull;

        /// <summary>
        /// Creates an empty {@link Section}.
        /// </summary>
        public Section()
        {
            this._offset = -1;
        }

        /// <summary>
        /// Constructs a <c>Section</c> by doing a deep copy of an
        /// existing <c>Section</c>. All nested <c>Property</c>
        /// instances, will be their mutable counterparts in the new
        /// <c>MutableSection</c>.
        /// </summary>
        /// <param name="s">The section Set to copy</param>
        public Section(Section s)
        {
            this._offset = -1;
            FormatID = (s.FormatID);
            foreach(Property p in s.properties.Values)
            {
                properties.Add(p.ID, new MutableProperty(p));
            }
            SetDictionary(s.Dictionary);
        }



        /// <summary>
        /// Creates a {@link Section} instance from a byte array.
        /// </summary>
        /// <param name="src">Contains the complete property Set stream.</param>
        /// <param name="offset">The position in the stream that points to the
        /// section's format ID.
        /// </param>
        /// 
        /// <exception name="UnsupportedEncodingException">if the section's codepage is not
        /// supported.
        /// </exception>  
        public Section(byte[] src, int offset)
        {
            /*
             * Read the format ID.
             */
            formatID = new ClassID(src, offset);

            /*
             * Read the offset from the stream's start and positions to
             * the section header.
             */
            int offFix = (int)LittleEndian.GetUInt(src, offset + ClassID.LENGTH);

            // some input files have a invalid (padded?) offset, which need to be fixed
            // search for beginning of size field
            if (src[offFix] == 0) {
                for (int i=0; i<3 && src[offFix] == 0; i++,offFix++);
                // cross check with propertyCount field and the property list field
                for (int i=0; i<3 && (src[offFix+3] != 0 || src[offFix+7] != 0 || src[offFix+11] != 0); i++,offFix--);
            }

            this._offset = offFix;

            LittleEndianByteArrayInputStream leis = new LittleEndianByteArrayInputStream(src, offFix);

            /*
             * Read the section length.
             */
            int size = (int)Math.Min(leis.ReadUInt(), src.Length-_offset);

            /*
             * Read the number of properties.
             */
            int propertyCount = (int)leis.ReadUInt();

            /*
             * Read the properties. The offset is positioned at the first
             * entry of the property list. There are two problems:
             *
             * 1. For each property we have to find out its length. In the
             *    property list we find each property's ID and its offset relative
             *    to the section's beginning. Unfortunately the properties in the
             *    property list need not to be in ascending order, so it is not
             *    possible to calculate the length as
             *    (offset of property(i+1) - offset of property(i)). Before we can
             *    that we first have to sort the property list by ascending offsets.
             *
             * 2. We have to read the property with ID 1 before we read other
             *    properties, at least before other properties containing strings.
             *    The reason is that property 1 specifies the codepage. If it is
             *    1200, all strings are in Unicode. In other words: Before we can
             *    read any strings we have to know whether they are in Unicode or
             *    not. Unfortunately property 1 is not guaranteed to be the first in
             *    a section.
             *
             *    The algorithm below reads the properties in two passes: The first
             *    one looks for property ID 1 and extracts the codepage number. The
             *    seconds pass reads the other properties.
             */
            /* Pass 1: Read the property list. */
            //TreeBidiDictionary<long, long> offset2Id = new TreeBidiDictionary<long, long>();
            BidirectionalDictionary<long, long> offset2Id = new ();
            for(int i = 0; i < propertyCount; i++)
            {
                /* Read the property ID. */
                long id = (int)leis.ReadUInt();

                /* Offset from the section's start. */
                long off = (int)leis.ReadUInt();

                offset2Id.Add(off, id);
            }

            long? cpOffset = offset2Id.ContainsValue((long)PropertyIDMap.PID_CODEPAGE) ? 
                offset2Id.GetKey((long)PropertyIDMap.PID_CODEPAGE) : null;
            /* Look for the codepage. */
            int codepage = -1;
            if(cpOffset.HasValue)
            {
                /* Read the property's value type. It must be VT_I2. */
                leis.SetReadIndex((int)(this._offset + cpOffset.Value));
                long type = leis.ReadUInt();

                if(type != Variant.VT_I2)
                {
                    throw new HPSFRuntimeException
                        ("Value type of property ID 1 is not VT_I2 but " + type + ".");
                }

                /* Read the codepage number. */
                codepage = leis.ReadUShort();
                Codepage = codepage;
            }


            /* Pass 2: Read all properties - including the codepage property,
             * if available. */
            foreach(KeyValuePair<long, long> me in offset2Id)
            {
                long off = me.Key;
                long id = me.Value;
                if (id == PropertyIDMap.PID_CODEPAGE)
                {
                    continue;
                }
                int pLen = propLen(offset2Id, off, size);
                leis.SetReadIndex((int)(this._offset + off));

                if (id == PropertyIDMap.PID_DICTIONARY)
                {
                    leis.Mark(100000);
                    if (!ReadDictionary(leis, pLen, codepage))
                    {
                        // there was an error reading the dictionary, maybe because the pid (0) was used wrong
                        // try reading a property instead
                        leis.Reset();
                        try
                        {
                            // fix id
                            id = Math.Max((long)PropertyIDMap.PID_MAX, offset2Id.Inverse.LastKey())+1;
                            SetProperty(new MutableProperty(id, leis, pLen, codepage));
                        }
                        catch (RuntimeException e)
                        {
                            //LOG.log(POILogger.INFO, "Dictionary fallback failed - ignoring property");
                        }
                    }
                }
                else
                {
                    SetProperty(new MutableProperty(id, leis, pLen, codepage));
                }
            }
            sectionBytes.Write(src, (int)_offset, size);
            padSectionBytes();
            /*
             * Extract the dictionary (if available).
             */
            dictionary = (Dictionary<long, String>) GetProperty(0);
        }

        /// <summary>
        /// Retrieves the length of the given property (by key)
        /// </summary>
        /// <param name="offset2Id">the offset to id map</param>
        /// <param name="entryOffset">the current entry key</param>
        /// <param name="maxSize">the maximum offset/size of the section stream</param>
        /// <return>length of the current property</return>
        private static int propLen(
            BidirectionalDictionary<long, long> offset2Id,
            long entryOffset,
            long maxSize)
        {
            long? nextKey = NextKey(offset2Id, entryOffset);
            long begin = entryOffset;
            long end = nextKey ?? maxSize;
            return (int) (end - begin);
        }

        private static long? NextKey(BidirectionalDictionary<long, long> offset2Id, long entryOffset)
        {
            var list = offset2Id.Keys.Where(k => k > entryOffset).ToList();
            list.Sort();
            return list.Count == 0 ? null : list.First();
        }

        /// <summary>
        /// get or set the format ID. The format ID is the "type" of the
        /// section. For example, if the format ID of the first {@link
        /// Section} contains the bytes specified by
        /// {@code NPOI.HPSF.wellknown.SectionIDMap.SUMMARY_INFORMATION_ID}
        /// the section (and thus the property Set) is a SummaryInformation.
        /// </summary>
        /// <return>format ID</return>
        public ClassID FormatID
        {
            get => this.formatID;
            set => this.formatID = value;
        }

        /// <summary>
        /// Sets the section's format ID.
        /// </summary>
        /// <param name="formatID">The section's format ID as a byte array. It components
        /// are in big-endian format.
        /// </param>
        public void SetFormatID(byte[] formatID)
        {
            ClassID fid = FormatID;
            if(fid == null)
            {
                fid = new ClassID();
                FormatID = (fid);
            }
            fid.Bytes = (formatID);
        }

        /// <summary>
        /// Returns the offset of the section in the stream.
        /// </summary>
        /// <return>offset of the section in the stream.</return>
        public long Offset => this._offset;

        /// <summary>
        /// Returns the number of properties in this section.
        /// </summary>
        /// <return>number of properties in this section.</return>
        public int PropertyCount => properties.Count;

        /// <summary>
        /// Returns this section's properties.
        /// </summary>
        /// <return>section's properties.</return>
        public Property[] Properties => properties.Values.ToArray();

        /// <summary>
        /// Sets this section's properties. Any former values are overwritten.
        /// </summary>
        /// <param name="properties">This section's new properties.</param>
        public void SetProperties(Property[] properties)
        {
            this.properties.Clear();
            foreach(Property p in properties)
            {
                SetProperty(p);
            }
        }

        /// <summary>
        /// Returns the value of the property with the specified ID. If
        /// the property is not available, <c>null</c> is returned
        /// and a subsequent call to <see cref="wasNull"/> will return
        /// <c>true</c>.
        /// </summary>
        /// <param name="id">The property's ID</param>
        /// 
        /// <return>property's value</return>
        public Object GetProperty(long id)
        {
            wasNull = !properties.ContainsKey(id);
            return (wasNull) ? null : properties[id].Value;
        }

        /// <summary>
        /// Sets the string value of the property with the specified ID.
        /// </summary>
        /// <param name="id">The property's ID</param>
        /// <param name="value">The property's value. It will be written as a Unicode
        /// string.
        /// </param>
        public void SetProperty(int id, String value)
        {
            SetProperty(id, Variant.VT_LPSTR, value);
        }

        /// <summary>
        /// Sets the int value of the property with the specified ID.
        /// </summary>
        /// <param name="id">The property's ID</param>
        /// <param name="value">The property's value.</param>
        /// 
        /// @see #setProperty(int, long, Object)
        /// @see #getProperty
        public void SetProperty(int id, int value)
        {
            SetProperty(id, Variant.VT_I4, value);
        }



        /// <summary>
        /// Sets the long value of the property with the specified ID.
        /// </summary>
        /// <param name="id">The property's ID</param>
        /// <param name="value">The property's value.</param>
        /// 
        /// @see #setProperty(int, long, Object)
        /// @see #getProperty
        public void SetProperty(int id, long value)
        {
            SetProperty(id, Variant.VT_I8, value);
        }



        /// <summary>
        /// Sets the bool value of the property with the specified ID.
        /// </summary>
        /// <param name="id">The property's ID</param>
        /// <param name="value">The property's value.</param>
        /// 
        /// @see #setProperty(int, long, Object)
        /// @see #getProperty
        public void SetProperty(int id, bool value)
        {
            SetProperty(id, Variant.VT_BOOL, value);
        }



        /// <summary>
        /// Sets the value and the variant type of the property with the
        /// specified ID. If a property with this ID is not yet present in
        /// the section, it will be added. An already present property with
        /// the specified ID will be overwritten. A default mapping will be
        /// used to choose the property's type.
        /// </summary>
        /// <param name="id">The property's ID.</param>
        /// <param name="variantType">The property's variant type.</param>
        /// <param name="value">The property's value.</param>
        /// 
        /// @see #setProperty(int, String)
        /// @see #getProperty
        /// @see Variant
        public void SetProperty(int id, long variantType, Object value)
        {
            SetProperty(new MutableProperty(id, variantType, value));
        }



        /// <summary>
        /// Sets a property.
        /// </summary>
        /// <param name="p">The property to be Set.</param>
        /// 
        /// @see #setProperty(int, long, Object)
        /// @see #getProperty
        /// @see Variant
        public void SetProperty(Property p)
        {
            Property old = properties.TryGetValue(p.ID, out Property value) ? value : null;
            if(old == null || !old.Equals(p))
            {
                properties[p.ID] = p;
                sectionBytes.Reset();
            }
        }

        /// <summary>
        /// Sets a property.
        /// </summary>
        /// <param name="id">The property ID.</param>
        /// <param name="value">The property's value. The value's class must be one of those
        /// supported by HPSF.
        /// </param>
        public void SetProperty(int id, Object value)
        {
            if(value is String s)
            {
                SetProperty(id, s);
            }
            else if(value is long l)
            {
                SetProperty(id, l);
            }
            else if(value is int i)
            {
                SetProperty(id, i);
            }
            else if(value is short s1)
            {
                SetProperty(id, s1);
            }
            else if(value is Boolean b)
            {
                SetProperty(id, b);
            }
            else if(value is Date)
            {
                SetProperty(id, Variant.VT_FILETIME, value);
            }
            else
            {
                throw new HPSFRuntimeException(
                        "HPSF does not support properties of type " +
                        value.GetType().Name + ".");
            }
        }

        /// <summary>
        /// Returns the value of the numeric property with the specified
        /// ID. If the property is not available, 0 is returned. A
        /// subsequent call to <see cref="wasNull"/> will return
        /// <c>true</c> to let the caller distinguish that case from
        /// a real property value of 0.
        /// </summary>
        /// <param name="id">The property's ID</param>
        /// 
        /// <return>property's value</return>
        protected internal int GetPropertyIntValue(long id)
        {
            int i;
            Object o = GetProperty(id);
            if(o == null)
            {
                return 0;
            }
            if(!(o is long || o is int))
            {
                throw new HPSFRuntimeException
                    ("This property is not an integer type, but " +
                     o.GetType().Name + ".");
            }
            i = (int) o;
            return i;
        }

        /// <summary>
        /// Returns the value of the bool property with the specified
        /// ID. If the property is not available, <c>false</c> is
        /// returned. A subsequent call to <see cref="wasNull"/> will return
        /// <c>true</c> to let the caller distinguish that case from
        /// a real property value of <c>false</c>.
        /// </summary>
        /// <param name="id">The property's ID</param>
        /// 
        /// <return>property's value</return>
        protected internal bool GetPropertyBooleanValue(int id)
        {
            object b = GetProperty(id);
            if(b == null)
            {
                return false;
            }
            return (bool) b;
        }

        /// <summary>
        /// Sets the value of the bool property with the specified
        /// ID.
        /// </summary>
        /// <param name="id">The property's ID</param>
        /// <param name="value">The property's value</param>
        /// 
        /// @see #setProperty(int, long, Object)
        /// @see #getProperty
        /// @see Variant
        protected void SetPropertyBooleanValue(int id, bool value)
        {
            SetProperty(id, Variant.VT_BOOL, value);
        }

        /// <summary>
        /// </summary>
        /// <return>section's size in bytes.</return>
        public int Size
        {
            get
            {
                int size = (int)sectionBytes.Position;
                if (size > 0)
                {
                    return size;
                }
                try
                {
                    return CalcSize();
                }
                catch (HPSFRuntimeException ex)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    throw new HPSFRuntimeException(ex);
                }
            }
        }

        /// <summary>
        /// Calculates the section's size. It is the sum of the lengths of the
        /// section's header (8), the properties list (16 times the number of
        /// properties) and the properties themselves.
        /// </summary>
        /// <return>section's length in bytes.</return>
        /// <throws name="WritingNotSupportedException">WritingNotSupportedException</throws>
        /// <throws name="IOException">IOException</throws>
        private int CalcSize()
        {
            sectionBytes.Reset();
            Write(sectionBytes);
            padSectionBytes();
            return (int)sectionBytes.Position;
        }

        private void padSectionBytes()
        {
            byte[] padArray = { 0, 0, 0 };
            /* Pad to multiple of 4 bytes so that even the Windows shell (explorer)
             * shows custom properties. */
            int pad = (4 - ((int)sectionBytes.Position & 0x3)) & 0x3;
            sectionBytes.Write(padArray, 0, pad);
        }


        /// <summary>
        /// Checks whether the property which the last call to {@link
        /// #getPropertyIntValue} or <see cref="getProperty"/> tried to access
        /// was available or not. This information might be important for
        /// callers of <see cref="getPropertyIntValue"/> since the latter
        /// returns 0 if the property does not exist. Using {@link
        /// #wasNull} the caller can distiguish this case from a property's
        /// real value of 0.
        /// </summary>
        /// <return>true} if the last call to {@link
        /// #getPropertyIntValue} or <see cref="getProperty"/> tried to access a
        /// property that was not available, else <c>false</c>.
        /// </return>
        public bool WasNull()
        {
            return wasNull;
        }



        /// <summary>
        /// Returns the PID string associated with a property ID. The ID
        /// is first looked up in the {@link Section}'s private
        /// dictionary. If it is not found there, the method calls {@link
        /// SectionIDMap#getPIDString}.
        /// </summary>
        /// <param name="pid">The property ID</param>
        /// 
        /// <return>property ID's string value</return>
        public String GetPIDString(long pid)
        {
            String s = null;
            if(dictionary != null)
            {
                s = (string) dictionary[pid];
            }
            if(s == null)
            {
                s = SectionIDMap.GetPIDString(FormatID, pid);
            }
            return s;
        }

        /// <summary>
        /// Removes all properties from the section including 0 (dictionary) and
        /// 1 (codepage).
        /// </summary>
        public void Clear()
        {
            Property[] properties = Properties;
            for(int i = 0; i < properties.Length; i++)
            {
                Property p = properties[i];
                RemoveProperty(p.ID);
            }
        }

        /// <summary>
        /// <para>
        /// Checks whether this section is equal to another object. The result is
        /// <c>false</c> if one of the the following conditions holds:
        /// </para>
        /// <para>
        /// <ul>
        /// </para>
        /// <para>
        /// <li>The other object is not a {@link Section}.</li>
        /// </para>
        /// <para>
        /// <li>The format IDs of the two sections are not equal.</li>
        /// </para>
        /// <para>
        /// <li>The sections have a different number of properties. However,
        /// properties with ID 1 (codepage) are not counted.</li>
        /// </para>
        /// <para>
        /// <li>The other object is not a {@link Section}.</li>
        /// </para>
        /// <para>
        /// <li>The properties have different values. The order of the properties
        /// is irrelevant.</li>
        /// </para>
        /// <para>
        /// </ul>
        /// </para>
        /// </summary>
        /// <param name="o">The object to compare this section with</param>
        /// <return>true} if the objects are equal, <c>false</c> if
        /// not
        /// </return>
        public override bool Equals(Object o)
        {
            if(o == null || o is not Section section)
            {
                return false;
            }

            if(!section.FormatID.Equals(FormatID))
            {
                return false;
            }

            /* Compare all properties except the dictionary (id 0) and
             * the codepage (id 1 / ignored) as they must be handled specially. */
            HashSet<long> propIds = [.. properties.Keys, .. section.properties.Keys];
            propIds.Remove(0L);
            propIds.Remove(1L);

            foreach (long id in propIds) {
                properties.TryGetValue(id, out Property p1);
                section.properties.TryGetValue(id, out Property p2);
                if (p1 == null || p2 == null || !p1.Equals(p2))
                {
                    return false;
                }
            }

            /* If the dictionaries are unequal the sections are unequal. */
            Dictionary<long, String> d1 = Dictionary;
            Dictionary<long, String> d2 = section.Dictionary;

            return (d1 == null && d2 == null) || (d1 != null && d2 != null && CompareDictionaries(d1, d2));
        }

        static bool CompareDictionaries(Dictionary<long, string> dict1, IDictionary dict2)
        {
            if(dict1.Count != dict2.Count)
            {
                return false;
            }

            foreach(KeyValuePair<long, string> pair in dict1)
            {
                if(!dict2.Contains(pair.Key) || !Equals(pair.Value, dict2[pair.Key]))
                {
                    return false;
                }
            }

            return true;
        }
        /// <summary>
        /// Removes a property.
        /// </summary>
        /// <param name="id">The ID of the property to be removed</param>
        public void RemoveProperty(long id)
        {
            if(properties.Remove(id))
            {
                sectionBytes.Reset();
            }
        }

        /// <summary>
        /// Removes a field from a property array. The resulting array is
        /// compactified and returned.
        /// </summary>
        /// <param name="pa">The property array.</param>
        /// <param name="i">The index of the field to be removed.</param>
        /// <return>compactified array.</return>
        [Obsolete]
        private static Property[] Remove(Property[] pa, int i)
        {
            Property[] h = new Property[pa.Length - 1];
            if(i > 0)
            {
                System.Array.Copy(pa, 0, h, 0, i);
            }
            System.Array.Copy(pa, i + 1, h, i, h.Length - i);
            return h;
        }
        /// <summary>
        /// <para>
        /// </para>
        /// <para>
        /// Writes this section into an output stream.
        /// </para>
        /// <para>
        /// Internally this is done by writing into three byte array output
        /// streams: one for the properties, one for the property list and one for
        /// the section as such. The two former are appended to the latter when they
        /// have received all their data.
        /// </para>
        /// </summary>
        /// <param name="out">The stream to write into.</param>
        /// 
        /// <return>number of bytes written, i.e. the section's size.</return>
        /// <exception name="IOException">if an I/O error occurs</exception>
        /// <exception name="WritingNotSupportedException">if HPSF does not yet support
        /// writing a property's variant type.
        /// </exception>
        public int Write(Stream out1)
        {

            /* Check whether we have already generated the bytes making out the
             * section. */
            if (sectionBytes.Position > 0)
            {
                out1.Write(sectionBytes.ToByteArray(), 0, (int)sectionBytes.Length);
                return (int)sectionBytes.Position;
            }

            /* Writing the section's dictionary it tricky. If there is a dictionary
             * (property 0) the codepage property (property 1) must be set, too. */
            int codepage = Codepage;
            if (codepage == -1)
            {
                //String msg =
                //    "The codepage property is not set although a dictionary is present. "+
                //    "Defaulting to ISO-8859-1.";
                //LOG.log(POILogger.WARN, msg);
                codepage = Property.DEFAULT_CODEPAGE;
            }

            /* The properties are written to this stream. */
            MemoryStream propertyStream = new MemoryStream();

            /* The property list is established here. After each property that has
             * been written to "propertyStream", a property list entry is written to
             * "propertyListStream". */
            MemoryStream propertyListStream = new MemoryStream();

            /* Maintain the current position in the list. */
            int position = 0;

            /* Increase the position variable by the size of the property list so
             * that it points behind the property list and to the beginning of the
             * properties themselves. */
            position += 2 * LittleEndianConsts.INT_SIZE + PropertyCount * 2 * LittleEndianConsts.INT_SIZE;

            /* Write the properties and the property list into their respective
             * streams: */
            foreach(Property p in properties.Values)
            {
                long id = p.ID;

                /* Write the property list entry. */
                LittleEndian.PutUInt(id, propertyListStream);
                LittleEndian.PutUInt(position, propertyListStream);
                /* If the property ID is not equal 0 we write the property and all
                 * is fine. However, if it equals 0 we have to write the section's
                 * dictionary which has an implicit type only and an explicit
                 * value. */
                if(id != 0)
                {
                    /* Write the property and update the position to the next
                     * property. */
                    position += p.Write(propertyStream, Codepage);
                }
                else
                {
                    if(codepage == -1)
                    {
                        throw new IllegalPropertySetDataException("Codepage (property 1) is undefined.");
                    }
                    position += WriteDictionary(propertyStream, codepage);
                }
            }
            
            /* Write the section: */
            int streamLength = LittleEndianConsts.INT_SIZE * 2 + 
                (int)propertyListStream.Length + (int)propertyStream.Length;

            /* Write the section's length: */
            LittleEndian.PutInt(streamLength, out1);

            /* Write the section's number of properties: */
            LittleEndian.PutInt(PropertyCount, out1);

            /* Write the property list: */
            propertyListStream.WriteTo(out1);

            /* Write the properties: */
            propertyStream.WriteTo(out1);

            return streamLength;
        }

        /**
         * Reads a dictionary.
         *
         * @param leis The byte stream containing the bytes making out the dictionary.
         * @param length The dictionary contains at most this many bytes.
         * @param codepage The codepage of the string values.
         *
         * @return {@code true} if dictionary was read successful, {@code false} otherwise
         *
         * @throws UnsupportedEncodingException if the dictionary's codepage is not
         *         (yet) supported.
         */
        private bool ReadDictionary(LittleEndianByteArrayInputStream leis, int length, int codepage)
        {
            Dictionary<long,String> dic = new Dictionary<long,String>();

            /*
             * Read the number of dictionary entries.
             */
            long nrEntries = leis.ReadUInt();

            long id = -1;
            bool isCorrupted = false;
            for (int i = 0; i < nrEntries; i++)
            {
                String errMsg =
                    "The property set's dictionary contains bogus data. "
                    + "All dictionary entries starting with the one with ID "
                    + id + " will be ignored.";

                /* The key. */
                id = leis.ReadUInt();

                /* The value (a string). The length is the either the
                 * number of (two-byte) characters if the character set is Unicode
                 * or the number of bytes if the character set is not Unicode.
                 * The length includes terminating 0x00 bytes which we have to strip
                 * off to create a Java string. */
                long sLength = leis.ReadUInt();

                /* Read the string - Strip 0x00 characters from the end of the string. */
                int cp = (codepage == -1) ? Property.DEFAULT_CODEPAGE : codepage;
                int nrBytes = (int)((sLength-1) * (cp == CodePageUtil.CP_UNICODE ? 2 : 1));
                if (nrBytes > 0xFFFFFF)
                {
                    //LOG.log(POILogger.WARN, errMsg);
                    isCorrupted = true;
                    break;
                }

                try 
                {
                    byte[] buf = new byte[nrBytes];
                    leis.ReadFully(buf, 0, nrBytes);
                    String str = CodePageUtil.GetStringFromCodePage(buf, 0, nrBytes, cp);

                    int pad = 1;
                    if (cp == CodePageUtil.CP_UNICODE)
                    {
                        pad = 2+((4 - ((nrBytes+2) & 0x3)) & 0x3);
                    }
                    leis.Skip(pad);

                    dic[id] = str;
                } 
                catch (RuntimeException ex)
                {
                    //LOG.log(POILogger.WARN, errMsg, ex);
                    isCorrupted = true;
                    break;
                }
            }
            SetDictionary(dic);
            return !isCorrupted;
        }

        /// <summary>
        /// Writes the section's dictionary.
        /// </summary>
        /// <param name="out">The output stream to write to.</param>
        /// <param name="dictionary">The dictionary.</param>
        /// <param name="codepage">The codepage to be used to write the dictionary items.</param>
        /// <return>number of bytes written</return>
        /// <exception name="IOException">if an I/O exception occurs.</exception>
        private int WriteDictionary(MemoryStream out1, int codepage)
        {
            byte[] padding = new byte[4];
            Dictionary<long,String> dic = Dictionary;
            LittleEndian.PutUInt(dic.Count, out1);
            int length = LittleEndianConsts.INT_SIZE;
            foreach(KeyValuePair<long, string> ls in dic)
            {
                LittleEndian.PutUInt(ls.Key, out1);
                length += LittleEndianConsts.INT_SIZE;

                String value = ls.Value+"\0";
                LittleEndian.PutUInt(value.Length, out1);
                length += LittleEndianConsts.INT_SIZE;

                byte[] bytes = CodePageUtil.GetBytesInCodePage(value, codepage);
                out1.Write(bytes, 0, bytes.Length);
                length += bytes.Length;

                if(codepage == CodePageUtil.CP_UNICODE)
                {
                    int pad = (4 - (length & 0x3)) & 0x3;
                    out1.Write(padding, 0, pad);
                    length += pad;
                }
            }
            int pad2 = (4 - (length & 0x3)) & 0x3;
            out1.Write(padding, 0, pad2);
            length += pad2;
            return length;
        }

        /// <summary>
        /// Sets the section's dictionary. All keys in the dictionary must be
        /// {@link java.lang.Long} instances, all values must be
        /// {@link java.lang.String}s. This method overwrites the properties with IDs
        /// 0 and 1 since they are reserved for the dictionary and the dictionary's
        /// codepage. Setting these properties explicitly might have surprising
        /// effects. An application should never do this but always use this
        /// method.
        /// </summary>
        /// <param name="dictionary">The dictionary</param>
        /// 
        /// <exception name="IllegalPropertySetDataException">if the dictionary's key and
        /// value types are not correct.
        /// </exception>
        /// 
        /// @see Section#getDictionary()
        public void SetDictionary(Dictionary<long, String> dictionary)
        {

            if(dictionary != null)
            {
                if(this.dictionary == null)
                {
                    this.dictionary = new Dictionary<long, String>();
                }
                foreach(var kv in dictionary)
                {
                    this.dictionary[kv.Key] = kv.Value;
                }
                
                /* If the codepage property (ID 1) for the strings (keys and values)
                 * used in the dictionary is not yet defined, set it to ISO-8859-1. */
                int cp = Codepage;
                if (cp == -1)
                {
                    Codepage = Property.DEFAULT_CODEPAGE;
                }
                /* Set the dictionary property (ID 0). Please note that the second
                 * parameter in the method call below is unused because dictionaries
                 * don't have a type. */
                SetProperty(PropertyIDMap.PID_DICTIONARY, -1, dictionary);

            }
            else
            {
                /* Setting the dictionary to null means to remove property 0.
                 * However, it does not mean to remove property 1 (codepage). */
                RemoveProperty(PropertyIDMap.PID_DICTIONARY);
                this.dictionary = null;
            }
        }



        /// <summary>
        /// </summary>
        /// @see Object#hashCode()
        public override int GetHashCode()
        {
            long hashCode = 0;
            hashCode += FormatID.GetHashCode();
            Property[] pa = Properties;
            for(int i = 0; i < pa.Length; i++)
            {
                hashCode += pa[i].GetHashCode();
            }
            int returnHashCode = (int)(hashCode & 0x0ffffffffL);
            return returnHashCode;
        }



        /// <summary>
        /// </summary>
        /// @see Object#toString()
        public override String ToString()
        {
            return ToString(null);
        }
        public String ToString(PropertyIDMap idMap)
        {
            using var b= ZString.CreateStringBuilder();
            Property[] pa = Properties;
            b.Append("\n\n\n");
            b.Append(GetType().Name);
            b.Append('[');
            b.Append("formatID: ");
            b.Append(FormatID);
            b.Append(", offset: ");
            b.Append(Offset);
            b.Append(", propertyCount: ");
            b.Append(PropertyCount);
            b.Append(", size: ");
            b.Append(Size);
            b.Append(", properties: [\n");
            int codepage = Codepage;
            if (codepage == -1)
            {
                codepage = Property.DEFAULT_CODEPAGE;
            }
            foreach(Property p in pa)
            {
                b.Append(p.ToString(codepage, idMap));
                b.Append(",\n");
            }
            b.Append(']');
            b.Append(']');
            return b.ToString();
        }



        /// <summary>
        /// Gets the section's dictionary. A dictionary allows an application to
        /// use human-readable property names instead of numeric property IDs. It
        /// contains mappings from property IDs to their associated string
        /// values. The dictionary is stored as the property with ID 0. The codepage
        /// for the strings in the dictionary is defined by property with ID 1.
        /// </summary>
        /// <return>dictionary or <c>null</c> if the section does not have
        /// a dictionary.
        /// </return>
        public Dictionary<long, String> Dictionary 
        {
            get
            {
                if(dictionary == null)
                {
                    dictionary = (Dictionary<long, String>)GetProperty(PropertyIDMap.PID_DICTIONARY);
                }
                return dictionary;
            }    
        }

        /// <summary>
        /// Gets the section's codepage, if any.
        /// </summary>
        /// <return>section's codepage if one is defined, else -1.</return>
        public int Codepage
        {
            get
            {
                object codepage = GetProperty(PropertyIDMap.PID_CODEPAGE);
                if(codepage == null)
                {
                    return -1;
                }
                int cp = (int)codepage;
                return cp;
            }
            set
            {
                SetProperty(PropertyIDMap.PID_CODEPAGE, Variant.VT_I2, value);
            }
        }
    }
}

