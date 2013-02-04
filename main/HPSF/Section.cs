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
    using System;
    using System.Text;
    using System.Collections;
    using NPOI.Util;
    using NPOI.HPSF.Wellknown;

    /// <summary>
    /// Represents a section in a {@link PropertySet}.
    /// @author Rainer Klute 
    /// <a href="mailto:klute@rainer-klute.de">&lt;klute@rainer-klute.de&gt;</a>
    /// @author Drew Varner (Drew.Varner allUpIn sc.edu)
    /// @since 2002-02-09
    /// </summary>
    public class Section
    {

        /**
         * Maps property IDs To section-private PID strings. These
         * strings can be found in the property with ID 0.
         */
        protected IDictionary dictionary;

        /**
         * The section's format ID, {@link #GetFormatID}.
         */
        protected ClassID formatID;


        /// <summary>
        /// Returns the format ID. The format ID is the "type" of the
        /// section. For example, if the format ID of the first {@link
        /// Section} Contains the bytes specified by 
        /// <c>org.apache.poi.hpsf.wellknown.SectionIDMap.SUMMARY_INFORMATION_ID</c>
        /// the section (and thus the property Set) is a SummaryInformation.
        /// </summary>
        /// <value>The format ID.</value>
        public ClassID FormatID
        {
            get { return formatID; }
        }

        protected long offset;


        /// <summary>
        /// Gets the offset of the section in the stream.
        /// </summary>
        /// <value>The offset of the section in the stream</value>
        public long OffSet
        {
            get { return offset; }
        }




        protected int size;


        /// <summary>
        /// Returns the section's size in bytes.
        /// </summary>
        /// <value>The section's size in bytes.</value>
        public virtual int Size
        {
            get { return size; }
        }



        /// <summary>
        /// Returns the number of properties in this section.
        /// </summary>
        /// <value>The number of properties in this section.</value> 
        public virtual int PropertyCount
        {
            get { return properties.Length; }
        }

        protected Property[] properties;


        /// <summary>
        /// Returns this section's properties.
        /// </summary>
        /// <value>This section's properties.</value>
        public virtual Property[] Properties
        {
            get { return properties; }
        }



        /// <summary>
        /// Creates an empty and uninitialized {@link Section}.
        /// </summary>
        protected Section()
        { }



        /// <summary>
        /// Creates a {@link Section} instance from a byte array.
        /// </summary>
        /// <param name="src">Contains the complete property Set stream.</param>
        /// <param name="offset">The position in the stream that points To the
        /// section's format ID.</param>
        public Section(byte[] src, int offset)
        {
            int o1 = offset;

            /*
             * Read the format ID.
             */
            formatID = new ClassID(src, o1);
            o1 += ClassID.LENGTH;

            /*
             * Read the offset from the stream's start and positions To
             * the section header.
             */
            this.offset = LittleEndian.GetUInt(src, o1);
            o1 = (int)this.offset;

            /*
             * Read the section Length.
             */
            size = (int)LittleEndian.GetUInt(src, o1);
            o1 += LittleEndianConsts.INT_SIZE;

            /*
             * Read the number of properties.
             */
            int propertyCount = (int)LittleEndian.GetUInt(src, o1);
            o1 += LittleEndianConsts.INT_SIZE;

            /*
             * Read the properties. The offset is positioned at the first
             * entry of the property list. There are two problems:
             * 
             * 1. For each property we have To Find out its Length. In the
             *    property list we Find each property's ID and its offset relative
             *    To the section's beginning. Unfortunately the properties in the
             *    property list need not To be in ascending order, so it is not
             *    possible To calculate the Length as
             *    (offset of property(i+1) - offset of property(i)). Before we can
             *    that we first have To sort the property list by ascending offsets.
             * 
             * 2. We have To Read the property with ID 1 before we Read other 
             *    properties, at least before other properties containing strings.
             *    The reason is that property 1 specifies the codepage. If it Is
             *    1200, all strings are in Unicode. In other words: Before we can
             *    Read any strings we have To know whether they are in Unicode or
             *    not. Unfortunately property 1 is not guaranteed To be the first in
             *    a section.
             *
             *    The algorithm below Reads the properties in two passes: The first
             *    one looks for property ID 1 and extracts the codepage number. The
             *    seconds pass Reads the other properties.
             */
            properties = new Property[propertyCount];

            /* Pass 1: Read the property list. */
            int pass1OffSet = o1;
            ArrayList propertyList = new ArrayList(propertyCount);
            PropertyListEntry ple;
            for (int i = 0; i < properties.Length; i++)
            {
                ple = new PropertyListEntry();

                /* Read the property ID. */
                ple.id = (int)LittleEndian.GetUInt(src, pass1OffSet);
                pass1OffSet += LittleEndianConsts.INT_SIZE;

                /* OffSet from the section's start. */
                ple.offset = (int)LittleEndian.GetUInt(src, pass1OffSet);
                pass1OffSet += LittleEndianConsts.INT_SIZE;

                /* Add the entry To the property list. */
                propertyList.Add(ple);
            }

            /* Sort the property list by ascending offsets: */
            propertyList.Sort();

            /* Calculate the properties' Lengths. */
            for (int i = 0; i < propertyCount - 1; i++)
            {
                PropertyListEntry ple1 =
                    (PropertyListEntry)propertyList[i];
                PropertyListEntry ple2 =
                    (PropertyListEntry)propertyList[i + 1];
                ple1.Length = ple2.offset - ple1.offset;
            }
            if (propertyCount > 0)
            {
                ple = (PropertyListEntry)propertyList[propertyCount - 1];
                ple.Length = size - ple.offset;
                //if (ple.Length <= 0)
                //{
                //    StringBuilder b = new StringBuilder();
                //    b.Append("The property Set claims To have a size of ");
                //    b.Append(size);
                //    b.Append(" bytes. However, it exceeds ");
                //    b.Append(ple.offset);
                //    b.Append(" bytes.");
                //    throw new IllegalPropertySetDataException(b.ToString());
                //}
            }

            /* Look for the codepage. */
            int codepage = -1;
            for (IEnumerator i = propertyList.GetEnumerator();
                 codepage == -1 && i.MoveNext(); )
            {
                ple = (PropertyListEntry)i.Current;

                /* Read the codepage if the property ID is 1. */
                if (ple.id == PropertyIDMap.PID_CODEPAGE)
                {
                    /* Read the property's value type. It must be
                     * VT_I2. */
                    int o = (int)(this.offset + ple.offset);
                    long type = LittleEndian.GetUInt(src, o);
                    o += LittleEndianConsts.INT_SIZE;

                    if (type != Variant.VT_I2)
                        throw new HPSFRuntimeException
                            ("Value type of property ID 1 is not VT_I2 but " +
                             type + ".");

                    /* Read the codepage number. */
                    codepage = LittleEndian.GetUShort(src, o);
                }
            }

            /* Pass 2: Read all properties - including the codepage property,
             * if available. */
            int i1 = 0;
            for (IEnumerator i = propertyList.GetEnumerator(); i.MoveNext(); )
            {
                ple = (PropertyListEntry)i.Current;
                Property p = new Property(ple.id, src,
                        this.offset + ple.offset,
                        ple.Length, codepage);
                if (p.ID == PropertyIDMap.PID_CODEPAGE)
                    p = new Property(p.ID, p.Type, codepage);
                properties[i1++] = p;
            }

            /*
             * Extract the dictionary (if available).
             * Tony changed the logic
             */
            this.dictionary = (IDictionary)GetProperty(0);
        }



        /**
         * Represents an entry in the property list and holds a property's ID and
         * its offset from the section's beginning.
         */
        class PropertyListEntry : IComparable
        {
            public int id;
            public int offset;
            public int Length;

            /**
             * Compares this {@link PropertyListEntry} with another one by their
             * offsets. A {@link PropertyListEntry} is "smaller" than another one if
             * its offset from the section's begin is smaller.
             *
             * @see Comparable#CompareTo(java.lang.Object)
             */
            public int CompareTo(Object o)
            {
                if (!(o is PropertyListEntry))
                    throw new InvalidCastException(o.ToString());
                int otherOffSet = ((PropertyListEntry)o).offset;
                if (offset < otherOffSet)
                    return -1;
                else if (offset == otherOffSet)
                    return 0;
                else
                    return 1;
            }

            public override String ToString()
            {
                StringBuilder b = new StringBuilder();
                b.Append(GetType().Name);
                b.Append("[id=");
                b.Append(id);
                b.Append(", offset=");
                b.Append(offset);
                b.Append(", Length=");
                b.Append(Length);
                b.Append(']');
                return b.ToString();
            }
        }



        /**
         * Returns the value of the property with the specified ID. If
         * the property is not available, <c>null</c> is returned
         * and a subsequent call To {@link #wasNull} will return
         * <c>true</c>.
         *
         * @param id The property's ID
         *
         * @return The property's value
         */
        public virtual Object GetProperty(long id)
        {
            wasNull = false;
            for (int i = 0; i < properties.Length; i++)
                if (id == properties[i].ID)
                    return properties[i].Value;
            wasNull = true;
            return null;
        }



        /**
         * Returns the value of the numeric property with the specified
         * ID. If the property is not available, 0 is returned. A
         * subsequent call To {@link #wasNull} will return
         * <c>true</c> To let the caller distinguish that case from
         * a real property value of 0.
         *
         * @param id The property's ID
         *
         * @return The property's value
         */
        public virtual int GetPropertyIntValue(long id)
        {
            Object o = GetProperty(id);
            if (o == null)
                return 0;
            if (!(o is long || o is int))
                throw new HPSFRuntimeException
                    ("This property is not an integer type, but " +
                     o.GetType().Name + ".");
            return (int)o;
        }



        /**
         * Returns the value of the bool property with the specified
         * ID. If the property is not available, <c>false</c> Is
         * returned. A subsequent call To {@link #wasNull} will return
         * <c>true</c> To let the caller distinguish that case from
         * a real property value of <c>false</c>.
         *
         * @param id The property's ID
         *
         * @return The property's value
         */
        public virtual bool GetPropertyBooleanValue(int id)
        {
            object tmp = GetProperty(id);
            if (tmp != null)
            {
                return (bool)GetProperty(id);
            }
            else
            {
                return false;
            }
        }



        /**
         * This member is <c>true</c> if the last call To {@link
         * #GetPropertyIntValue} or {@link #GetProperty} tried To access a
         * property that was not available, else <c>false</c>.
         */
        private bool wasNull;


        /// <summary>
        /// Checks whether the property which the last call To {@link
        /// #GetPropertyIntValue} or {@link #GetProperty} tried To access
        /// was available or not. This information might be important for
        /// callers of {@link #GetPropertyIntValue} since the latter
        /// returns 0 if the property does not exist. Using {@link
        /// #wasNull} the caller can distiguish this case from a property's
        /// real value of 0.
        /// </summary>
        /// <value><c>true</c> if the last call To {@link
        /// #GetPropertyIntValue} or {@link #GetProperty} tried To access a
        /// property that was not available; otherwise, <c>false</c>.</value>
        public virtual bool WasNull
        {
            get { return wasNull; }
        }



        /// <summary>
        /// Returns the PID string associated with a property ID. The ID
        /// is first looked up in the {@link Section}'s private
        /// dictionary. If it is not found there, the method calls {@link
        /// SectionIDMap#GetPIDString}.
        /// </summary>
        /// <param name="pid">The property ID.</param>
        /// <returns>The property ID's string value</returns>
        public String GetPIDString(long pid)
        {
            String s = null;
            if (dictionary != null)
                s = (String)dictionary[pid];
            if (s == null)
                s = SectionIDMap.GetPIDString(FormatID.Bytes, pid);
            if (s == null)
                s = SectionIDMap.UNDEFINED;
            return s;
        }



        /**
         * Checks whether this section is equal To another object. The result Is
         * <c>false</c> if one of the the following conditions holds:
         * 
         * <ul>
         * 
         * <li>The other object is not a {@link Section}.</li>
         * 
         * <li>The format IDs of the two sections are not equal.</li>
         *   
         * <li>The sections have a different number of properties. However,
         * properties with ID 1 (codepage) are not counted.</li>
         * 
         * <li>The other object is not a {@link Section}.</li>
         * 
         * <li>The properties have different values. The order of the properties
         * is irrelevant.</li>
         * 
         * </ul>
         * 
         * @param o The object To Compare this section with
         * @return <c>true</c> if the objects are equal, <c>false</c> if
         * not
         */
        public override bool Equals(Object o)
        {
            if (o == null || !(o is Section))
                return false;
            Section s = (Section)o;
            if (!s.FormatID.Equals(FormatID))
                return false;

            /* Compare all properties except 0 and 1 as they must be handled 
             * specially. */
            Property[] pa1 = new Property[Properties.Length];
            Property[] pa2 = new Property[s.Properties.Length];
            System.Array.Copy(Properties, 0, pa1, 0, pa1.Length);
            System.Array.Copy(s.Properties, 0, pa2, 0, pa2.Length);

            /* Extract properties 0 and 1 and Remove them from the copy of the
             * arrays. */
            Property p10 = null;
            Property p20 = null;
            for (int i = 0; i < pa1.Length; i++)
            {
                long id = pa1[i].ID;
                if (id == 0)
                {
                    p10 = pa1[i];
                    pa1 = Remove(pa1, i);
                    i--;
                }
                if (id == 1)
                {
                    // p11 = pa1[i];
                    pa1 = Remove(pa1, i);
                    i--;
                }
            }
            for (int i = 0; i < pa2.Length; i++)
            {
                long id = pa2[i].ID;
                if (id == 0)
                {
                    p20 = pa2[i];
                    pa2 = Remove(pa2, i);
                    i--;
                }
                if (id == 1)
                {
                    // p21 = pa2[i];
                    pa2 = Remove(pa2, i);
                    i--;
                }
            }

            /* If the number of properties (not counting property 1) is unequal the
             * sections are unequal. */
            if (pa1.Length != pa2.Length)
                return false;

            /* If the dictionaries are unequal the sections are unequal. */
            bool dictionaryEqual = true;
            if (p10 != null && p20 != null)
            {
                //tony qu fixed this issue
                Hashtable a=(Hashtable)p10.Value;
                Hashtable b = (Hashtable)p20.Value;
                dictionaryEqual = a.Count==b.Count;
            }
            else if (p10 != null || p20 != null)
            {
                dictionaryEqual = false;
            }
            if (!dictionaryEqual)
                return false;
            else
                return Util.AreEqual(pa1, pa2);
        }



        /// <summary>
        /// Removes a field from a property array. The resulting array Is
        /// compactified and returned.
        /// </summary>
        /// <param name="pa">The property array.</param>
        /// <param name="i">The index of the field To be Removed.</param>
        /// <returns>the compactified array.</returns>
        private Property[] Remove(Property[] pa, int i)
        {
            Property[] h = new Property[pa.Length - 1];
            if (i > 0)
                System.Array.Copy(pa, 0, h, 0, i);
            System.Array.Copy(pa, i + 1, h, i, h.Length - i);
            return h;
        }


        /// <summary>
        /// Serves as a hash function for a particular type.
        /// </summary>
        /// <returns>
        /// A hash code for the current <see cref="T:System.Object"/>.
        /// </returns>
        public override int GetHashCode()
        {
            long GetHashCode = 0;
            GetHashCode += FormatID.GetHashCode();
            Property[] pa = Properties;
            for (int i = 0; i < pa.Length; i++)
                GetHashCode += pa[i].GetHashCode();
            int returnHashCode = (int)(GetHashCode & 0x0ffffffffL);
            return returnHashCode;
        }




        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override String ToString()
        {
            StringBuilder b = new StringBuilder();
            Property[] pa = Properties;
            b.Append(GetType().Name);
            b.Append('[');
            b.Append("formatID: ");
            b.Append(FormatID);
            b.Append(", offset: ");
            b.Append(OffSet);
            b.Append(", propertyCount: ");
            b.Append(PropertyCount);
            b.Append(", size: ");
            b.Append(Size);
            b.Append(", properties: [\n");
            for (int i = 0; i < pa.Length; i++)
            {
                b.Append(pa[i].ToString());
                b.Append(",\n");
            }
            b.Append(']');
            b.Append(']');
            return b.ToString();
        }



        /// <summary>
        /// Gets the section's dictionary. A dictionary allows an application To
        /// use human-Readable property names instead of numeric property IDs. It
        /// Contains mappings from property IDs To their associated string
        /// values. The dictionary is stored as the property with ID 0. The codepage
        /// for the strings in the dictionary is defined by property with ID 1.
        /// </summary>
        /// <value>the dictionary or null
        ///  if the section does not have
        /// a dictionary.</value>
        public virtual IDictionary Dictionary
        {
            get {
                if (dictionary == null)
                    dictionary = new Hashtable();
                return dictionary;
            }
            set { 
                dictionary = value; 
            }
        }



        /// <summary>
        /// Gets the section's codepage, if any.
        /// </summary>
        /// <value>The section's codepage if one is defined, else -1.</value>
        public int Codepage
        {
            get
            {
                if (GetProperty(PropertyIDMap.PID_CODEPAGE) == null)
                    return -1;
                int codepage =
                    (int)GetProperty(PropertyIDMap.PID_CODEPAGE);
                return codepage;
            }
        }

    }
}