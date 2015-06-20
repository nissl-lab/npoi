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

    /// <summary>
    /// A property in a {@link Section} of a {@link PropertySet}.
    /// The property's ID gives the property a meaning
    /// in the context of its {@link Section}. Each {@link Section} spans
    /// its own name space of property IDs.
    /// The property's type determines how its
    /// value  is interpreted. For example, if the type Is
    /// {@link Variant#VT_LPSTR} (byte string), the value consists of a
    /// DWord telling how many bytes the string Contains. The bytes follow
    /// immediately, including any null bytes that terminate the
    /// string. The type {@link Variant#VT_I4} denotes a four-byte integer
    /// value, {@link Variant#VT_FILETIME} some DateTime and time (of a
    /// file).
    /// Please note that not all {@link Variant} types yet. This might Change
    /// over time but largely depends on your feedback so that the POI team knows
    /// which variant types are really needed. So please feel free To submit error
    /// reports or patches for the types you need.
    /// Microsoft documentation: 
    /// <a href="http://msdn.microsoft.com/library/en-us/stg/stg/property_Set_display_name_dictionary.asp?frame=true">
    /// Property Set Display Name Dictionary</a>
    /// .
    /// @author Rainer Klute 
    /// <a href="mailto:klute@rainer-klute.de">&lt;klute@rainer-klute.de&gt;</a>
    /// @author Drew Varner (Drew.Varner InAndAround sc.edu)
    /// @see Section
    /// @see Variant
    /// @since 2002-02-09
    /// </summary>
    public class Property
    {

        /** The property's ID. */
        protected long id;


        /**
         * Returns the property's ID.
         *
         * @return The ID value
         */
        public virtual long ID
        {
            get { return id; }
            set { id = value; }
        }



        /** The property's type. */
        protected long type;


        /**
         * Returns the property's type.
         *
         * @return The type value
         */
        public virtual long Type
        {
            get { return type; }
            set { type = value; }
        }



        /** The property's value. */
        protected Object value;


        /// <summary>
        /// Gets the property's value.
        /// </summary>
        /// <value>The property's value</value>
        public virtual object Value
        {
            get { return this.value; }
            set { this.value = value; }
        }



        /// <summary>
        /// Initializes a new instance of the <see cref="Property"/> class.
        /// </summary>
        /// <param name="id">the property's ID.</param>
        /// <param name="type">the property's type, see {@link Variant}.</param>
        /// <param name="value">the property's value. Only certain types are allowed, see
        /// {@link Variant}.</param>
        public Property(long id, long type, Object value)
        {
            this.id = id;
            this.type = type;
            this.value = value;
        }



        /// <summary>
        /// Initializes a new instance of the <see cref="Property"/> class.
        /// </summary>
        /// <param name="id">The property's ID.</param>
        /// <param name="src">The bytes the property Set stream consists of.</param>
        /// <param name="offset">The property's type/value pair's offset in the
        /// section.</param>
        /// <param name="Length">The property's type/value pair's Length in bytes.</param>
        /// <param name="codepage">The section's and thus the property's
        /// codepage. It is needed only when Reading string values</param>
        public Property(long id, byte[] src, long offset,
                        int Length, int codepage)
        {
            this.id = id;

            /*
             * ID 0 is a special case since it specifies a dictionary of
             * property IDs and property names.
             */
            if (id == 0)
            {
                value = ReadDictionary(src, offset, Length, codepage);
                return;
            }

            int o = (int)offset;
            type = LittleEndian.GetUInt(src, o);
            o += LittleEndianConsts.INT_SIZE;

            try
            {
                value = VariantSupport.Read(src, o, Length, (int)type, codepage);
            }
            catch (UnsupportedVariantTypeException ex)
            {
                VariantSupport.WriteUnsupportedTypeMessage(ex);
                value = ex.Value;
            }
        }



        /// <summary>
        /// Initializes a new instance of the <see cref="Property"/> class.
        /// </summary>
        protected Property()
        { }



        /// <summary>
        /// Reads the dictionary.
        /// </summary>
        /// <param name="src">The byte array containing the bytes making out the dictionary.</param>
        /// <param name="offset">At this offset within src the dictionary starts.</param>
        /// <param name="Length">The dictionary Contains at most this many bytes.</param>
        /// <param name="codepage">The codepage of the string values.</param>
        /// <returns>The dictonary</returns>
        protected IDictionary ReadDictionary(byte[] src, long offset,
                                     int Length, int codepage)
        {
            /* Check whether "offset" points into the "src" array". */
            if (offset < 0 || offset > src.Length)
                throw new HPSFRuntimeException
                    ("Illegal offset " + offset + " while HPSF stream Contains " +
                     Length + " bytes.");
            int o = (int)offset;

            /*
             * Read the number of dictionary entries.
             */
            long nrEntries = LittleEndian.GetUInt(src, o);
            o += LittleEndianConsts.INT_SIZE;

            Hashtable m = new Hashtable((int)nrEntries, (float)1.0);

            try
            {
                for (int i = 0; i < nrEntries; i++)
                {
                    /* The key. */
                    long id = LittleEndian.GetUInt(src, o);
                    o += LittleEndianConsts.INT_SIZE;

                    /* The value (a string). The Length is the either the
                     * number of (two-byte) characters if the character Set is Unicode
                     * or the number of bytes if the character Set is not Unicode.
                     * The Length includes terminating 0x00 bytes which we have To strip
                     * off To Create a Java string. */
                    long sLength = LittleEndian.GetUInt(src, o);
                    o += LittleEndianConsts.INT_SIZE;

                    /* Read the string. */
                    StringBuilder b = new StringBuilder();
                    switch (codepage)
                    {
                        case -1:
                            {
                                /* Without a codepage the Length is equal To the number of
                                 * bytes. */
                                b.Append(Encoding.UTF8.GetString(src, o, (int)sLength));
                                break;
                            }
                        case CodePageUtil.CP_UNICODE:
                            {
                                /* The Length is the number of characters, i.e. the number
                                 * of bytes is twice the number of the characters. */
                                int nrBytes = (int)(sLength * 2);
                                byte[] h = new byte[nrBytes];
                                for (int i2 = 0; i2 < nrBytes; i2++)
                                {
                                    h[i2] = src[o + i2];
                                }
                                b.Append(Encoding.GetEncoding(codepage).GetString(h, 0, nrBytes-2));
                                break;
                            }
                        default:
                            {
                                /* For encodings other than Unicode the Length is the number
                                 * of bytes. */
                                b.Append(Encoding.GetEncoding(codepage).GetString(src, o, (int)sLength));
                                break;
                            }
                    }

                    /* Strip 0x00 characters from the end of the string: */
                    while (b.Length > 0 && b[b.Length - 1] == 0x00)
                        b.Length = b.Length - 1;
                    if (codepage == CodePageUtil.CP_UNICODE)
                    {
                        if (sLength % 2 == 1)
                            sLength++;
                        o += (int)(sLength + sLength);
                    }
                    else
                        o += (int)sLength;
                    m[id] = b.ToString();
                }
            }
            catch (Exception ex)
            {
                POILogger l = POILogFactory.GetLogger(typeof(Property));
                l.Log(POILogger.WARN,
                        "The property Set's dictionary Contains bogus data. "
                        + "All dictionary entries starting with the one with ID "
                        + id + " will be ignored.", ex);
            }
            return m;
        }



        /// <summary>
        /// Gets the property's size in bytes. This is always a multiple of
        /// 4.
        /// </summary>
        /// <value>the property's size in bytes</value>
        public int Count
        {
            get
            {
                int Length = VariantSupport.GetVariantLength(type);
                if (Length >= 0)
                    return Length; /* Fixed Length */
                if (Length == -2)
                    /* Unknown Length */
                    throw new WritingNotSupportedException(type, null);

                /* Variable Length: */
                int PAddING = 4; /* Pad To multiples of 4. */
                switch ((int)type)
                {
                    case Variant.VT_LPSTR:
                        {
                            int l = ((String)value).Length + 1;
                            int r = l % PAddING;
                            if (r > 0)
                                l += PAddING - r;
                            Length += l;
                            break;
                        }
                    case Variant.VT_EMPTY:
                        break;
                    default:
                        throw new WritingNotSupportedException(type, value);
                }
                return Length;
            }
        }



        /// <summary>
        /// Compares two properties.
        /// Please beware that a property with
        /// ID == 0 is a special case: It does not have a type, and its value is the
        /// section's dictionary. Another special case are strings: Two properties
        /// may have the different types Variant.VT_LPSTR and Variant.VT_LPWSTR;
        /// </summary>
        /// <param name="o">The o.</param>
        /// <returns></returns>
        public override bool Equals(Object o)
        {
            if (!(o is Property))
                return false;
            Property p = (Property)o;
            Object pValue = p.Value;
            long pId = p.ID;
            if (id != pId || (id != 0 && !TypesAreEqual(type, p.Type)))
                return false;
            if (value == null && pValue == null)
                return true;
            if (value == null || pValue == null)
                return false;

            /* It's clear now that both values are non-null. */
            Type valueClass = value.GetType();
            Type pValueClass = pValue.GetType();
            if (!(valueClass.IsAssignableFrom(pValueClass)) &&
                !(pValueClass.IsAssignableFrom(valueClass)))
                return false;

            if (value is byte[])
                return Arrays.Equals((byte[])value, (byte[])pValue);

            return value.Equals(pValue);
        }



        /// <summary>
        /// Typeses the are equal.
        /// </summary>
        /// <param name="t1">The t1.</param>
        /// <param name="t2">The t2.</param>
        /// <returns></returns>
        private bool TypesAreEqual(long t1, long t2)
        {
            if (t1 == t2 ||
                (t1 == Variant.VT_LPSTR && t2 == Variant.VT_LPWSTR) ||
                (t2 == Variant.VT_LPSTR && t1 == Variant.VT_LPWSTR))
                return true;
            else
                return false;
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
            GetHashCode += id;
            GetHashCode += type;
            if (value != null)
                GetHashCode += value.GetHashCode();
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
            b.Append(this.GetType().Name);
            b.Append('[');
            b.Append("id: ");
            b.Append(ID);
            b.Append(", type: ");
            b.Append(GetType());
            Object value = Value;
            b.Append(", value: ");
            if (value is String)
            {
                b.Append(value.ToString());
                String s = value.ToString();
                int l = s.Length;

                byte[] bytes = new byte[l*2];
                for (int i = 0; i < l; i++)
                {
                    char c = s[i];
                    byte high = (byte) ((c & 0x00ff00) >> 8);
                    byte low = (byte) ((c & 0x0000ff) >> 0);
                    bytes[i*2] = high;
                    bytes[i*2 + 1] = low;
                }
                b.Append(" [");
                if (bytes.Length > 0)
                {
                    String hex = HexDump.Dump(bytes, 0L, 0);
                    b.Append(hex);
                }
                b.Append("]");
            }
            else if (value is byte[])
            {
                byte[] bytes = (byte[]) value;
                if (bytes.Length > 0)
                {
                    String hex = HexDump.Dump(bytes, 0L, 0);
                    b.Append(hex);
                }
            }
            else
            {
                b.Append(value.ToString());
            }
            b.Append(']');
            return b.ToString();
        }

    }
}