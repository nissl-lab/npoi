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

namespace NPOI.HPSF
{
    using System;
    using System.IO;
    using System.Text;
    using System.Collections;
    using NPOI.Util;
    using NPOI.HPSF.Wellknown;
    using System.Globalization;

    /// <summary>
    /// Adds writing capability To the {@link Section} class.
    /// Please be aware that this class' functionality will be merged into the
    /// {@link Section} class at a later time, so the API will Change.
    /// @since 2002-02-20
    /// </summary>
    public class MutableSection : Section
    {
        /**
         * If the "dirty" flag is true, the section's size must be
         * (re-)calculated before the section is written.
         */
        private bool dirty = true;



        /**
         * List To assemble the properties. Unfortunately a wrong
         * decision has been taken when specifying the "properties" field
         * as an Property[]. It should have been a {@link java.util.List}.
         */
        private ArrayList preprops;



        /**
         * Contains the bytes making out the section. This byte array is
         * established when the section's size is calculated and can be reused
         * later. It is valid only if the "dirty" flag is false.
         */
        private byte[] sectionBytes;



        /// <summary>
        /// Initializes a new instance of the <see cref="MutableSection"/> class.
        /// </summary>
        public MutableSection()
        {
            dirty = true;
            formatID = null;
            offset = -1;
            preprops = new ArrayList();
        }



        /// <summary>
        /// Constructs a <c>MutableSection</c> by doing a deep copy of an
        /// existing <c>Section</c>. All nested <c>Property</c>
        /// instances, will be their mutable counterparts in the new
        /// <c>MutableSection</c>.
        /// </summary>
        /// <param name="s">The section Set To copy</param>
        public MutableSection(Section s)
        {
            SetFormatID(s.FormatID);
            Property[] pa = s.Properties;
            MutableProperty[] mpa = new MutableProperty[pa.Length];
            for (int i = 0; i < pa.Length; i++)
                mpa[i] = new MutableProperty(pa[i]);
            SetProperties(mpa);
            this.Dictionary=(s.Dictionary);
        }



        /// <summary>
        /// Sets the section's format ID.
        /// </summary>
        /// <param name="formatID">The section's format ID</param>
        public void SetFormatID(ClassID formatID)
        {
            this.formatID = formatID;
        }



        /// <summary>
        /// Sets the section's format ID.
        /// </summary>
        /// <param name="formatID">The section's format ID as a byte array. It components
        /// are in big-endian format.</param>
        public void SetFormatID(byte[] formatID)
        {
            ClassID fid = this.FormatID;
            if (fid == null)
            {
                fid = new ClassID();
                SetFormatID(fid);
            }
            fid.Bytes=formatID;
        }



        /// <summary>
        /// Sets this section's properties. Any former values are overwritten.
        /// </summary>
        /// <param name="properties">This section's new properties.</param>
        public void SetProperties(Property[] properties)
        {
            this.properties = properties;
            preprops = new ArrayList();
            for (int i = 0; i < properties.Length; i++)
                preprops.Add(properties[i]);
            dirty = true;
        }



        /// <summary>
        /// Sets the string value of the property with the specified ID.
        /// </summary>
        /// <param name="id">The property's ID</param>
        /// <param name="value">The property's value. It will be written as a Unicode
        /// string.</param>
        public void SetProperty(int id, String value)
        {
            SetProperty(id, Variant.VT_LPWSTR, value);
            dirty = true;
        }



        /// <summary>
        /// Sets the int value of the property with the specified ID.
        /// </summary>
        /// <param name="id">The property's ID</param>
        /// <param name="value">The property's value.</param>
        public void SetProperty(int id, int value)
        {
            SetProperty(id, Variant.VT_I4, value);
            dirty = true;
        }



        /// <summary>
        /// Sets the long value of the property with the specified ID.
        /// </summary>
        /// <param name="id">The property's ID</param>
        /// <param name="value">The property's value.</param>
        public void SetProperty(int id, long value)
        {
            SetProperty(id, Variant.VT_I8, value);
            dirty = true;
        }



        /// <summary>
        /// Sets the bool value of the property with the specified ID.
        /// </summary>
        /// <param name="id">The property's ID</param>
        /// <param name="value">The property's value.</param>
        public void SetProperty(int id, bool value)
        {
            SetProperty(id, Variant.VT_BOOL, value);
            dirty = true;
        }



        /// <summary>
        /// Sets the value and the variant type of the property with the
        /// specified ID. If a property with this ID is not yet present in
        /// the section, it will be Added. An alReady present property with
        /// the specified ID will be overwritten. A default mapping will be
        /// used To choose the property's type.
        /// </summary>
        /// <param name="id">The property's ID.</param>
        /// <param name="variantType">The property's variant type.</param>
        /// <param name="value">The property's value.</param>
        public void SetProperty(int id, long variantType,
                                Object value)
        {
            MutableProperty p = new MutableProperty();
            p.ID=id;
            p.Type=variantType;
            p.Value=value;
            SetProperty(p);
            dirty = true;
        }



        /// <summary>
        /// Sets the property.
        /// </summary>
        /// <param name="p">The property To be Set.</param>
        public void SetProperty(Property p)
        {
            long id = p.ID;
            RemoveProperty(id);
            preprops.Add(p);
            dirty = true;
        }



        /// <summary>
        /// Removes the property.
        /// </summary>
        /// <param name="id">The ID of the property To be Removed</param>
        public void RemoveProperty(long id)
        {
            for (IEnumerator i = preprops.GetEnumerator(); i.MoveNext(); )
                if (((Property)i.Current).ID == id)
                {
                    preprops.Remove(i.Current);
                    break;
                }
            dirty = true;
        }



        /// <summary>
        /// Sets the value of the bool property with the specified
        /// ID.
        /// </summary>
        /// <param name="id">The property's ID</param>
        /// <param name="value">The property's value</param>
        protected void SetPropertyBooleanValue(int id, bool value)
        {
            SetProperty(id, Variant.VT_BOOL, value);
        }



        /// <summary>
        /// Returns the section's size in bytes.
        /// </summary>
        /// <value>The section's size in bytes.</value>
        public override int Size
        {
            get
            {
                if (dirty)
                {
                    try
                    {
                        size = CalcSize();
                        dirty = false;
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                }
                return size;
            }
        }



        /// <summary>
        /// Calculates the section's size. It is the sum of the Lengths of the
        /// section's header (8), the properties list (16 times the number of
        /// properties) and the properties themselves.
        /// </summary>
        /// <returns>the section's Length in bytes.</returns>
        private int CalcSize()
        {
            using (MemoryStream out1 = new MemoryStream())
            {
                Write(out1);
                /* Pad To multiple of 4 bytes so that even the Windows shell (explorer)
                 * shows custom properties. */
                sectionBytes = Util.Pad4(out1.ToArray());
                return sectionBytes.Length;
            }
        }


        private class PropertyComparer : IComparer
        {
            #region IComparer Members

            int IComparer.Compare(object o1, object o2)
            {
                Property p1 = (Property)o1;
                Property p2 = (Property)o2;
                if (p1.ID < p2.ID)
                    return -1;
                else if (p1.ID == p2.ID)
                    return 0;
                else
                    return 1;
            }

            #endregion
        }
        /// <summary>
        /// Writes this section into an output stream.
        /// Internally this is done by writing into three byte array output
        /// streams: one for the properties, one for the property list and one for
        /// the section as such. The two former are Appended To the latter when they
        /// have received all their data.
        /// </summary>
        /// <param name="out1">The stream To Write into.</param>
        /// <returns>The number of bytes written, i.e. the section's size.</returns>
        public int Write(Stream out1)
        {
            /* Check whether we have alReady generated the bytes making out the
             * section. */
            if (!dirty && sectionBytes != null)
            {
                out1.Write(sectionBytes,0,sectionBytes.Length);
                return sectionBytes.Length;
            }

            /* The properties are written To this stream. */
            using (MemoryStream propertyStream =
                new MemoryStream())
            {

                /* The property list is established here. After each property that has
                 * been written To "propertyStream", a property list entry is written To
                 * "propertyListStream". */
                using (MemoryStream propertyListStream =
                     new MemoryStream())
                {

                    /* Maintain the current position in the list. */
                    int position = 0;

                    /* Increase the position variable by the size of the property list so
                     * that it points behind the property list and To the beginning of the
                     * properties themselves. */
                    position += 2 * LittleEndianConsts.INT_SIZE +
                                PropertyCount * 2 * LittleEndianConsts.INT_SIZE;

                    /* Writing the section's dictionary it tricky. If there is a dictionary
                     * (property 0) the codepage property (property 1) must be Set, Too. */
                    int codepage = -1;
                    if (GetProperty(PropertyIDMap.PID_DICTIONARY) != null)
                    {
                        Object p1 = GetProperty(PropertyIDMap.PID_CODEPAGE);
                        if (p1 != null)
                        {
                            if (!(p1 is int))
                                throw new IllegalPropertySetDataException
                                    ("The codepage property (ID = 1) must be an " +
                                     "Integer object.");
                        }
                        else
                            /* Warning: The codepage property is not Set although a
                             * dictionary is present. In order To cope with this problem we
                             * Add the codepage property and Set it To Unicode. */
                            SetProperty(PropertyIDMap.PID_CODEPAGE, Variant.VT_I2,
                                        CodePageUtil.CP_UNICODE);
                        codepage = Codepage;
                    }



                    /* Sort the property list by their property IDs: */
                    preprops.Sort(new PropertyComparer());

                    /* Write the properties and the property list into their respective
                     * streams: */
                    for (int i = 0; i < preprops.Count; i++)
                    {
                        MutableProperty p = (MutableProperty)preprops[i];
                        long id = p.ID;

                        /* Write the property list entry. */
                        TypeWriter.WriteUIntToStream(propertyListStream, (uint)p.ID);
                        TypeWriter.WriteUIntToStream(propertyListStream, (uint)position);

                        /* If the property ID is not equal 0 we Write the property and all
                         * is fine. However, if it Equals 0 we have To Write the section's
                         * dictionary which has an implicit type only and an explicit
                         * value. */
                        if (id != 0)
                        {
                            /* Write the property and update the position To the next
                             * property. */
                            position += p.Write(propertyStream, Codepage);
                        }
                        else
                        {
                            if (codepage == -1)
                                throw new IllegalPropertySetDataException
                                    ("Codepage (property 1) is undefined.");
                            position += WriteDictionary(propertyStream, dictionary,
                                                        codepage);
                        }
                    }
                    propertyStream.Flush();
                    propertyListStream.Flush();

                    /* Write the section: */
                    byte[] pb1 = propertyListStream.ToArray();
                    byte[] pb2 = propertyStream.ToArray();

                    /* Write the section's Length: */
                    TypeWriter.WriteToStream(out1, LittleEndianConsts.INT_SIZE * 2 +
                                                  pb1.Length + pb2.Length);

                    /* Write the section's number of properties: */
                    TypeWriter.WriteToStream(out1, PropertyCount);

                    /* Write the property list: */
                    out1.Write(pb1, 0, pb1.Length);

                    /* Write the properties: */
                    out1.Write(pb2, 0, pb2.Length);

                    int streamLength = LittleEndianConsts.INT_SIZE * 2 + pb1.Length + pb2.Length;
                    return streamLength;
                }
            }
        }



        /// <summary>
        /// Writes the section's dictionary
        /// </summary>
        /// <param name="out1">The output stream To Write To.</param>
        /// <param name="dictionary">The dictionary.</param>
        /// <param name="codepage">The codepage to be used to Write the dictionary items.</param>
        /// <returns>The number of bytes written</returns>
        /// <remarks>
        /// see MSDN KB: http://msdn.microsoft.com/en-us/library/aa380065(VS.85).aspx
        /// </remarks>
        private static int WriteDictionary(Stream out1,
                                           IDictionary dictionary, int codepage)
        {
            int length = TypeWriter.WriteUIntToStream(out1, (uint)dictionary.Count);
            for (IEnumerator i = dictionary.Keys.GetEnumerator(); i.MoveNext(); )
            {
                long key = Convert.ToInt64(i.Current, CultureInfo.InvariantCulture);
                String value = (String)dictionary[key];
                //tony qu added: some key is int32 instead of int64
                if(value==null)
                    value = (String)dictionary[(int)key];

                if (codepage == CodePageUtil.CP_UNICODE)
                {
                    /* Write the dictionary item in Unicode. */
                    int sLength = value.Length + 1;
                    if (sLength % 2 == 1)
                        sLength++;
                    length += TypeWriter.WriteUIntToStream(out1, (uint)key);
                    length += TypeWriter.WriteUIntToStream(out1, (uint)sLength);
                    byte[] ca =
                        Encoding.GetEncoding(codepage).GetBytes(value);
                    for (int j =0; j < ca.Length; j++)   
                    {
                        out1.WriteByte(ca[j]);
                        length ++;
                    }
                    sLength -= value.Length;
                    while (sLength > 0)
                    {
                        out1.WriteByte(0x00);
                        out1.WriteByte(0x00);
                        length += 2;
                        sLength--;
                    }
                }
                else
                {
                    /* Write the dictionary item in another codepage than
                     * Unicode. */
                    length += TypeWriter.WriteUIntToStream(out1, (uint)key);
                    length += TypeWriter.WriteUIntToStream(out1, (uint)value.Length + 1);

                    try
                    {
                        byte[] ba =
                            Encoding.GetEncoding(codepage).GetBytes(value);
                        for (int j = 0; j < ba.Length; j++)
                        {
                            out1.WriteByte(ba[j]);
                            length++;
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new IllegalPropertySetDataException(ex);
                    }

                    out1.WriteByte(0x00);
                    length++;
                }
            }
            return length;
        }



        /// <summary>
        /// OverWrites the base class' method To cope with a redundancy:
        /// the property count is maintained in a separate member variable, but
        /// shouldn't.
        /// </summary>
        /// <value>The number of properties in this section.</value>
        public override int PropertyCount
        {
            get { return preprops.Count; }
        }



        /// <summary>
        /// Returns this section's properties.
        /// </summary>
        /// <value>This section's properties.</value>
        public override Property[] Properties
        {
            get
            {
                EnsureProperties();
                return properties;
            }
        }
        /// <summary>
        /// Ensures the properties.
        /// </summary>
        public void EnsureProperties()
        {
            properties = (Property[])preprops.ToArray(typeof(Property));
        }


        /// <summary>
        /// Gets a property.
        /// </summary>
        /// <param name="id">The ID of the property To Get</param>
        /// <returns>The property or null  if there is no such property</returns>
        public override Object GetProperty(long id)
        {
            /* Calling Properties ensures that properties and preprops are in
             * sync. */
            EnsureProperties();
            return base.GetProperty(id);
        }



        /// <summary>
        /// Sets the section's dictionary. All keys in the dictionary must be
        /// {@link java.lang.long} instances, all values must be
        /// {@link java.lang.String}s. This method overWrites the properties with IDs
        /// 0 and 1 since they are reserved for the dictionary and the dictionary's
        /// codepage. Setting these properties explicitly might have surprising
        /// effects. An application should never do this but always use this
        /// method.
        /// </summary>
        /// <value>
        /// the dictionary
        /// </value>
        public override IDictionary Dictionary
        {
            get {
                return this.dictionary;
            }
            set
            {
                if (value != null)
                {
                    for (IEnumerator i = value.Keys.GetEnumerator();
                         i.MoveNext(); )
                        if (!(i.Current is Int64 || i.Current is Int32))
                            throw new IllegalPropertySetDataException
                                ("Dictionary keys must be of type long. but it's " + i.Current + ","+i.Current.GetType().Name+" now");

                    this.dictionary = value;

                    /* Set the dictionary property (ID 0). Please note that the second
                     * parameter in the method call below is unused because dictionaries
                     * don't have a type. */
                    SetProperty(PropertyIDMap.PID_DICTIONARY, -1, value);

                    /* If the codepage property (ID 1) for the strings (keys and
                     * values) used in the dictionary is not yet defined, Set it To
                     * Unicode. */

                    if (GetProperty(PropertyIDMap.PID_CODEPAGE) == null)
                    {
                        SetProperty(PropertyIDMap.PID_CODEPAGE, Variant.VT_I2,
                                    CodePageUtil.CP_UNICODE);
                    }

                }
                else
                {
                    /* Setting the dictionary To null means To Remove property 0.
                     * However, it does not mean To Remove property 1 (codepage). */
                    RemoveProperty(PropertyIDMap.PID_DICTIONARY);
                }
            }
        }



        /// <summary>
        /// Sets the property.
        /// </summary>
        /// <param name="id">The property ID.</param>
        /// <param name="value">The property's value. The value's class must be one of those
        /// supported by HPSF.</param>
        public void SetProperty(int id, Object value)
        {
            if (value is String)
                SetProperty(id, (String)value);
            else if (value is long)
                SetProperty(id, ((long)value));
            else if (value is int)
                SetProperty(id, value);
            else if (value is short)
                SetProperty(id, (short)value);
            else if (value is bool)
                SetProperty(id, (bool)value);
            else if (value is DateTime)
                SetProperty(id, Variant.VT_FILETIME, value);
            else
                throw new HPSFRuntimeException(
                        "HPSF does not support properties of type " +
                        value.GetType().Name + ".");
        }



        /// <summary>
        /// Removes all properties from the section including 0 (dictionary) and
        /// 1 (codepage).
        /// </summary>
        public void Clear()
        {
            Property[] properties = Properties;
            for (int i = 0; i < properties.Length; i++)
            {
                Property p = properties[i];
                RemoveProperty(p.ID);
            }
        }

        /// <summary>
        /// Gets the section's codepage, if any.
        /// </summary>
        /// <value>The section's codepage if one is defined, else -1.</value>
        public new int Codepage
        {
            get { return base.Codepage; }
            set
            {
                SetProperty(PropertyIDMap.PID_CODEPAGE, Variant.VT_I2,
                    value);
            }
        }

    }
}