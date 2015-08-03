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

using System.Collections.Generic;

namespace NPOI.HPSF
{
    using System;
    using System.IO;
    using NPOI.HPSF.Wellknown;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;
    using System.Text;

    /// <summary>
    /// Abstract superclass for the convenience classes {@link
    /// SummaryInformation} and {@link DocumentSummaryInformation}.
    /// The motivation behind this class is quite nasty if you look
    /// behind the scenes, but it serves the application programmer well by
    /// providing him with the easy-to-use {@link SummaryInformation} and
    /// {@link DocumentSummaryInformation} classes. When parsing the data a
    /// property Set stream consists of (possibly coming from an {@link
    /// java.io.Stream}) we want To Read and process each byte only
    /// once. Since we don't know in advance which kind of property Set we
    /// have, we can expect only the most general {@link
    /// PropertySet}. Creating a special subclass should be as easy as
    /// calling the special subclass' constructor and pass the general
    /// {@link PropertySet} in. To make things easy internally, the special
    /// class just holds a reference To the general {@link PropertySet} and
    /// delegates all method calls To it.
    /// A cleaner implementation would have been like this: The {@link
    /// PropertySetFactory} parses the stream data into some internal
    /// object first.  Then it Finds out whether the stream is a {@link
    /// SummaryInformation}, a {@link DocumentSummaryInformation} or a
    /// general {@link PropertySet}.  However, the current implementation
    /// went the other way round historically: the convenience classes came
    /// only late To my mind.
    /// @author Rainer Klute 
    /// klute@rainer-klute.de
    /// @since 2002-02-09
    /// </summary>
    [Serializable]
    public abstract class SpecialPropertySet : MutablePropertySet
    {
        /**
     * The id to name mapping of the properties
     *  in this set.
     */
        public abstract PropertyIDMap PropertySetIDMap{get;}

        /**
         * The "real" property Set <c>SpecialPropertySet</c>
         * delegates To.
         */
        private MutablePropertySet delegate1;



        /// <summary>
        /// Initializes a new instance of the <see cref="SpecialPropertySet"/> class.
        /// </summary>
        /// <param name="ps">The property Set To be encapsulated by the <c>SpecialPropertySet</c></param>
        public SpecialPropertySet(PropertySet ps)
        {
            delegate1 = new MutablePropertySet(ps);
        }



        /// <summary>
        /// Initializes a new instance of the <see cref="SpecialPropertySet"/> class.
        /// </summary>
        /// <param name="ps">The mutable property Set To be encapsulated by the <c>SpecialPropertySet</c></param>
        public SpecialPropertySet(MutablePropertySet ps)
        {
            delegate1 = ps;
        }


        /// <summary>
        /// gets or sets the "byteOrder" property.
        /// </summary>
        /// <value>the byteOrder value To Set</value>
        public override int ByteOrder
        {
            get { return delegate1.ByteOrder; }
            set { delegate1.ByteOrder = value; }
        }


        /// <summary>
        /// gets or sets the "format" property
        /// </summary>
        /// <value>the format value To Set</value>
        public override int Format
        {
            get { return delegate1.Format; }
            set { delegate1.Format = value; }
        }

        /// <summary>
        /// gets or sets the property Set stream's low-level "class ID"
        /// field.
        /// </summary>
        /// <value>The property Set stream's low-level "class ID" field</value>
        public override ClassID ClassID
        {
            get { return delegate1.ClassID; }
            set { delegate1.ClassID = value; }
        }


        /// <summary>
        /// Returns the number of {@link Section}s in the property
        /// Set.
        /// </summary>
        /// <value>The number of {@link Section}s in the property Set.</value>
        public override int SectionCount
        {
            get { return delegate1.SectionCount; }
        }


        public override List<Section> Sections
        {
            get { return delegate1.Sections; }
        }


        /// <summary>
        /// Checks whether this {@link PropertySet} represents a Summary
        /// Information.
        /// </summary>
        /// <value>
        /// 	<c>true</c> Checks whether this {@link PropertySet} represents a Summary
        /// Information; otherwise, <c>false</c>.
        /// </value>
        public override bool IsSummaryInformation
        {
            get{return delegate1.IsSummaryInformation;}
        }

        public override Stream ToInputStream() 
        {
            return delegate1.ToInputStream();
        }

        /// <summary>
        /// Gets a value indicating whether this instance is document summary information.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is document summary information; otherwise, <c>false</c>.
        /// </value>
        /// Checks whether this {@link PropertySet} is a Document
        /// Summary Information.
        /// @return
        /// <c>true</c>
        /// if this {@link PropertySet}
        /// represents a Document Summary Information, else
        /// <c>false</c>
        public override bool IsDocumentSummaryInformation
        {
            get{return delegate1.IsDocumentSummaryInformation;}
        }


        /// <summary>
        /// Gets the PropertySet's first section.
        /// </summary>
        /// <value>The {@link PropertySet}'s first section.</value>
        public override Section FirstSection
        {
            get { return delegate1.FirstSection; }
        }

        /// <summary>
        /// Adds a section To this property set.
        /// </summary>
        /// <param name="section">The {@link Section} To Add. It will be Appended
        /// after any sections that are alReady present in the property Set
        /// and thus become the last section.</param>
        public override void AddSection(Section section)
        {
            delegate1.AddSection(section);
        }


        /// <summary>
        /// Removes all sections from this property Set.
        /// </summary>
        public override void ClearSections()
        {
            delegate1.ClearSections();
        }


        /// <summary>
        /// gets or sets the "osVersion" property
        /// </summary>
        /// <value> the osVersion value To Set</value>
        public override int OSVersion
        {
            set { delegate1.OSVersion=value; }
            get { return delegate1.OSVersion; }
        }

        /// <summary>
        /// Writes a property Set To a document in a POI filesystem directory.
        /// </summary>
        /// <param name="dir">The directory in the POI filesystem To Write the document To</param>
        /// <param name="name">The document's name. If there is alReady a document with the
        /// same name in the directory the latter will be overwritten.</param>
        public override void Write(DirectoryEntry dir, String name)
        {
            delegate1.Write(dir, name);
        }



        /// <summary>
        /// Writes the property Set To an output stream.
        /// </summary>
        /// <param name="out1">the output stream To Write the section To</param>
        public override void Write(Stream out1)
        {
            delegate1.Write(out1);
        }



        /// <summary>
        /// Returns <c>true</c> if the <c>PropertySet</c> is equal
        /// To the specified parameter, else <c>false</c>.
        /// </summary>
        /// <param name="o">the object To Compare this
        /// <c>PropertySet</c>
        /// with</param>
        /// <returns>
        /// 	<c>true</c>
        /// if the objects are equal,
        /// <c>false</c>
        /// if not
        /// </returns>
        public override bool Equals(Object o)
        {
            return delegate1.Equals(o);
        }




        /// <summary>
        /// Convenience method returning the {@link Property} array
        /// contained in this property Set. It is a shortcut for Getting
        /// the {@link PropertySet}'s {@link Section}s list and then
        /// Getting the {@link Property} array from the first {@link
        /// Section}.
        /// </summary>
        /// <value>
        /// The properties of the only {@link Section} of this
        /// {@link PropertySet}.
        /// </value>
        public override Property[] Properties
        {
            get { return delegate1.Properties; }
        }



        /// <summary>
        /// Convenience method returning the value of the property with
        /// the specified ID. If the property is not available,
        /// <c>null</c> is returned and a subsequent call To {@link
        /// #WasNull} will return <c>true</c> .
        /// </summary>
        /// <param name="id">The property ID</param>
        /// <returns>The property value</returns>
        public override Object GetProperty(int id)
        {
            return delegate1.GetProperty(id);
        }


        /// <summary>
        /// Convenience method returning the value of a bool property
        /// with the specified ID. If the property is not available,
        /// <c>false</c> is returned. A subsequent call To {@link
        /// #WasNull} will return <c>true</c> To let the caller
        /// distinguish that case from a real property value of
        /// <c>false</c>.
        /// </summary>
        /// <param name="id">The property ID</param>
        /// <returns>The property value</returns>
        public override bool GetPropertyBooleanValue(int id)
        {
            return delegate1.GetPropertyBooleanValue(id);
        }



        /// <summary>
        /// Convenience method returning the value of the numeric
        /// property with the specified ID. If the property is not
        /// available, 0 is returned. A subsequent call To {@link #WasNull}
        /// will return <c>true</c> To let the caller distinguish
        /// that case from a real property value of 0.
        /// </summary>
        /// <param name="id">The property ID</param>
        /// <returns>The propertyIntValue value</returns>
        public override int GetPropertyIntValue(int id)
        {
            return delegate1.GetPropertyIntValue(id);
        }

        /**
         * Fetches the property with the given ID, then does its
         *  best to return it as a String
         * @return The property as a String, or null if unavailable
         */
        protected String GetPropertyStringValue(int propertyId)
        {
            Object propertyValue = GetProperty(propertyId);
            return GetPropertyStringValue(propertyValue);
        }
        protected static String GetPropertyStringValue(Object propertyValue) 
        {
            // Normal cases
            if (propertyValue == null) return null;
            if (propertyValue is String) return (String)propertyValue;

            // Do our best with some edge cases
            if (propertyValue is byte[])
            {
                byte[] b = (byte[])propertyValue;
                if (b.Length == 0)
                {
                    return "";
                }
                if (b.Length == 1)
                {
                    return b[0].ToString();
                }
                if (b.Length == 2)
                {
                    return LittleEndian.GetUShort(b).ToString();
                }
                if (b.Length == 4)
                {
                    return LittleEndian.GetUInt(b).ToString();
                }
                // Maybe it's a string? who knows!
                return Encoding.UTF8.GetString(b);
            }
            return propertyValue.ToString();
        }

        /// <summary>
        /// Serves as a hash function for a particular type.
        /// </summary>
        /// <returns>
        /// A hash code for the current <see cref="T:System.Object"/>.
        /// </returns>
        public override int GetHashCode()
        {
            return delegate1.GetHashCode();
        }



        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override String ToString()
        {
            return delegate1.ToString();
        }



        /// <summary>
        /// Checks whether the property which the last call To {@link
        /// #GetPropertyIntValue} or {@link #GetProperty} tried To access
        /// Was available or not. This information might be important for
        /// callers of {@link #GetPropertyIntValue} since the latter
        /// returns 0 if the property does not exist. Using {@link
        /// #WasNull}, the caller can distiguish this case from a
        /// property's real value of 0.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if the last call To {@link
        /// #GetPropertyIntValue} or {@link #GetProperty} tried To access a
        /// property that Was not available; otherwise, <c>false</c>.
        /// </value>
        public override bool WasNull
        {
            get { return delegate1.WasNull; }
        }

    }
}