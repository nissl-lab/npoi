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
    using System.Collections;
    using NPOI.HPSF.Wellknown;
    using System.Text;


    /// <summary>
    /// Maintains the instances of {@link CustomProperty} that belong To a
    /// {@link DocumentSummaryInformation}. The class maintains the names of the
    /// custom properties in a dictionary. It implements the {@link Map} interface
    /// and by this provides a simplified view on custom properties: A property's
    /// name is the key that maps To a typed value. This implementation hides
    /// property IDs from the developer and regards the property names as keys To
    /// typed values.
    /// While this class provides a simple API To custom properties, it ignores
    /// the fact that not names, but IDs are the real keys To properties. Under the
    /// hood this class maintains a 1:1 relationship between IDs and names. Therefore
    /// you should not use this class To process property Sets with several IDs
    /// mapping To the same name or with properties without a name: the result will
    /// contain only a subSet of the original properties. If you really need To deal
    /// such property Sets, use HPSF's low-level access methods.
    /// An application can call the {@link #isPure} method To check whether a
    /// property Set parsed by {@link CustomProperties} is still pure (i.e.
    /// unmodified) or whether one or more properties have been dropped.
    /// This class is not thRead-safe; concurrent access To instances of this
    /// class must be syncronized.
    /// @author Rainer Klute 
    /// <a href="mailto:klute@rainer-klute.de">&lt;klute@rainer-klute.de&gt;</a>
    /// @since 2006-02-09
    /// </summary>
    public class CustomProperties : Hashtable
    {

        /**
         * Maps property IDs To property names.
         */
        private Hashtable dictionaryIDToName = new Hashtable();

        /**
         * Maps property names To property IDs.
         */
        private Hashtable dictionaryNameToID = new Hashtable();

        /**
         * Tells whether this object is pure or not.
         */
        private bool isPure = true;



        /// <summary>
        /// Puts a {@link CustomProperty} into this map. It is assumed that the
        /// {@link CustomProperty} alReady has a valid ID. Otherwise use
        /// {@link #Put(CustomProperty)}.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="cp">The custom property.</param>
        /// <returns></returns>
        public CustomProperty Put(string name, CustomProperty cp)
        {
            if (string.IsNullOrEmpty((string)name))     //tony qu changed the code
            {
                /* Ignoring a property without a name. */
                isPure = false;
                return null;
            }
            if (!(name is String))
                throw new ArgumentException("The name of a custom property must " +
                        "be a String, but it is a " +
                        name.GetType().Name);
            if (!(name.Equals(cp.Name)))
                throw new ArgumentException("Parameter \"name\" (" + name +
                        ") and custom property's name (" + cp.Name +
                        ") do not match.");

            /* Register name and ID in the dictionary. Mapping in both directions is possible. If there is alReady a  */
            long idKey = cp.ID;
            Object oldID = dictionaryNameToID[name];
            if(oldID!=null)
                dictionaryIDToName.Remove(oldID);
            dictionaryNameToID[name]=idKey;
            dictionaryIDToName[idKey]= name;

            /* Put the custom property into this map. */
            if (oldID != null)
                base.Remove(oldID);

            base[idKey]= cp;
            return cp;
        }

        /**
     * Returns a set of all the names of our
     *  custom properties. Equivalent to 
     *  {@link #nameSet()}
     */
        public ICollection KeySet()
        {
            return dictionaryNameToID.Keys;
        }

        /**
         * Returns a set of all the names of our
         *  custom properties
         */
        public ICollection NameSet()
        {
            return dictionaryNameToID.Keys;
        }

        /**
         * Returns a set of all the IDs of our
         *  custom properties
         */
        public ICollection IdSet()
        {
            return dictionaryNameToID.Keys;
        }

        /// <summary>
        /// Puts a {@link CustomProperty} that has not yet a valid ID into this
        /// map. The method will allocate a suitable ID for the custom property:
        /// <ul>
        /// 	<li>If there is alReady a property with the same name, take the ID
        /// of that property.</li>
        /// 	<li>Otherwise Find the highest ID and use its value plus one.</li>
        /// </ul>
        /// </summary>
        /// <param name="customProperty">The custom property.</param>
        /// <returns>If the was alReady a property with the same name, the</returns>
        private Object Put(CustomProperty customProperty)
        {
            String name = customProperty.Name;
            
            /* Check whether a property with this name is in the map alReady. */
            object oldId = dictionaryNameToID[(name)];
            if (oldId!=null)
            {
                customProperty.ID = (long)oldId;
            }
            else
            {
                long max = 1;
                for (IEnumerator i = dictionaryIDToName.Keys.GetEnumerator(); i.MoveNext(); )
                {
                    long id = (long)i.Current;
                    if (id > max)
                        max = id;
                }
                customProperty.ID = max + 1;
            }
            return this.Put(name, customProperty);
        }



        /// <summary>
        /// Removes a custom property.
        /// </summary>
        /// <param name="name">The name of the custom property To Remove</param>
        /// <returns>The Removed property or 
        /// <c>null</c>
        ///  if the specified property was not found.</returns>
        public object Remove(String name)
        {
            if (dictionaryNameToID[name] == null)
                return null;
            long id = (long)dictionaryNameToID[name];
            dictionaryIDToName.Remove(id);
            dictionaryNameToID.Remove(name);
            CustomProperty tmp = (CustomProperty)this[id];
            this.Remove(id);
            return tmp;
        }

        /// <summary>
        /// Adds a named string property.
        /// </summary>
        /// <param name="name">The property's name.</param>
        /// <param name="value">The property's value.</param>
        /// <returns>the property that was stored under the specified name before, or
        /// <c>null</c>
        ///  if there was no such property before.</returns>
        public Object Put(String name, String value)
        {
            MutableProperty p = new MutableProperty();
            p.ID=-1;
            p.Type=Variant.VT_LPWSTR;
            p.Value=value;
            CustomProperty cp = new CustomProperty(p, name);
            return Put(cp);
        }

        /// <summary>
        /// Adds a named long property
        /// </summary>
        /// <param name="name">The property's name.</param>
        /// <param name="value">The property's value.</param>
        /// <returns>the property that was stored under the specified name before, or
        /// <c>null</c>
        ///  if there was no such property before.</returns>
        public Object Put(String name, long value)
        {
            MutableProperty p = new MutableProperty();
            p.ID=-1;
            p.Type=Variant.VT_I8;
            p.Value=value;
            CustomProperty cp = new CustomProperty(p, name);
            return Put(cp);
        }

        /// <summary>
        /// Adds a named double property.
        /// </summary>
        /// <param name="name">The property's name.</param>
        /// <param name="value">The property's value.</param>
        /// <returns>the property that was stored under the specified name before, or
        /// <c>null</c>
        ///  if there was no such property before.</returns>
        public Object Put(String name, Double value)
        {
            MutableProperty p = new MutableProperty();
            p.ID=-1;
            p.Type=Variant.VT_R8;
            p.Value=value;
            CustomProperty cp = new CustomProperty(p, name);
            return Put(cp);
        }

        /// <summary>
        /// Adds a named integer property.
        /// </summary>
        /// <param name="name">The property's name.</param>
        /// <param name="value">The property's value.</param>
        /// <returns>the property that was stored under the specified name before, or
        /// <c>null</c>
        ///  if there was no such property before.</returns>
        public Object Put(String name, int value)
        {
            MutableProperty p = new MutableProperty();
            p.ID=-1;
            p.Type=Variant.VT_I4;
            p.Value=value;
            CustomProperty cp = new CustomProperty(p, name);
            return Put(cp);
        }

        /// <summary>
        /// Adds a named bool property.
        /// </summary>
        /// <param name="name">The property's name.</param>
        /// <param name="value">The property's value.</param>
        /// <returns>the property that was stored under the specified name before, or
        /// <c>null</c>
        ///  if there was no such property before.</returns>
        public Object Put(String name, bool value)
        {
            MutableProperty p = new MutableProperty();
            p.ID=-1;
            p.Type=Variant.VT_BOOL;
            p.Value=value;
            CustomProperty cp = new CustomProperty(p, name);
            return Put(cp);
        }


        /// <summary>
        /// Adds a named date property.
        /// </summary>
        /// <param name="name">The property's name.</param>
        /// <param name="value">The property's value.</param>
        /// <returns>the property that was stored under the specified name before, or
        /// <c>null</c>
        ///  if there was no such property before.</returns>
        public Object Put(String name,DateTime value)
        {
            MutableProperty p = new MutableProperty();
            p.ID=-1;
            p.Type=Variant.VT_FILETIME;
            p.Value=value;
            CustomProperty cp = new CustomProperty(p, name);
            return Put(cp);
        }

        /// <summary>
        /// Gets the <see cref="System.Object"/> with the specified name.
        /// </summary>
        /// <value>the value or 
        /// <c>null</c>
        ///  if a value with the specified
        /// name is not found in the custom properties.</value>
        public Object this[string name]
        {
            get
            {
                object x = dictionaryNameToID[name];
                //string.Equals seems not support Unicode string
                if (x == null)
                {
                    IEnumerator dic = dictionaryNameToID.GetEnumerator();
                    while (dic.MoveNext())
                    {
                        string key = ((DictionaryEntry)dic.Current).Key as string;

                        int codepage = this.Codepage;
                        if (codepage < 0)
                            codepage = (int)Constants.CP_UNICODE;
                        byte[] a= Encoding.GetEncoding(codepage).GetBytes(key);
                        byte[] b = Encoding.UTF8.GetBytes(name);
                        if (NPOI.Util.Arrays.Equals(a, b))
                            x = ((DictionaryEntry)dic.Current).Value;
                    }
                    if (x == null)
                    {
                        return null;
                    }
                }
                long id = (long)x;
                CustomProperty cp = (CustomProperty)base[id];
                return cp != null ? cp.Value : null;
            }
        }
        /**
     * Checks against both String Name and Long ID
     */
        public override bool ContainsKey(Object key)
        {
            if (key is long)
            {
                return base.ContainsKey((long)key);
            }
            if (key is String)
            {
                return base.ContainsKey((long)dictionaryNameToID[(key)]);
            }
            return false;
        }

        /**
         * Checks against both the property, and its values. 
         */
        public override bool ContainsValue(Object value)
        {
            if (value is CustomProperty)
            {
                return base.ContainsValue(value);
            }
            else
            {
                foreach (object cp in base.Values)
                {
                    if ((cp as CustomProperty).Value == value)
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Gets the dictionary which Contains IDs and names of the named custom
        /// properties.
        /// </summary>
        /// <value>The dictionary.</value>
        public IDictionary Dictionary
        {
            get { return dictionaryIDToName; }
        }


        /// <summary>
        /// Gets or sets the codepage.
        /// </summary>
        /// <value>The codepage.</value>
        public int Codepage
        {
            get
            {
                int codepage = -1;
                for (IEnumerator i = this.Values.GetEnumerator(); codepage == -1 && i.MoveNext(); )
                {
                    CustomProperty cp = (CustomProperty)i.Current;
                    if (cp.ID == PropertyIDMap.PID_CODEPAGE)
                        codepage = (int)cp.Value;
                }
                return codepage;
            }
            set 
            {
                MutableProperty p = new MutableProperty();
                p.ID=PropertyIDMap.PID_CODEPAGE;
                p.Type=Variant.VT_I2;
                p.Value=value;
                Put(new CustomProperty(p));
            }
        }



        /// <summary>
        /// Tells whether this {@link CustomProperties} instance is pure or one or
        /// more properties of the underlying low-level property Set has been
        /// dropped.
        /// </summary>
        /// <value><c>true</c> if this instance is pure; otherwise, <c>false</c>.</value>
        public bool IsPure
        {
            get { return isPure; }
            set { this.isPure = value; }
        }

    }
}