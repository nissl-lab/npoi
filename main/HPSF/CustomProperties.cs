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

using System.Collections.Generic;
using System.Linq;

namespace NPOI.HPSF
{
    using NPOI.HPSF.Wellknown;
    using System;

    /// <summary>
    /// <para>
    /// Maintains the instances of <see cref="CustomProperty"/> that belong to a
    /// <see cref="DocumentSummaryInformation"/>. The class maintains the names of the
    /// custom properties in a dictionary. It : the <see cref="Dictionary{TKey, TValue}"/> interface
    /// and by this provides a simplified view on custom properties: A property's
    /// name is the key that maps to a typed value. This implementation hides
    /// property IDs from the developer and regards the property names as keys to
    /// typed values.
    /// </para>
    /// <para>
    /// While this class provides a simple API to custom properties, it ignores
    /// the fact that not names, but IDs are the real keys to properties. Under the
    /// hood this class maintains a 1:1 relationship between IDs and names. Therefore
    /// you should not use this class to process property Sets with several IDs
    /// mapping to the same name or with properties without a name: the result will
    /// contain only a subset of the original properties. If you really need to deal
    /// such property Sets, use HPSF's low-level access methods.
    /// </para>
    /// <para>
    /// An application can call the <see cref="IsPure"/> method to check whether a
    /// property Set parsed by <see cref="CustomProperties"/> is still pure (i.e.
    /// unmodified) or whether one or more properties have been dropped.
    /// </para>
    /// <para>
    /// This class is not thread-safe; concurrent access to instances of this
    /// class must be synchronized.
    /// </para>
    /// <para>
    /// While this class is roughly Dictionary&lt;Long,CustomProperty&gt;, that's the
    /// internal representation. To external calls, it should appear as
    /// Dictionary&lt;String,Object&gt; mapping between Names and Custom Property Values.
    /// </para>
    /// </summary>
    public class CustomProperties : Dictionary<long, CustomProperty>
    {
        /// <summary>
        /// Maps property IDs to property names and vice versa.
        /// </summary>
        //private TreeBidiDictionary<long, string> dictionary = new TreeBidiDictionary<long, string>();

        private BidirectionalDictionary<long, string> dictionary = new();
        /// <summary>
        /// Tells whether this object is pure or not.
        /// </summary>
        private bool isPure = true;


        /// <summary>
        /// Puts a <see cref="CustomProperty"/> into this map. It is assumed that the
        /// <see cref="CustomProperty"/> already has a valid ID. Otherwise use
        /// <see cref="Put(CustomProperty)"/>.
        /// </summary>
        /// <param name="name">the property name</param>
        /// <param name="cp">the property</param>
        /// 
        /// <return>previous property stored under this name</return>
        public CustomProperty Put(String name, CustomProperty cp)
        {
            if(name == null)
            {
                /* Ignoring a property without a name. */
                isPure = false;
                return null;
            }

            if(!name.Equals(cp.Name))
            {
                throw new ArgumentException("Parameter \"name\" (" + name +
                        ") and custom property's name (" + cp.Name +
                        ") do not match.");
            }

            /* Register name and ID in the dictionary. Mapping in both directions is possible. If there is already a  */
            if(dictionary.ContainsValue(name))
            {
                long k = dictionary.GetKey(name);
                dictionary.Remove(k);
                base.Remove(k);
            }

            dictionary[cp.ID] = name;

            /* Put the custom property into this map. */
            base.Add(cp.ID, cp);
            return cp;
        }



        /// <summary>
        /// <para>
        /// Puts a <see cref="CustomProperty"/> that has not yet a valid ID into this
        /// map. The method will allocate a suitable ID for the custom property:
        /// </para>
        /// <para>
        /// If there is already a property with the same name, take the ID
        /// of that property.
        /// </para>
        /// <para>
        /// Otherwise find the highest ID and use its value plus one.
        /// </para>
        /// </summary>
        /// <param name="customProperty">customProperty</param>
        /// <return>there was already a property with the same name, the old property</return>
        /// <exception cref="ClassCastException">ClassCastException</exception>
        private object Put(CustomProperty customProperty)
        {
            string name = customProperty.Name;

            /* Check whether a property with this name is in the map already. */
            long? oldId = (name == null|| !dictionary.ContainsValue(name)) ? null : dictionary.GetKey(name);
            if(oldId != null)
            {
                customProperty.ID = (oldId.Value);
            }
            else
            {
                long lastKey = (dictionary.Count == 0) ? 0 : dictionary.LastKey();
                customProperty.ID = (Math.Max(lastKey, PropertyIDMap.PID_MAX) + 1);
            }
            return this.Put(name, customProperty);
        }



        /// <summary>
        /// Removes a custom property.
        /// </summary>
        /// <param name="name">The name of the custom property to remove</param>
        /// <return>removed property or <c>null</c> if the specified property was not found.</return>
        /// 
        /// @see java.Util.HashSet#remove(java.lang.Object)
        public Object Remove(String name)
        {
            long id = dictionary.RemoveValue(name);
            CustomProperty cp = null;
            if(base.ContainsKey(id))
                cp = this[id];

            base.Remove(id);
            return cp;
        }

        /// <summary>
        /// Adds a named string property.
        /// </summary>
        /// <param name="name">The property's name.</param>
        /// <param name="value">The property's value.</param>
        /// <return>property that was stored under the specified name before, or
        /// <c>null</c> if there was no such property before.
        /// </return>
        public Object Put(String name, String value)
        {
            Property p = new Property(-1, Variant.VT_LPWSTR, value);
            return Put(new CustomProperty(p, name));
        }

        /// <summary>
        /// Adds a named long property.
        /// </summary>
        /// <param name="name">The property's name.</param>
        /// <param name="value">The property's value.</param>
        /// <return>property that was stored under the specified name before, or
        /// <c>null</c> if there was no such property before.
        /// </return>
        public Object Put(String name, long value)
        {
            Property p = new Property(-1, Variant.VT_I8, value);
            return Put(new CustomProperty(p, name));
        }

        /// <summary>
        /// Adds a named double property.
        /// </summary>
        /// <param name="name">The property's name.</param>
        /// <param name="value">The property's value.</param>
        /// <return>property that was stored under the specified name before, or
        /// <c>null</c> if there was no such property before.
        /// </return>
        public Object Put(String name, Double value)
        {
            Property p = new Property(-1, Variant.VT_R8, value);
            return Put(new CustomProperty(p, name));
        }

        /// <summary>
        /// Adds a named integer property.
        /// </summary>
        /// <param name="name">The property's name.</param>
        /// <param name="value">The property's value.</param>
        /// <return>property that was stored under the specified name before, or
        /// <c>null</c> if there was no such property before.
        /// </return>
        public Object Put(String name, int value)
        {
            Property p = new Property(-1, Variant.VT_I4, value);
            return Put(new CustomProperty(p, name));
        }

        /// <summary>
        /// Adds a named bool property.
        /// </summary>
        /// <param name="name">The property's name.</param>
        /// <param name="value">The property's value.</param>
        /// <return>property that was stored under the specified name before, or
        /// <c>null</c> if there was no such property before.
        /// </return>
        public Object Put(String name, Boolean value)
        {
            Property p = new Property(-1, Variant.VT_BOOL, value);
            return Put(new CustomProperty(p, name));
        }


        /// <summary>
        /// Gets a named value from the custom properties.
        /// </summary>
        /// <param name="name">the name of the value to Get</param>
        /// <return>value or <c>null</c> if a value with the specified
        /// name is not found in the custom properties.
        /// </return>
        public Object Get(String name)
        {
            long? id = dictionary.ContainsValue(name) ? dictionary.GetKey(name) : null;
            CustomProperty cp = id.HasValue ? (base.ContainsKey(id.Value) ? base[id.Value]: null) : null;
            return cp?.Value;
        }

        public object this[string name] => Get(name);

        /// <summary>
        /// Adds a named date property.
        /// </summary>
        /// <param name="name">The property's name.</param>
        /// <param name="value">The property's value.</param>
        /// <return>property that was stored under the specified name before, or
        /// <c>null</c> if there was no such property before.
        /// </return>
        public Object Put(String name, DateTime value)
        {
            Property p = new Property(-1, Variant.VT_FILETIME, value);
            return Put(new CustomProperty(p, name));
        }

        /// <summary>
        /// Returns a Set of all the names of our custom properties.
        /// Equivalent to {@link #nameSet()}
        /// </summary>
        /// <return>set of all the names of our custom properties</return>
        public HashSet<string> KeySet()
        {
            return new HashSet<string>(dictionary.Values);
        }

        /// <summary>
        /// Returns a Set of all the names of our custom properties
        /// </summary>
        /// <return>set of all the names of our custom properties</return>
        public HashSet<string> NameSet()
        {
            return new HashSet<string>(dictionary.Values);
        }

        /// <summary>
        /// Returns a Set of all the IDs of our custom properties
        /// </summary>
        /// <return>set of all the IDs of our custom properties</return>
        public HashSet<string> IdSet()
        {
            return new HashSet<string>(dictionary.Values);
        }


        /// <summary>
        /// Sets the codepage.
        /// </summary>
        /// <param name="codepage">the codepage</param>
        public void SetCodepage(int codepage)
        {
            Property p = new Property(PropertyIDMap.PID_CODEPAGE, Variant.VT_I2, codepage);
            Put(new CustomProperty(p));
        }



        /// <summary>
        /// <para>
        /// </para>
        /// <para>
        /// Gets the dictionary which contains IDs and names of the named custom
        /// properties.
        /// </para>
        /// </summary>
        /// <return>dictionary.</return>
        public Dictionary<long, String> GetDictionary()
        {
            return dictionary._baseDictionary;
        }


        /// <summary>
        /// Checks against both String Name and Long ID
        /// </summary>
        public bool ContainsKey(Object key)
        {
            //key is long

            return (key is long l && dictionary.ContainsKey(l))
                   || dictionary.ContainsValue(key.ToString());
        }

        /// <summary>
        /// Checks against both the property, and its values.
        /// </summary>
        public bool ContainsValue(Object value)
        {
            if(value is CustomProperty cp1)
            {
                return base.ContainsValue(cp1);
            }

            foreach(CustomProperty cp in base.Values)
            {
                if(cp.Value == value)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Gets the codepage.
        /// </summary>
        /// <return>codepage or -1 if the codepage is undefined.</return>
        public int GetCodepage()
        {
            CustomProperty cp = ContainsKey(PropertyIDMap.PID_CODEPAGE) ? this[PropertyIDMap.PID_CODEPAGE] as CustomProperty
                : null;
            return (cp == null) ? -1 : (int) cp.Value;
        }

        /// <summary>
        /// Tells whether this <see cref="CustomProperties"/> instance is pure or one or
        /// more properties of the underlying low-level property Set has been
        /// dropped.
        /// </summary>
        /// <return>true} if the <see cref="CustomProperties"/> is pure, else
        /// <c>false</c>.
        /// </return>
        public bool IsPure
        {
            get => isPure;
            set => isPure = value;
        }

    }
}

