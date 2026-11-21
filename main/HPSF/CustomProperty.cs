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

using System;

namespace NPOI.HPSF

{
    /// <summary>
    /// This class represents custom properties in the document summary
    /// information stream. The difference to normal properties is that custom
    /// properties have an optional name. If the name is not <c>null</c> it
    /// will be maintained in the section's dictionary.
    /// </summary>
    public class CustomProperty : MutableProperty
    {

        private String name;

        /// <summary>
        /// Creates an empty <see cref="CustomProperty"/>. The Set methods must be
        /// called to make it usable.
        /// </summary>
        public CustomProperty()
        {
            this.name = null;
        }

        /// <summary>
        /// Creates a <see cref="CustomProperty"/> without a name by copying the
        /// underlying {@link Property}' attributes.
        /// </summary>
        /// <param name="property">the property to copy</param>
        public CustomProperty(Property property)
            : this(property, null)
        {

        }

        /// <summary>
        /// Creates a <see cref="CustomProperty"/> with a name.
        /// </summary>
        /// <param name="property">This property's attributes are copied to the new custom
        /// property.
        /// </param>
        /// <param name="name">The new custom property's name.</param>
        public CustomProperty(Property property, String name)
            : base(property)
        {
            this.name = name;
        }

        /// <summary>
        /// Get or set the property's name.
        /// </summary>
        /// <return>property's name.</return>
        public String Name
        {
            get { return this.name; }
            set { this.name = value; }
        }


        /// <summary>
        /// Compares two custom properties for equality. The method returns
        /// <c>true</c> if all attributes of the two custom properties are
        /// equal.
        /// </summary>
        /// <param name="o">The custom property to compare with.</param>
        /// <return>true} if both custom properties are equal, else
        /// <c>false</c>.
        /// </return>
        /// 
        /// @see java.Util.AbstractSet#equals(java.lang.Object)
        public bool EqualsContents(Object o)
        {
            CustomProperty c = (CustomProperty)o;
            String name1 = c.Name;
            String name2 = this.Name;
            bool equalNames = true;
            if(name1 == null)
            {
                equalNames = name2 == null;
            }
            else
            {
                equalNames = name1.Equals(name2);
            }
            return equalNames && c.ID == this.ID
                    && c.GetType() == this.GetType()
                    && c.Value.Equals(this.Value);
        }

        /// <summary>
        /// </summary>
        /// @see java.Util.AbstractSet#hashCode()
        public override int GetHashCode()
        {
            return (int) this.ID;
        }
        public override bool Equals(Object o)
        {
            return (o is CustomProperty) ? EqualsContents(o) : false;
        }
    }
}

