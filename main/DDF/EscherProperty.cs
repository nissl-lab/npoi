
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

namespace NPOI.DDF
{
    using System;
    using System.Text;

    /// <summary>
    /// This is the abstract base class for all escher properties.
    /// @see EscherOptRecord
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    abstract public class EscherProperty
    {
        protected short id;

        /// <summary>
        /// Initializes a new instance of the <see cref="EscherProperty"/> class.
        /// </summary>
        /// <param name="id">The id is distinct from the actual property number.  The id includes the property number the blip id
        /// flag and an indicator whether the property is complex or not.</param>
        public EscherProperty(short id)
        {
            this.id = id;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EscherProperty"/> class.The three parameters are combined to form a property
        /// id.
        /// </summary>
        /// <param name="propertyNumber">The property number.</param>
        /// <param name="isComplex">if set to <c>true</c> [is complex].</param>
        /// <param name="isBlipId">if set to <c>true</c> [is blip id].</param> 
        public EscherProperty(short propertyNumber, bool isComplex, bool isBlipId)
        {
            this.id = (short)(propertyNumber +
                    (isComplex ? unchecked((short)0x8000) : (short)0x0) +
                    (isBlipId ? (short)0x4000 : (short)0x0));
        }

        /// <summary>
        /// Gets the id.
        /// </summary>
        /// <value>The id.</value>
        public virtual  short Id
        {
            get { return id; }
        }

        /// <summary>
        /// Gets the property number.
        /// </summary>
        /// <value>The property number.</value>
        public virtual  short PropertyNumber
        {
            get { return (short)(id & (short)0x3FFF); }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is complex.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is complex; otherwise, <c>false</c>.
        /// </value>
        public virtual  bool IsComplex
        {
            get { return (id & unchecked((short)0x8000)) != 0; }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is blip id.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is blip id; otherwise, <c>false</c>.
        /// </value>
        public virtual  bool IsBlipId
        {
            get { return (id & (short)0x4000) != 0; }
        }

        /// <summary>
        /// Gets the name.
        /// </summary>
        /// <value>The name.</value>
        public virtual String Name
        {
            get { return EscherProperties.GetPropertyName(PropertyNumber); }
        }

        /// <summary>
        /// Most properties are just 6 bytes in Length.  Override this if we're
        /// dealing with complex properties.
        /// </summary>
        /// <value>The size of the property.</value>
        public virtual int PropertySize
        {
            get { return 6; }
        }

        public virtual String ToXml(String tab)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(tab)
                   .Append("<")
                   .Append(GetType().Name)
                   .Append(" id=\"")
                   .Append(Id)
                   .Append("\" name=\"")
                   .Append(Name)
                   .Append("\" blipId=\"")
                   .Append(IsBlipId).Append("\"/>\n");
            return builder.ToString();
        }

        /// <summary>
        /// Escher properties consist of a simple fixed Length part and a complex variable Length part.
        /// The fixed Length part is Serialized first.
        /// </summary>
        /// <param name="data">The data.</param>
        /// <param name="pos">The pos.</param>
        /// <returns></returns>
        abstract public int SerializeSimplePart(byte[] data, int pos);
        /// <summary>
        /// Escher properties consist of a simple fixed Length part and a complex variable Length part.
        /// The fixed Length part is Serialized first.
        /// </summary>
        /// <param name="data">The data.</param>
        /// <param name="pos">The pos.</param>
        /// <returns></returns>
        abstract public int SerializeComplexPart(byte[] data, int pos);
    }
}