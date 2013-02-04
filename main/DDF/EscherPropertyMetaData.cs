
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

    /// <summary>
    /// This class stores the type and description of an escher property.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    public class EscherPropertyMetaData
    {
        // Escher property types.
        public const byte TYPE_UNKNOWN = (byte)0;
        public const byte TYPE_bool = (byte)1;
        public const byte TYPE_RGB = (byte)2;
        public const byte TYPE_SHAPEPATH = (byte)3;
        public const byte TYPE_SIMPLE = (byte)4;
        public const byte TYPE_ARRAY = (byte)5;

        private String description;
        private byte type;


        /// <summary>
        /// Initializes a new instance of the <see cref="EscherPropertyMetaData"/> class.
        /// </summary>
        /// <param name="description">The description of the escher property.</param>
        public EscherPropertyMetaData(String description)
        {
            this.description = description;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EscherPropertyMetaData"/> class.
        /// </summary>
        /// <param name="description">The description of the escher property.</param>
        /// <param name="type">The type of the property.</param> 
        public EscherPropertyMetaData(String description, byte type)
        {
            this.description = description;
            this.type = type;
        }

        /// <summary>
        /// Gets the description.
        /// </summary>
        /// <value>The description.</value>
        public String Description
        {
            get { return description; }
        }

        /// <summary>
        /// Gets the type.
        /// </summary>
        /// <value>The type.</value>
        public byte Type
        {
            get { return type; }
        }

    }
}