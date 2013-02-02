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

/* ================================================================
 * About NPOI
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * Author's Blog: tonyqus.wordpress.com.cn (wp.tonyqus.cn)
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/

using System.IO;
using System.Collections.Generic;

using NPOI.POIFS.Properties;
using NPOI.POIFS.Common;

namespace NPOI.POIFS.Storage
{
    /// <summary>
    /// A block of Property instances
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public class PropertyBlock : BigBlock
    {
        private class AnonymousProperty : Property
        {
            public override void PreWrite()
            {
            }

            public override bool IsDirectory
            {
                get
                {
                    return false;
                }
            }
        }



        private Property[]       _properties;

        /// <summary>
        /// Create a single instance initialized with default values
        /// </summary>
        /// <param name="bigBlockSize"></param>
        /// <param name="properties">the properties to be inserted</param>
        /// <param name="offset">the offset into the properties array</param>
        protected PropertyBlock(POIFSBigBlockSize bigBlockSize, Property[] properties, int offset) : base(bigBlockSize)
        {
            _properties = new Property[ bigBlockSize.GetPropertiesPerBlock() ];
            for (int j = 0; j < _properties.Length; j++)
            {
                _properties[ j ] = properties[ j + offset ];
            }
        }

        /// <summary>
        /// Create an array of PropertyBlocks from an array of Property
        /// instances, creating empty Property instances to make up any
        /// shortfall
        /// </summary>
        /// <param name="bigBlockSize"></param>
        /// <param name="properties">the Property instances to be converted into PropertyBlocks, in a java List</param>
        /// <returns>the array of newly created PropertyBlock instances</returns>
        public static BlockWritable [] CreatePropertyBlockArray( POIFSBigBlockSize bigBlockSize,
                                        List<Property> properties)
            {
            int _properties_per_block = bigBlockSize.GetPropertiesPerBlock();

            int blockCount = (properties.Count + _properties_per_block - 1) / _properties_per_block;

            Property[] toBeWritten = new Property[blockCount * _properties_per_block];

            System.Array.Copy(properties.ToArray(), 0, toBeWritten, 0, properties.Count);

            for (int i = properties.Count; i < toBeWritten.Length; i++)
            {
                toBeWritten[i] = new AnonymousProperty();
            }

            BlockWritable[] rvalue = new BlockWritable[blockCount];

            for (int i = 0; i < blockCount; i++)
                rvalue[i] = new PropertyBlock(bigBlockSize, toBeWritten, i * _properties_per_block);

            return rvalue;
        }

        /// <summary>
        /// Write the block's data to an OutputStream
        /// </summary>
        /// <param name="stream">the OutputStream to which the stored data should be written</param>
        public override void WriteData(Stream stream)
        {
            int _properties_per_block = bigBlockSize.GetPropertiesPerBlock();

            for (int i = 0; i < _properties_per_block; i++)
                _properties[i].WriteData(stream);
        }
    }
}
