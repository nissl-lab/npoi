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

using NPOI.POIFS.Storage;
using NPOI.POIFS.Common;

namespace NPOI.POIFS.Properties
{
    public class PropertyTable : PropertyTableBase, BlockWritable
    {

        private POIFSBigBlockSize _bigBigBlockSize;
        private BlockWritable[] _blocks;
        /**
         * Default constructor
         */
        public PropertyTable(HeaderBlock headerBlock) : base(headerBlock)
        {
            _bigBigBlockSize = headerBlock.BigBlockSize;
            _blocks = null;
        }

        /**
         * reading constructor (used when we've read in a file and we want
         * to extract the property table from it). Populates the
         * properties thoroughly
         *
         * @param startBlock the first block of the property table
         * @param blockList the list of blocks
         *
         * @exception IOException if anything goes wrong (which should be
         *            a result of the input being NFG)
         */
        public PropertyTable(HeaderBlock headerBlock, 
                             RawDataBlockList blockList)
            : base(headerBlock, 
                    PropertyFactory.ConvertToProperties( blockList.FetchBlocks(headerBlock.PropertyStart, -1) ) )
        {
            _bigBigBlockSize = headerBlock.BigBlockSize;
            _blocks      = null;

        }

        /**
         * Prepare to be written Leon
         */

        public void PreWrite()
        {

            List<Property> properties = new List<Property>(_properties.Count);

            for (int i = 0; i < _properties.Count; i++)
                properties.Add(_properties[i]);


            // give each property its index
            for (int k = 0; k < properties.Count; k++)
            {
                properties[ k ].Index = k;
            }

            // allocate the blocks for the property table
            _blocks = PropertyBlock.CreatePropertyBlockArray(_bigBigBlockSize, properties);

            // prepare each property for writing
            for (int k = 0; k < properties.Count; k++)
            {
                properties[ k ].PreWrite();
            }
        }



        /* ********** START implementation of BATManaged ********** */

        /**
         * Return the number of BigBlock's this instance uses
         *
         * @return count of BigBlock instances
         */

        public override int CountBlocks
        {
            get { return (_blocks == null) ? 0 : _blocks.Length; }
        }

        /**
         * Write the storage to an Stream
         *
         * @param stream the Stream to which the stored data should
         *               be written
         *
         * @exception IOException on problems writing to the specified
         *            stream
         */

        public void WriteBlocks(Stream stream)
        {
            if (_blocks != null)
            {
                for (int j = 0; j < _blocks.Length; j++)
                {
                    _blocks[ j ].WriteBlocks(stream);
                }
            }
        }
    }
}
