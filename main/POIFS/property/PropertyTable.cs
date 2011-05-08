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
using System;
using System.Collections;
using System.IO;

using NPOI.POIFS.Storage;
using NPOI.POIFS.FileSystem;
using NPOI.POIFS.Common;

namespace NPOI.POIFS.Properties
{
    public class PropertyTable:BATManaged, BlockWritable
    {
        private int             _start_block;
        private IList            _properties;
        private BlockWritable[] _blocks=null;

        /**
         * Default constructor
         */

        public PropertyTable()
        {
            _start_block = POIFSConstants.END_OF_CHAIN;
            _properties  = new ArrayList();
            AddProperty(new RootProperty());
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

        public PropertyTable(int startBlock,
                             RawDataBlockList blockList)
        {
            _start_block = POIFSConstants.END_OF_CHAIN;
            _blocks      = null;
            _properties  =
                PropertyFactory
                    .ConvertToProperties(blockList.FetchBlocks(startBlock,-1));
            PopulatePropertyTree(( DirectoryProperty ) _properties[0]);
        }

        /**
         * Add a property to the list of properties we manage
         *
         * @param property the new Property to manage
         */

        public void AddProperty(Property property)
        {
            _properties.Add(property);
        }

        /**
         * Remove a property from the list of properties we manage
         *
         * @param property the Property to be Removed
         */

        public void RemoveProperty(Property property)
        {
            _properties.Remove(property);
        }

        /**
         * Get the root property
         *
         * @return the root property
         */

        public RootProperty Root
        {

            // it's always the first element in the List
            get{return ( RootProperty ) _properties[0];}
        }

        /**
         * Prepare to be written
         */

        public void PreWrite()
        {
            Property[] array = new Property[this._properties.Count];
            this._properties.CopyTo(array, 0);

            Property[] properties = array;

            // give each property its index
            for (int k = 0; k < properties.Length; k++)
            {
                properties[ k ].Index=k;
            }

            // allocate the blocks for the property table
            _blocks = PropertyBlock.CreatePropertyBlockArray(_properties);

            // prepare each property for writing
            for (int k = 0; k < properties.Length; k++)
            {
                properties[ k ].PreWrite();
            }
        }

        /**
         * Get the start block for the property table
         *
         * @return start block index
         */

        public int StartBlock
        {
            get { return _start_block; }
            set { _start_block = value; }
        }

        private void PopulatePropertyTree(DirectoryProperty root)
        {
            int index = root.ChildIndex;

            if (!Property.IsValidIndex(index))
            {

                // property has no children
                return;
            }
            Stack children = new Stack();

            children.Push(_properties[index]);
            while (!(children.Count==0))
            {
                Property property = ( Property ) children.Pop();
                if (property == null)
                {
                    continue;
                }
                root.AddChild(property);
                
                if (property.IsDirectory)
                {
                    PopulatePropertyTree(( DirectoryProperty ) property);
                }
                index = property.PreviousChildIndex;
                if (Property.IsValidIndex(index))
                {
                    children.Push(_properties[index]);
                }
                index = property.NextChildIndex;
                if (Property.IsValidIndex(index))
                {
                    children.Push(_properties[index]);
                }
            }
        }

        /* ********** START implementation of BATManaged ********** */

        /**
         * Return the number of BigBlock's this instance uses
         *
         * @return count of BigBlock instances
         */

        public int CountBlocks
        {
            get{return (_blocks == null) ? 0
                                     : _blocks.Length;
            }
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
