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
using System.Collections.Generic;
using System.IO;
using NPOI.POIFS.FileSystem;
using NPOI.POIFS.Storage;

namespace NPOI.POIFS.Properties
{
    public abstract class PropertyTableBase : BATManaged
    {
        protected HeaderBlock _header_block;
        protected List<Property> _properties;

        public PropertyTableBase(HeaderBlock header_block)
        {
            _header_block = header_block;
            _properties = new List<Property>();
            AddProperty(new RootProperty());
        }

        public PropertyTableBase(HeaderBlock header_block, List<Property> properties)
        {
            _header_block = header_block;
            _properties = properties;
            PopulatePropertyTree((DirectoryProperty)_properties[0]);
        }

        public void AddProperty(Property property)
        {
            _properties.Add(property);
        }

        public void RemoveProperty(Property property)
        {
            _properties.Remove(property);
        }

        public RootProperty Root
        {
            get { return (RootProperty)(_properties[0]); }
        }

        protected void PopulatePropertyTree(DirectoryProperty root)
        {
            try
            {
                int index = root.ChildIndex;

                if (!Property.IsValidIndex(index))
                    return;

                Stack<Property> children = new Stack<Property>();

                children.Push(_properties[index]);

                while (children.Count != 0)
                {
                    Property property = children.Pop();
                    if (property == null)
                    {
                        // unknown / unsupported / corrupted property, skip
                        continue;
                    }
                    root.AddChild(property);

                    if (property.IsDirectory)
                    {
                        PopulatePropertyTree((DirectoryProperty)property);
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
            catch (IOException ex)
            {
                throw ex;
            }
        }


        #region BATManaged Members

        public virtual int StartBlock
        {
            get { return _header_block.PropertyStart; }
            set { _header_block.PropertyStart = value; }
        }


        public virtual int CountBlocks
        {
            get { throw new NotImplementedException(); }
        }

        #endregion
    }
}
