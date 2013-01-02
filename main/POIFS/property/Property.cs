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
using System.Text;
using System.Collections;
using System.IO;

using NPOI.POIFS.Dev;
using NPOI.POIFS.Common;
using NPOI.Util;

namespace NPOI.POIFS.Properties
{
    /// <summary>
    /// This abstract base class is the ancestor of all classes
    /// implementing POIFS Property behavior.
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public abstract class Property:Child, POIFSViewable
    {
        private const byte   _default_fill             = ( byte ) 0x00;
        private const int    _name_size_offset         = 0x40;
        private const int _max_name_length = (_name_size_offset / LittleEndianConsts.SHORT_SIZE) - 1;

        protected const int  _NO_INDEX                 = -1;

        // useful offsets
        private const int    _node_color_offset        = 0x43;
        private const int    _previous_property_offset = 0x44;
        private const int    _next_property_offset     = 0x48;
        private const int    _child_property_offset    = 0x4C;
        private const int    _storage_clsid_offset     = 0x50;
        private const int    _user_flags_offset        = 0x60;
        private const int    _seconds_1_offset         = 0x64;
        private const int    _days_1_offset            = 0x68;
        private const int    _seconds_2_offset         = 0x6C;
        private const int    _days_2_offset            = 0x70;
        private const int    _start_block_offset       = 0x74;
        private const int    _size_offset              = 0x78;

        // node colors
        protected const byte _NODE_BLACK               = 1;
        protected const byte _NODE_RED                 = 0;

        // documents must be at least this size to be stored in big blocks
        private const int _big_block_minimum_bytes = POIFSConstants.BIG_BLOCK_MINIMUM_DOCUMENT_SIZE;   //4096;
        private String              _name;
        private ShortField          _name_size;
        private ByteField           _property_type;
        private ByteField           _node_color;
        private IntegerField        _previous_property;
        private IntegerField        _next_property;
        private IntegerField        _child_property;
        private ClassID             _storage_clsid;
        private IntegerField        _user_flags;
        private IntegerField        _seconds_1;
        private IntegerField        _days_1;
        private IntegerField        _seconds_2;
        private IntegerField        _days_2;
        private IntegerField        _start_block;
        private IntegerField        _size;
        private byte[]              _raw_data;
        private int                 _index;
        private Child               _next_child;
        private Child               _previous_child;

        /// <summary>
        /// Initializes a new instance of the <see cref="Property"/> class.
        /// </summary>
        protected Property()
        {
            _raw_data = new byte[POIFSConstants.PROPERTY_SIZE];
            for (int i = 0; i < this._raw_data.Length; i++)
            {
                this._raw_data[i] = _default_fill;
            }
            _name_size         = new ShortField(_name_size_offset);
            _property_type     =
                new ByteField(PropertyConstants.PROPERTY_TYPE_OFFSET);
            _node_color        = new ByteField(_node_color_offset);
            _previous_property = new IntegerField(_previous_property_offset,
                                                  _NO_INDEX, _raw_data);
            _next_property     = new IntegerField(_next_property_offset,
                                                  _NO_INDEX, _raw_data);
            _child_property    = new IntegerField(_child_property_offset,
                                                  _NO_INDEX, _raw_data);
            _storage_clsid     = new ClassID(_raw_data,_storage_clsid_offset);
            _user_flags        = new IntegerField(_user_flags_offset, 0, _raw_data);
            _seconds_1         = new IntegerField(_seconds_1_offset, 0,
                                                  _raw_data);
            _days_1            = new IntegerField(_days_1_offset, 0, _raw_data);
            _seconds_2         = new IntegerField(_seconds_2_offset, 0,
                                                  _raw_data);
            _days_2            = new IntegerField(_days_2_offset, 0, _raw_data);
            _start_block       = new IntegerField(_start_block_offset);
            _size              = new IntegerField(_size_offset, 0, _raw_data);
            _index             = _NO_INDEX;

            this.Name="";
            this.NextChild=null;
            this.PreviousChild=null;
        }

        /// <summary>
        /// Constructor from byte data
        /// </summary>
        /// <param name="index">index number</param>
        /// <param name="array">byte data</param>
        /// <param name="offset">offset into byte data</param>
        protected Property(int index, byte [] array, int offset)
        {
            _raw_data = new byte[ POIFSConstants.PROPERTY_SIZE ];
            System.Array.Copy(array, offset, _raw_data, 0, POIFSConstants.PROPERTY_SIZE);
            _name_size         = new ShortField(_name_size_offset, _raw_data);
            _property_type     =
                new ByteField(PropertyConstants.PROPERTY_TYPE_OFFSET, _raw_data);
            _node_color        = new ByteField(_node_color_offset, _raw_data);
            _previous_property = new IntegerField(_previous_property_offset,
                                                  _raw_data);
            _next_property     = new IntegerField(_next_property_offset,
                                                  _raw_data);
            _child_property    = new IntegerField(_child_property_offset,
                                                  _raw_data);
            _storage_clsid     = new ClassID(_raw_data,_storage_clsid_offset);
            _user_flags        = new IntegerField(_user_flags_offset, 0, _raw_data);
            _seconds_1         = new IntegerField(_seconds_1_offset, _raw_data);
            _days_1            = new IntegerField(_days_1_offset, _raw_data);
            _seconds_2         = new IntegerField(_seconds_2_offset, _raw_data);
            _days_2            = new IntegerField(_days_2_offset, _raw_data);
            _start_block       = new IntegerField(_start_block_offset, _raw_data);
            _size              = new IntegerField(_size_offset, _raw_data);
            _index             = index;
            int name_length = (_name_size.Value / LittleEndianConsts.SHORT_SIZE)
                              - 1;

            if (name_length < 1)
            {
                _name = "";
            }
            else
            {
                char[] char_array  = new char[ name_length ];
                int    name_offset = 0;

                for (int j = 0; j < name_length; j++)
                {
                    char_array[ j ] = ( char ) new ShortField(name_offset,
                                                              _raw_data).Value;
                    name_offset     += LittleEndianConsts.SHORT_SIZE;
                }
                _name = new String(char_array, 0, name_length);
            }
            _next_child     = null;
            _previous_child = null;
        }

        /// <summary>
        /// Write the raw data to an OutputStream.
        /// </summary>
        /// <param name="stream">the OutputStream to which the data Should be
        /// written.</param>
        public void WriteData(Stream stream)
        {
            stream.Write(_raw_data,0,this._raw_data.Length);
        }

        /// <summary>
        /// Gets or sets the start block for the document referred to by this
        /// Property.
        /// </summary>
        /// <value>the start block index</value>
        public int StartBlock
        {
            set
            {
                _start_block.Set(value, _raw_data);
            }
            get 
            {
                return _start_block.Value; 
            }
        }

        /// <summary>
        /// Based on the currently defined size, Should this property use
        /// small blocks?
        /// </summary>
        /// <returns>true if the size Is less than _big_block_minimum_bytes</returns>
        public bool ShouldUseSmallBlocks
        {
            get { return Property.IsSmall(_size.Value); }
        }

        /// <summary>
        /// does the length indicate a small document?
        /// </summary>
        /// <param name="length">length in bytes</param>
        /// <returns>
        /// 	<c>true</c> if the length Is less than
        /// _big_block_minimum_bytes; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsSmall(int length)
        {
            return length < _big_block_minimum_bytes;
        }

        /// <summary>
        /// Gets or sets the name of this property
        /// </summary>
        /// <value>property name</value>
        public String Name
        {
            get { return _name; }
            set
            {
                char[] char_array = value.ToCharArray();
                int limit = Math.Min(char_array.Length, _max_name_length);

                _name = new String(char_array, 0, limit);
                short offset = 0;
                int j = 0;

                for (; j < limit; j++)
                {
                    ShortField.Write(offset, (short)char_array[j], ref _raw_data);
                    offset += LittleEndianConsts.SHORT_SIZE;
                }
                for (; j < _max_name_length + 1; j++)
                {
                    ShortField.Write(offset, (short)0, ref _raw_data);
                    offset += LittleEndianConsts.SHORT_SIZE;
                }

                // double the count, and include the null at the end
                _name_size
                    .Set((short)((limit + 1)
                                    * LittleEndianConsts.SHORT_SIZE), ref _raw_data);
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is directory.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if a directory type Property; otherwise, <c>false</c>.
        /// </value>
        public virtual bool IsDirectory
        {
            get { return false; }
        }

        /// <summary>
        /// Gets or sets the storage class ID for this property stream. ThIs Is the Class ID
        /// of the COM object which can read and write this property stream </summary>
        /// <value>Storage Class ID</value>
        public ClassID StorageClsid 
        {
            get
            {
                return _storage_clsid;
            }
            set
            {
                _storage_clsid = value;
                if (value == null)
                {
                    for (int i = _storage_clsid_offset; i < _storage_clsid_offset + ClassID.LENGTH; i++)
                        _raw_data[i] = (byte)0;
                }
                else
                {
                    value.Write(_raw_data, _storage_clsid_offset);
                }
               
            }
        }
        /// <summary>
        /// Set the property type. Makes no attempt to validate the value.
        /// </summary>
        /// <value>the property type (root, file, directory)</value>
        public byte PropertyType
        {
            set
            {
                _property_type.Set(value, _raw_data);
            }
        }

        /// <summary>
        /// Sets the color of the node.
        /// </summary>
        /// <value>the node color (red or black)</value>
        public byte NodeColor
        {
            set
            {
                _node_color.Set(value, _raw_data);
            }
        }

        /// <summary>
        /// Sets the child property.
        /// </summary>
        /// <value>the child property's index in the Property Table</value>
        public int ChildProperty
        {
            set
            {
                _child_property.Set(value, _raw_data);
            }
        }

        /// <summary>
        /// Get the child property (its index in the Property Table)
        /// </summary>
        /// <value>The index of the child.</value>
        public int ChildIndex
        {
            get
            {
                return _child_property.Value;
            }
        }

        /// <summary>
        /// Gets or sets the size of the document associated with this Property
        /// </summary>
        /// <value>the size of the document, in bytes</value>
        public virtual int Size
        {
            set{
                _size.Set(value, _raw_data);
            }
            get 
            {
                return _size.Value;
            }
        }

        /// <summary>
        /// Gets or sets the index.
        /// </summary>
        /// <value>The index.</value>
        /// Get the index for this Property
        /// @return the index of this Property within its Property Table
        public int Index
        {
            get { return _index; }
            set { _index = value; }
        }

        /// <summary>
        /// Perform whatever activities need to be performed prior to
        /// writing
        /// </summary>
        public abstract void PreWrite();

        /// <summary>
        /// Gets the index of the next child.
        /// </summary>
        /// <value>The index of the next child.</value>
        public int NextChildIndex
        {
            get
            {
                return _next_property.Value;
            }
        }

        /// <summary>
        /// Gets the index of the previous child.
        /// </summary>
        /// <value>The index of the previous child.</value>
        public int PreviousChildIndex
        {
            get{
                return _previous_property.Value;
            }
        }

        /// <summary>
        /// Determines whether the specified index Is valid
        /// </summary>
        /// <param name="index">value to be checked</param>
        /// <returns>
        /// 	<c>true</c> if the index Is valid; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsValidIndex(int index)
        {
            return index != _NO_INDEX;
        }

        #region Child Members

        /// <summary>
        /// Gets or sets the previous child.
        /// </summary>
        /// <value>the new 'previous' child; may be null, which has
        /// the effect of saying there Is no 'previous' child</value>
        public Child PreviousChild
        {
            set
            {
                _previous_child = value;
                _previous_property.Set((value == null) ? _NO_INDEX
                                                       : ((Property)value)
                                                           .Index, _raw_data);
            }
            get
            {
                return _previous_child;
            }
        }
        /// <summary>
        /// Gets or sets the next Child
        /// </summary>
        /// <value> the new 'next' child; may be null, which has the
        /// effect of saying there Is no 'next' child</value>
        public Child NextChild
        {
            set
            {
                _next_child = value;
                _next_property.Set((value == null) ? _NO_INDEX
                                                   : ((Property)value)
                                                       .Index, _raw_data);
            }
            get
            {
                return _next_child;
            }
        }

        #endregion

        #region POIFSViewable Members

        /// <summary>
        /// Get an array of objects, some of which may implement
        /// POIFSViewable
        /// </summary>
        /// <value>an array of Object; may not be null, but may be empty</value>
        public Array ViewableArray
        {
            get
            {
                Array results = new string[5];

                results.SetValue("Name          = \"" + Name + "\"", 0);
                results.SetValue("Property Type = " + _property_type.Value, 1);
                results.SetValue("Node Color    = " + _node_color.Value, 2);
                long time = _days_1.Value;

                time <<= 32;
                time += ((long)_seconds_1.Value) & 0x0000FFFFL;
                results.SetValue("Time 1        = " + time, 3);
                time = _days_2.Value;
                time <<= 32;
                time += ((long)_seconds_2.Value) & 0x0000FFFFL;
                results.SetValue("Time 2        = " + time, 4);
                return results;
            }
        }

        /// <summary>
        /// Get an Iterator of objects, some of which may implement POIFSViewable
        /// </summary>
        /// <value> may not be null, but may have an empty
        /// back end store</value>
        public IEnumerator ViewableIterator
        {
            get
            {
                return ArrayList.ReadOnly(new ArrayList()).GetEnumerator();
            }
        }

        /// <summary>
        /// Give viewers a hint as to whether to call GetViewableArray or
        /// GetViewableIterator
        /// </summary>
        /// <value><c>true</c> if a viewer Should call GetViewableArray; otherwise, <c>false</c>
        /// if a viewer Should call GetViewableIterator
        /// </value>
        public bool PreferArray
        {
            get { return true; }
        }

        /// <summary>
        /// Provides a short description of the object, to be used when a
        /// POIFSViewable object has not provided its contents.
        /// </summary>
        /// <value>The short description.</value>
        public String ShortDescription
        {
            get
            {
                StringBuilder buffer = new StringBuilder();

                buffer.Append("Property: \"").Append(Name).Append("\"");
                return buffer.ToString();
            }
        }

        #endregion
    }
}
