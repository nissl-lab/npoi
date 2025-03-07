
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License Is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HWPF.Model
{
    using System;
    using System.Collections;
    using NPOI.Util;


    /**
     * common data structure in a Word file. Contains an array of 4 byte ints in
     * the front that relate to an array of abitrary data structures in the back.
     *
     *
     * @author Ryan Ackley
     */
    public class PlexOfCps
    {
        private int _iMac;
        private int _offset;
        private int _sizeOfStruct;
        private ArrayList _props;


        public PlexOfCps(int sizeOfStruct)
        {
            _props = new ArrayList();
            _sizeOfStruct = sizeOfStruct;
        }

        /**
         * Constructor
         *
         * @param size The size in bytes of this PlexOfCps
         * @param sizeOfStruct The size of the data structure type stored in
         *        this PlexOfCps.
         */
        public PlexOfCps(byte[] buf, int start, int size, int sizeOfStruct)
        {
            // Figure out the number we hold
            _iMac = (size - 4) / (4 + sizeOfStruct);

            _sizeOfStruct = sizeOfStruct;
            _props = new ArrayList(_iMac);

            for (int x = 0; x < _iMac; x++)
            {
                _props.Add(GetProperty(x, buf, start));
            }
        }

        public GenericPropertyNode GetProperty(int index)
        {
            return (GenericPropertyNode)_props[index];
        }

        public void AddProperty(GenericPropertyNode node)
        {
            _props.Add(node);
        }

        public byte[] ToByteArray()
        {
            int size = _props.Count;
            int cpBufSize = ((size + 1) * LittleEndianConsts.INT_SIZE);
            int structBufSize = +(_sizeOfStruct * size);
            int bufSize = cpBufSize + structBufSize;

            byte[] buf = new byte[bufSize];

            GenericPropertyNode node = null;
            for (int x = 0; x < size; x++)
            {
                node = (GenericPropertyNode)_props[x];

                // put the starting offset of the property into the plcf.
                LittleEndian.PutInt(buf, (LittleEndianConsts.INT_SIZE * x), node.Start);

                // put the struct into the plcf
                System.Array.Copy(node.Bytes, 0, buf, cpBufSize + (x * _sizeOfStruct),
                                 _sizeOfStruct);
            }
            // put the ending offset of the last property into the plcf.
            LittleEndian.PutInt(buf, LittleEndianConsts.INT_SIZE * size, node.End);

            return buf;

        }

        private GenericPropertyNode GetProperty(int index, byte[] buf, int offset)
        {
            int start = LittleEndian.GetInt(buf, offset + GetIntOffset(index));
            int end = LittleEndian.GetInt(buf, offset + GetIntOffset(index + 1));

            byte[] structure = new byte[_sizeOfStruct];
            System.Array.Copy(buf, offset + GetStructOffset(index), structure, 0, _sizeOfStruct);

            return new GenericPropertyNode(start, end, structure);
        }

        private int GetIntOffset(int index)
        {
            return index * 4;
        }

        /**
         * returns the number of data structures in this PlexOfCps.
         *
         * @return The number of data structures in this PlexOfCps
         */
        public int Length
        {
            get
            {
                return _iMac;
            }
        }


        internal GenericPropertyNode[] ToPropertiesArray()
        {
            if (_props == null || _props.Count==0)
                return Array.Empty<GenericPropertyNode>();

            return (GenericPropertyNode[])_props.ToArray(typeof(GenericPropertyNode));
        }

        /**
         * Returns the offset, in bytes, from the beginning if this PlexOfCps to
         * the data structure at index.
         *
         * @param index The index of the data structure.
         *
         * @return The offset, in bytes, from the beginning if this PlexOfCps to
         *         the data structure at index.
         */
        private int GetStructOffset(int index)
        {
            return (4 * (_iMac + 1)) + (_sizeOfStruct * index);
        }
    }
}