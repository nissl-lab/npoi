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

namespace NPOI.HWPF.Model
{
    using System;
    using NPOI.Util;

    /**
     * Handles the fibRgW / FibRgW97 part of
     *  the FIB (File Information Block)
     */
    public class FIBshortHandler
    {
        public const int MAGICCREATED = 0;
        public const int MAGICREVISED = 1;
        public const int MAGICCREATEDPRIVATE = 2;
        public const int MAGICREVISEDPRIVATE = 3;
        public const int LIDFE = 13;

        internal const int START = 0x20;

        short[] _shorts;

        public FIBshortHandler(byte[] mainStream)
        {
            int offset = START;
            int shortCount = LittleEndian.GetShort(mainStream, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            _shorts = new short[shortCount];

            for (int x = 0; x < shortCount; x++)
            {
                _shorts[x] = LittleEndian.GetShort(mainStream, offset);
                offset += LittleEndianConsts.SHORT_SIZE;
            }
        }

        public short Getshort(int shortCode)
        {
            return _shorts[shortCode];
        }

        internal int SizeInBytes()
        {
            return (_shorts.Length * LittleEndianConsts.SHORT_SIZE) + LittleEndianConsts.SHORT_SIZE;
        }

        internal void Serialize(byte[] mainStream)
        {
            int offset = START;
            LittleEndian.PutShort(mainStream, offset, (short)_shorts.Length);
            offset += LittleEndianConsts.SHORT_SIZE;
            //mainStream.Write(holder);

            for (int x = 0; x < _shorts.Length; x++)
            {
                LittleEndian.PutShort(mainStream, offset, _shorts[x]);
                offset += LittleEndianConsts.SHORT_SIZE;
            }
        }


    }


}