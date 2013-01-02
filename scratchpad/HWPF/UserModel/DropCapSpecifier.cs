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

namespace NPOI.HWPF.UserModel
{
    using System;
    using NPOI.Util;
    using NPOI.HWPF.Model;

    /**
     * This data structure Is used by a paragraph to determine how it should drop
     * its first letter. I think its the visual effect that will show a giant first
     * letter to a paragraph. I've seen this used in the first paragraph of a
     * book
     *
     * @author Ryan Ackley
     */
    public class DropCapSpecifier:BaseObject
    {
        private short _info;
        private static BitField _type = BitFieldFactory.GetInstance(0x07);
        private static BitField _lines = BitFieldFactory.GetInstance(0xf8);

        public DropCapSpecifier()
        {
            this._info = 0;
        }

        public DropCapSpecifier(byte[] buf, int offset)
            : this(LittleEndian.GetShort(buf, offset))
        {

        }

        public DropCapSpecifier(short info)
        {
            _info = info;
        }

        public short ToShort()
        {
            return _info;
        }
    }
}