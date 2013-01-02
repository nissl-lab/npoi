
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

    public class ShadingDescriptor : BaseObject
    {
        public static int SIZE = 2;

        private short _info;
        private static BitField _icoFore = BitFieldFactory.GetInstance(0x1f);
        private static BitField _icoBack = BitFieldFactory.GetInstance(0x3e0);
        private static BitField _ipat = BitFieldFactory.GetInstance(0xfc00);

        public ShadingDescriptor()
        {
        }

        public ShadingDescriptor(byte[] buf, int offset)
            : this(LittleEndian.GetShort(buf, offset))
        {

        }

        public ShadingDescriptor(short info)
        {
            _info = info;
        }

        public short ToShort()
        {
            return _info;
        }

        public void Serialize(byte[] buf, int offset)
        {
            LittleEndian.PutShort(buf, offset, _info);
        }
    }
}
