
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

    public class UPX
    {
        private byte[] _upx;

        public UPX(byte[] upx)
        {
            _upx = upx;
        }

        public byte[] GetUPX()
        {
                return _upx;
        }
        public int Size
        {
            get
            {
                return _upx.Length;
            }
        }

        public override bool Equals(Object o)
        {
            UPX upx = (UPX)o;
            return Arrays.Equals(_upx, upx._upx);
        }
    }
}