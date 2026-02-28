/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */
using NPOI.Util;

namespace NPOI.HPSF
{
    internal class GUID
    {
        private int _data1;
        private short _data2;
        private short _data3;
        private long _data4;
        internal GUID() {}
    
        internal void Read( LittleEndianByteArrayInputStream lei ) {
            _data1 = lei.ReadInt();
            _data2 = lei.ReadShort();
            _data3 = lei.ReadShort();
            _data4 = lei.ReadLong();
        }
    }
}