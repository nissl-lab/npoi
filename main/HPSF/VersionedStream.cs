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
namespace NPOI.HPSF
{
    public class VersionedStream
    {
        private GUID _versionGuid;
        private IndirectPropertyName _streamName;

        public VersionedStream(byte[] data, int offset)
        {
            _versionGuid = new GUID(data, offset);
            _streamName = new IndirectPropertyName(data, offset + GUID.SIZE);
        }

        public int Size
        {
            get { return GUID.SIZE + _streamName.Size; }
        }
    }
}