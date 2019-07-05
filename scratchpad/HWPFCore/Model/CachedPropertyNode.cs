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

    using NPOI.HWPF.SPRM;
    using System;


    public class CachedPropertyNode
      : PropertyNode
    {
        protected object _propCache;

        public CachedPropertyNode(int start, int end, SprmBuffer buf)
            : base(start, end, buf)
        {

        }

        protected void FillCache(Object ref1)
        {
            _propCache = ref1;
        }

        protected Object GetCacheContents()
        {
            return _propCache == null ? null : _propCache;
        }

        /**
         * @return This property's property in compressed form.
         */
        public SprmBuffer GetSprmBuf()
        {
            return (SprmBuffer)_buf;
        }


    }
}


