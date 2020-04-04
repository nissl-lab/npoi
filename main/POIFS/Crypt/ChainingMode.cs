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

namespace NPOI.POIFS.Crypt
{
    using System;

    public class ChainingMode
    {
        // ecb - only for standard encryption
        public static readonly ChainingMode ecb = new ChainingMode("ECB", 1);
        public static readonly ChainingMode cbc = new ChainingMode("CBC", 2);
        /* Cipher feedback chaining (CFB), with an 8-bit window */
        public static readonly ChainingMode cfb = new ChainingMode("CFB8", 3);

        public String jceId { get; set; }
        public int ecmaId { get; set; }
        public ChainingMode(String jceId, int ecmaId)
        {
            this.jceId = jceId;
            this.ecmaId = ecmaId;
        }
    }
}