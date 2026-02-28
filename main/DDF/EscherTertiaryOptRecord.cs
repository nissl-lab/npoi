/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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
using System;
namespace NPOI.DDF
{

    /**
     * "The OfficeArtTertiaryFOPT record specifies a table of OfficeArtRGFOPTE properties, as defined in section 2.3.1."
     * -- [MS-ODRAW] -- v20110608; Office Drawing Binary File Format
     *
     * @author Sergey Vladimirov (vlsergey {at} gmail {dot} com)
     */
    public class EscherTertiaryOptRecord : AbstractEscherOptRecord
    {
        public const short RECORD_ID = unchecked((short)0xF122);

        public override String RecordName
        {
            get
            {
                return "TertiaryOpt";
            }
        }
    }
}