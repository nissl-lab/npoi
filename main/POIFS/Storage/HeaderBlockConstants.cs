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

/* ================================================================
 * About NPOI
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * Author's Blog: tonyqus.wordpress.com.cn (wp.tonyqus.cn)
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/

using NPOI.POIFS.Common;
using NPOI.Util;

namespace NPOI.POIFS.Storage
{
    /// <summary>
    /// Constants used in reading/writing the Header block
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public class HeaderBlockConstants
    {
        //public const long _signature = unchecked((long)16220472316735377360L);
        public const long _signature = unchecked((long)0xE11AB1A1E011CFD0L);
        public const int  _bat_array_offset        = 0x4c;
        public const int _max_bats_in_header =
            (POIFSConstants.BIG_BLOCK_SIZE - _bat_array_offset)
            / LittleEndianConsts.INT_SIZE;

        // useful offsets
        public const int _signature_offset = 0;
        public const int _bat_count_offset = 0x2C;
        public const int _property_start_offset = 0x30;
        public const int _sbat_start_offset = 0x3C;
        public const int _sbat_block_count_offset = 0x40;
        public const int _xbat_start_offset = 0x44;
        public const int _xbat_count_offset = 0x48;
    } 
}
