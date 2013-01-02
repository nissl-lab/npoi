
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

namespace NPOI.HWPF.SPRM
{
    using System;
    using System.Collections;



    public class SprmUncompressor
    {
        public SprmUncompressor()
        {
        }

        /**
         * Converts an int into a boolean. If the int Is non-zero, it returns true.
         * Otherwise it returns false.
         *
         * @param x The int to convert.
         *
         * @return A bool whose value depends on x.
         */
        public static bool GetFlag(int x)
        {
            if (x != 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }


    }
}