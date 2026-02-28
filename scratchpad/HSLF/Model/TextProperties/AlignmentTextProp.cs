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

namespace NPOI.HSLF.Model.TextProperties
{

    /**
     * DefInition for the alignment text property.
     */
    public class AlignmentTextProp : TextProp
    {
        public static int LEFT = 0;
        public static int CENTER = 1;
        public static int RIGHT = 2;
        public static int JUSTIFY = 3;
        public static int THAIDISTRIBUTED = 5;
        public static int JUSTIFYLOW = 6;

        public AlignmentTextProp():base(2, 0x800, "alignment")
        {
            
        }
    }

}