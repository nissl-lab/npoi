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

namespace NPOI.XSSF.Util
{

    public class NumericRanges
    {

        public const int NO_OVERLAPS = -1;
        public const int OVERLAPS_1_MINOR = 0;
        public const int OVERLAPS_2_MINOR = 1;
        public const int OVERLAPS_1_WRAPS = 2;
        public const int OVERLAPS_2_WRAPS = 3;

        public static long[] GetOverlappingRange(long[] range1, long[] range2)
        {
            int overlappingType = GetOverlappingType(range1, range2);
            if (overlappingType == OVERLAPS_1_MINOR)
            {
                return new long[] { range2[0], range1[1] };
            }
            else if (overlappingType == OVERLAPS_2_MINOR)
            {
                return new long[] { range1[0], range2[1] };
            }
            else if (overlappingType == OVERLAPS_2_WRAPS)
            {
                return range1;
            }
            else if (overlappingType == OVERLAPS_1_WRAPS)
            {
                return range2;
            }
            return new long[] { -1, -1 };
        }

        public static int GetOverlappingType(long[] range1, long[] range2)
        {
            long min1 = range1[0];
            long max1 = range1[1];
            long min2 = range2[0];
            long max2 = range2[1];
            if (min1 >= min2 && max1 <= max2)
            {
                return OVERLAPS_2_WRAPS;
            }
            else if (min2 >= min1 && max2 <= max1)
            {
                return OVERLAPS_1_WRAPS;
            }
            else if ((min2 >= min1 && min2 <= max1) && max2 >= max1)
            {
                return OVERLAPS_1_MINOR;
            }
            else if ((min1 >= min2 && min1 <= max2) && max1 >= max2)
            {
                return OVERLAPS_2_MINOR;
            }
            return NO_OVERLAPS;

        }

    }
}