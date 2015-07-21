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
==================================================================== */using System;
namespace NPOI.Util
{

    /**
     * @author Yegor Kozlov
     */
    public class Units
    {
        public static int EMU_PER_PIXEL = 9525;
        public static int EMU_PER_POINT = 12700;

        public static int ToEMU(double value)
        {
            return (int)Math.Round(EMU_PER_POINT * value);
        }

        public static double ToPoints(long emu)
        {
            return (double)emu / EMU_PER_POINT;
        }

        /**
         * Converts a value of type FixedPoint to a decimal number
         *
         * @param fixedPoint
         * @return decimal number
         * 
         * @see <a href="http://msdn.microsoft.com/en-us/library/dd910765(v=office.12).aspx">[MS-OSHARED] - 2.2.1.6 FixedPoint</a>
         */
        public static double FixedPointToDecimal(int fixedPoint) {
        int i = (fixedPoint >> 16);
        int f = (fixedPoint >> 0) & 0xFFFF;
        double decimal1 = (i + f/65536.0);
        return decimal1;
    }
    }
}