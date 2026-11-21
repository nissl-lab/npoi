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

using NPOI.HSSF.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.OOXML.XSSF.UserModel
{
    /// <summary>
    /// Uses the legacy colors defined in HSSF for index lookups
    /// </summary>
    public class DefaultIndexedColorMap:IIndexedColorMap
    {
        public byte[] GetRGB(int index)
        {
            return GetDefaultRGB(index);
        }
        /// <summary>
        /// RGB bytes from HSSF default color by index
        /// </summary>
        /// <param name="index"></param>
        /// <returns>RGB bytes from HSSF default color by index</returns>
        public static byte[] GetDefaultRGB(int index)
        {
            var colorDict = HSSFColor.GetIndexHash();
            colorDict.TryGetValue(index, out HSSFColor hssfColor);
            //HSSFColor hssfColor = HSSFColor.GetIndexHash()[index];
            if (hssfColor == null) return null;
            byte[] rgbShort = hssfColor.GetTriplet();
            return rgbShort;
        }
    }
}
