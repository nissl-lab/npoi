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

namespace NPOI.XDDF.UserModel
{
    using NPOI.OpenXmlFormats.Dml;
    public class XDDFPositiveSize2D
    {
        private CT_PositiveSize2D size;
        private long x;
        private long y;

        public XDDFPositiveSize2D(CT_PositiveSize2D size)
        {
            this.size = size;
        }

        public XDDFPositiveSize2D(long x, long y)
        {
            if(x <0 || y < 0)
            {
                throw new ArgumentException("x and y must be positive");
            }
            this.x = x;
            this.y = y;
        }

        public long X
        {
            get
            {
                if(size == null)
                {
                    return x;
                }
                else
                {
                    return size.cx;
                }
            }
        }

        public long Y
        {
            get
            {
                if(size == null)
                {
                    return y;
                }
                else
                {
                    return size.cy;
                }
            }
        }
    }
}
