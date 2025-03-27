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


namespace NPOI.XDDF.UserModel
{
    using NPOI.OpenXmlFormats.Dml;
    public class XDDFPoint2D
    {
        private CT_Point2D point;
        private long x;
        private long y;

        public XDDFPoint2D(CT_Point2D point)
        {
            this.point = point;
        }

        public XDDFPoint2D(long x, long y)
        {
            this.x = x;
            this.y = y;
        }

        public long X
        {
            get
            {
                if(point == null)
                {
                    return x;
                }
                else
                {
                    return point.x;
                }
            }
        }

        public long Y
        {
            get
            {
                if(point == null)
                {
                    return y;
                }
                else
                {
                    return point.y;
                }
            }
            
        }
    }
}


