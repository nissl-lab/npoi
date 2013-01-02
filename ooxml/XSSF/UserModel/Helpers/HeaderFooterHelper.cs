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

namespace NPOI.XSSF.UserModel.Helpers
{
    using System;
    using System.Text;

    public class HeaderFooterHelper
    {
        // Note - XmlBeans handles entity encoding for us,
        //  so these should be & forms, not the &amp; ones!
        private static String HeaderFooterEntity_L = "&L";
        private static String HeaderFooterEntity_C = "&C";
        private static String HeaderFooterEntity_R = "&R";

        // These are other entities that may be used in the
        //  left, center or right. Not exhaustive
        public static String HeaderFooterEntity_File = "&F";
        public static String HeaderFooterEntity_Date = "&D";
        public static String HeaderFooterEntity_Time = "&T";

        public String GetLeftSection(String str)
        {
            return GetParts(str)[0];
        }
        public String GetCenterSection(String str)
        {
            return GetParts(str)[1];
        }
        public String GetRightSection(String str)
        {
            return GetParts(str)[2];
        }

        public String SetLeftSection(String str, String newLeft)
        {
            String[] parts = GetParts(str);
            parts[0] = newLeft;
            return JoinParts(parts);
        }
        public String SetCenterSection(String str, String newCenter)
        {
            String[] parts = GetParts(str);
            parts[1] = newCenter;
            return JoinParts(parts);
        }
        public String SetRightSection(String str, String newRight)
        {
            String[] parts = GetParts(str);
            parts[2] = newRight;
            return JoinParts(parts);
        }

        /**
         * Split into left, center, right
         */
        private String[] GetParts(String str)
        {
            String[] parts = new String[] { "", "", "" };
            if (str == null)
                return parts;

            // They can come in any order, which is just nasty
            // Work backwards from the end, picking the last
            //  on off each time as we go
            int lAt = 0;
            int cAt = 0;
            int rAt = 0;

            while (
                // Ensure all indicies get updated, then -1 tested
                (lAt = str.IndexOf(HeaderFooterEntity_L)) > -2 &&
                (cAt = str.IndexOf(HeaderFooterEntity_C)) > -2 &&
                (rAt = str.IndexOf(HeaderFooterEntity_R)) > -2 &&
                (lAt > -1 || cAt > -1 || rAt > -1)
            )
            {
                // Pick off the last one
                if (rAt > cAt && rAt > lAt)
                {
                    parts[2] = str.Substring(rAt + HeaderFooterEntity_R.Length);
                    str = str.Substring(0, rAt);
                }
                else if (cAt > rAt && cAt > lAt)
                {
                    parts[1] = str.Substring(cAt + HeaderFooterEntity_C.Length);
                    str = str.Substring(0, cAt);
                }
                else
                {
                    parts[0] = str.Substring(lAt + HeaderFooterEntity_L.Length);
                    str = str.Substring(0, lAt);
                }
            }

            return parts;
        }
        private String JoinParts(String[] parts)
        {
            return JoinParts(parts[0], parts[1], parts[2]);
        }
        private String JoinParts(String l, String c, String r)
        {
            StringBuilder ret = new StringBuilder();

            // Join as c, l, r
            if (c.Length > 0)
            {
                ret.Append(HeaderFooterEntity_C);
                ret.Append(c);
            }
            if (l.Length > 0)
            {
                ret.Append(HeaderFooterEntity_L);
                ret.Append(l);
            }
            if (r.Length > 0)
            {
                ret.Append(HeaderFooterEntity_R);
                ret.Append(r);
            }

            return ret.ToString();
        }
    }


}