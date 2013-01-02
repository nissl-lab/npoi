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
namespace NPOI.XWPF.UserModel
{
    using System;
    using NPOI.OpenXmlFormats.Wordprocessing;

    /**
     * @author Philipp Epp
     *
     */
    public class XWPFNum
    {
        private CT_Num ctNum;
        protected XWPFNumbering numbering;

        public XWPFNum()
        {
            this.ctNum = null;
            this.numbering = null;
        }

        public XWPFNum(CT_Num ctNum)
        {
            this.ctNum = ctNum;
            this.numbering = null;
        }

        public XWPFNum(XWPFNumbering numbering)
        {
            this.ctNum = null;
            this.numbering = numbering;
        }

        public XWPFNum(CT_Num ctNum, XWPFNumbering numbering)
        {
            this.ctNum = ctNum;
            this.numbering = numbering;
        }

        public XWPFNumbering GetNumbering()
        {
            return numbering;
        }

        public CT_Num GetCTNum()
        {
            return ctNum;
        }

        public void SetNumbering(XWPFNumbering numbering)
        {
            this.numbering = numbering;
        }

        public void SetCTNum(CT_Num ctNum)
        {
            this.ctNum = ctNum;
        }
    }
}