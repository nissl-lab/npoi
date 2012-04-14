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
    public class XWPFAbstractNum
    {
        private CT_AbstractNum ctAbstractNum;
        protected XWPFNumbering numbering;

        protected XWPFAbstractNum()
        {
            this.ctAbstractNum = null;
            this.numbering = null;

        }
        public XWPFAbstractNum(CT_AbstractNum abstractNum)
        {
            this.ctAbstractNum = abstractNum;
        }

        public XWPFAbstractNum(CT_AbstractNum ctAbstractNum, XWPFNumbering numbering)
        {
            this.ctAbstractNum = ctAbstractNum;
            this.numbering = numbering;
        }
        public CT_AbstractNum GetAbstractNum()
        {
            return ctAbstractNum;
        }

        public XWPFNumbering GetNumbering()
        {
            return numbering;
        }

        public CT_AbstractNum GetCTAbstractNum()
        {
            return ctAbstractNum;
        }

        public void SetNumbering(XWPFNumbering numbering)
        {
            this.numbering = numbering;
        }

    }

}