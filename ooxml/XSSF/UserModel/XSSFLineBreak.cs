/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for Additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

using NPOI.OpenXmlFormats.Dml;
using System;
namespace NPOI.XSSF.UserModel
{

    public class XSSFLineBreak : XSSFTextRun
    {
        private CT_TextCharacterProperties _brProps;

        public XSSFLineBreak(CT_RegularTextRun r, XSSFTextParagraph p, CT_TextCharacterProperties brProps)
            : base(r, p)
        {
            _brProps = brProps;
        }


        protected CT_TextCharacterProperties GetRPr()
        {
            return _brProps;
        }

        /**
         * Always . You cannot change text of a line break.
         */
        public void SetText(String text)
        {
            throw new InvalidOperationException("You cannot change text of a line break, it is always '\\n'");
        }

    }

}