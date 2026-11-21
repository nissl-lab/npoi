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
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.XSSF.Binary
{
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using NPOI.XSSF.UserModel;

    /// <summary>
    /// Wrapper class around String so that we can use it in Comment.
    /// Nothing has been implemented yet except for <see cref="getString()" />.
    /// </summary>
    /// @since 3.16-beta3
    public class XSSFBRichTextString : XSSFRichTextString
    {

        private  String @string;

        public XSSFBRichTextString(String @string)
        {
            this.@string = @string;
        }

        public new void ApplyFont(int startIndex, int endIndex, short fontIndex)
        {

        }

        public new void ApplyFont(int startIndex, int endIndex, IFont font)
        {

        }

        public new void ApplyFont(IFont font)
        {

        }

        public new void ClearFormatting()
        {

        }
        public override String String => @string;

        public override int Length => @string.Length;


        public override int NumFormattingRuns => 0;

        public int GetIndexOfFormattingRun(int index)
        {
            return 0;
        }

        public new void ApplyFont(short fontIndex)
        {

        }
    }
}

