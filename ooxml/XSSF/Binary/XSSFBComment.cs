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
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.Util;
    using NPOI.XSSF.UserModel;

    /// <summary>
    /// </summary>
    /// @since 3.16-beta3
    public class XSSFBComment : XSSFComment
    {

        private  CellAddress cellAddress;
        private  String author;
        private  XSSFBRichTextString comment;
        private bool visible = true;

        public XSSFBComment(CellAddress cellAddress, String author, String comment)
        : base(null, null, null)
        {
            ;
            this.cellAddress = cellAddress;
            this.author = author;
            this.comment = new XSSFBRichTextString(comment);
        }

        public override bool Visible
        {
            get => visible;
            set => throw new ArgumentException("XSSFBComment is read only.");
        }
        public override CellAddress Address
        {
            get => cellAddress;
            set => throw new ArgumentException("XSSFBComment is read only");
        }

        public override void SetAddress(int row, int col)
        {
            throw new ArgumentException("XSSFBComment is read only");

        }
        public override int Row
        {
            get => cellAddress.Row;
            set => throw new ArgumentException("XSSFBComment is read only");
        }

        public override int Column
        {
            get => cellAddress.Column;
            set => throw new ArgumentException("XSSFBComment is read only");
        }

        public override String Author
        {
            get => author;
            set => throw new ArgumentException("XSSFBComment is read only");
        }

        public override IRichTextString String
        {
            get => comment;
            set => throw new ArgumentException("XSSFBComment is read only");
        }

        public override IClientAnchor ClientAnchor => null;
    }
}
