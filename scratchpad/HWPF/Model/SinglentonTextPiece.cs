
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
using System.Collections.Generic;
using System.Text;

namespace NPOI.HWPF.Model
{
    public class SinglentonTextPiece : TextPiece
    {

        public SinglentonTextPiece(StringBuilder buffer) :
            base(0, buffer.Length, Encoding.GetEncoding("UTF-16LE").GetBytes(buffer.ToString()),
                   new PieceDescriptor(new byte[8], 0))
        {

        }

        public override int BytesLength
        {
            get
            {
                return GetStringBuilder().Length * 2;
            }
        }

        public override int CharacterLength
        {
            get
            {
                return GetStringBuilder().Length;
            }
        }

        public override int GetCP()
        {
            return 0;
        }

        public int GetEnd()
        {
            return CharacterLength;
        }

        public int GetStart()
        {
            return 0;
        }

        public override String ToString()
        {
            return "SinglentonTextPiece (" + CharacterLength + " chars)";
        }
    }

}