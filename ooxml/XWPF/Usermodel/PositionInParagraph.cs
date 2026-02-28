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


    /**
     * postion of a character in a paragrapho
    * 1st RunPositon
    * 2nd TextPosition
    * 3rd CharacterPosition 
    * 
    *
    */
    public class PositionInParagraph
    {
        private int posRun = 0, posText = 0, posChar = 0;

        public PositionInParagraph()
        {
        }

        public PositionInParagraph(int posRun, int posText, int posChar)
        {
            this.posRun = posRun;
            this.posChar = posChar;
            this.posText = posText;
        }

        public int Run
        {
            get
            {
                return posRun;
            }
            set
            {
                this.posRun = value;
            }
        }

        public int Text
        {
            get
            {
                return posText;
            }
            set
            {
                this.posText = value;
            }
        }


        public int Char
        {
            get
            {
                return posChar;
            }
            set
            {
                this.posChar = value;
            }
        }
    }

}