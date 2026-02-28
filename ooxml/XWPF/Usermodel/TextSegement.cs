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
     * saves the begin and end position  of a text in a Paragraph
    */
    public class TextSegment
    {
        private PositionInParagraph beginPos;
        private PositionInParagraph endPos;

        public TextSegment()
        {
            this.beginPos = new PositionInParagraph();
            this.endPos = new PositionInParagraph();
        }

        public TextSegment(int beginRun, int endRun, int beginText, int endText, int beginChar, int endChar)
        {
            PositionInParagraph beginPos = new PositionInParagraph(beginRun, beginText, beginChar);
            PositionInParagraph endPos = new PositionInParagraph(endRun, endText, endChar);
            this.beginPos = beginPos;
            this.endPos = endPos;
        }

        public TextSegment(PositionInParagraph beginPos, PositionInParagraph endPos)
        {
            this.beginPos = beginPos;
            this.endPos = endPos;
        }

        public PositionInParagraph BeginPos
        {
            get
            {
                return beginPos;
            }
            set
            {
                beginPos = value;
            }
        }

        public PositionInParagraph EndPos
        {
            get
            {
                return endPos;
            }
        }
        /// <summary>
        /// The index of the start run
        /// </summary>
        public int BeginRun
        {
            get
            {
                return beginPos.Run;
            }
            set 
            {
                beginPos.Run = value;
            }
        }
        /// <summary>
        /// The index of the start CT_Text
        /// </summary>
        public int BeginText
        {
            get
            {
                return beginPos.Text;
            }
            set 
            {
                beginPos.Text = value;
            }
        }
        /// <summary>
        /// The index of the start text character
        /// </summary>
        public int BeginChar
        {
            get
            {
                return beginPos.Char;
            }
            set 
            {
                beginPos.Char = value;
            }
        }
        public int EndRun
        {
            get
            {
                return endPos.Run;
            }
            set 
            {
                endPos.Run = value;
            }
        }
        /// <summary>
        /// the index of the end CT_Text 
        /// </summary>
        public int EndText
        {
            get
            {
                return endPos.Text;
            }
            set 
            {
                endPos.Text = value;
            }
        }
        /// <summary>
        /// the index of the end text character
        /// </summary>
        public int EndChar
        {
            get
            {
                return endPos.Char;
            }
            set 
            {
                endPos.Char = value;
            }
        }
    }

}