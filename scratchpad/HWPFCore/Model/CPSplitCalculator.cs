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

namespace NPOI.HWPF.Model
{

    using NPOI.HWPF;

    /**
     * Helper class for {@link HWPFDocument}, which figures out
     *  where different kinds of text can be found within the
     *  overall CP splurge.
     */
    public class CPSplitCalculator
    {
        private FileInformationBlock fib;
        public CPSplitCalculator(FileInformationBlock fib)
        {
            this.fib = fib;
        }

        /**
         * Where the main document text starts. Always 0.
         */
        public int GetMainDocumentStart()
        {
            return 0;
        }
        /**
         * Where the main document text ends.
         * Given by FibRgLw97.ccpText
         */
        public int GetMainDocumentEnd()
        {
            return fib.GetCcpText();
        }

        /**
         * Where the Footnotes text starts.
         * Follows straight on from the main text.
         */
        public int GetFootnoteStart()
        {
            return GetMainDocumentEnd();
        }
        /**
         * Where the Footnotes text ends.
         * Length comes from FibRgLw97.ccpFtn
         */
        public int GetFootnoteEnd()
        {
            return GetFootnoteStart() +
                fib.GetCcpFtn();
        }

        /**
         * Where the "Header Story" text starts.
         * Follows straight on from the footnotes.
         */
        public int GetHeaderStoryStart()
        {
            return GetFootnoteEnd();
        }
        /**
         * Where the "Header Story" text ends.
         * Length comes from FibRgLw97.ccpHdd
         */
        public int GetHeaderStoryEnd()
        {
            return GetHeaderStoryStart() +
                fib.GetCcpHdd();
        }

        /**
         * Where the Comment (Atn) text starts.
         * Follows straight on from the header stories.
         */
        public int GetCommentsStart()
        {
            return GetHeaderStoryEnd();
        }
        /**
         * Where the Comment (Atn) text ends.
         * Length comes from FibRgLw97.ccpAtn
         */
        public int GetCommentsEnd()
        {
            return GetCommentsStart() +
                fib.GetCcpCommentAtn();
        }

        /**
         * Where the End Note text starts.
         * Follows straight on from the comments.
         */
        public int GetEndNoteStart()
        {
            return GetCommentsEnd();
        }
        /**
         * Where the End Note text ends.
         * Length comes from FibRgLw97.ccpEdn
         */
        public int GetEndNoteEnd()
        {
            return GetEndNoteStart() +
                fib.GetCcpEdn();
        }

        /**
         * Where the Main Textbox text starts.
         * Follows straight on from the end note.
         */
        public int GetMainTextboxStart()
        {
            return GetEndNoteEnd();
        }
        /**
         * Where the Main textbox text ends.
         * Length comes from FibRgLw97.ccpTxBx
         */
        public int GetMainTextboxEnd()
        {
            return GetMainTextboxStart() +
                fib.GetCcpTxtBx();
        }

        /**
         * Where the Header Textbox text starts.
         * Follows straight on from the main textbox.
         */
        public int GetHeaderTextboxStart()
        {
            return GetMainTextboxEnd();
        }
        /**
         * Where the Header textbox text ends.
         * Length comes from FibRgLw97.ccpHdrTxBx
         */
        public int GetHeaderTextboxEnd()
        {
            return GetHeaderTextboxStart() +
                fib.GetCcpHdrTxtBx();
        }
    }
}

