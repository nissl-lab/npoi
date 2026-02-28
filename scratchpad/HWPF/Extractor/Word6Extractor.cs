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

using NPOI.POIFS.FileSystem;
using System.IO;
using System;
using System.Text;
using NPOI.HWPF.UserModel;
namespace NPOI.HWPF.Extractor
{


    /**
     * Class to extract the text from old (Word 6 / Word 95) Word Documents.
     *
     * This should only be used on the older files, for most uses you
     *  should call {@link WordExtractor} which deals properly 
     *  with HWPF.
     *
     * @author Nick Burch
     */
    public class Word6Extractor : POIOLE2TextExtractor
    {
        private HWPFOldDocument doc;

        /**
         * Create a new Word Extractor
         * @param is InputStream Containing the word file
         */
        public Word6Extractor(Stream is1)
            : this(new POIFSFileSystem(is1))
        {

        }

        /**
         * Create a new Word Extractor
         * @param fs POIFSFileSystem Containing the word file
         */
        public Word6Extractor(POIFSFileSystem fs)
            : this(fs.Root, fs)
        {

        }
        public Word6Extractor(DirectoryNode dir, POIFSFileSystem fs)
            : this(new HWPFOldDocument(dir, fs))
        {

        }

        /**
         * Create a new Word Extractor
         * @param doc The HWPFOldDocument to extract from
         */
        public Word6Extractor(HWPFOldDocument doc)
            : base(doc)
        {

            this.doc = doc;
        }

        /**
         * Get the text from the word file, as an array with one String
         *  per paragraph
         */
        public String[] ParagraphText
        {
            get
            {
                String[] ret;

                // Extract using the model code
                try
                {
                    Range r = doc.GetRange();

                    ret = WordExtractor.GetParagraphText(r);
                }
                catch (Exception)
                {
                    // Something's up with turning the text pieces into paragraphs
                    // Fall back to ripping out the text pieces
                    ret = new String[doc.TextTable.TextPieces.Count];
                    for (int i = 0; i < ret.Length; i++)
                    {
                        ret[i] = doc.TextTable.TextPieces[i].GetStringBuilder().ToString();

                        // Fix the line endings
                        ret[i].Replace("\r", "\ufffe");
                        ret[i].Replace("\ufffe", "\r\n");
                    }
                }

                return ret;
            }
        }

        public override String Text
        {
            get
            {
                StringBuilder text = new StringBuilder();

                foreach (String t in ParagraphText)
                {
                    text.Append(t);
                }

                return text.ToString();
            }
        }
    }
}
