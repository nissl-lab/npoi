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


namespace NPOI.HWPF.Extractor
{
    using System.IO;
    using System;
    using NPOI.HWPF.UserModel;
    using System.Text;
    using NPOI.HWPF.Model;
    using System.Collections.Generic;
    using NPOI.POIFS.FileSystem;
    /**
     * Class to extract the text from a Word Document.
     *
     * You should use either GetParagraphText() or GetText() unless
     *  you have a strong reason otherwise.
     *
     * @author Nick Burch
     */
    public class WordExtractor : POIOLE2TextExtractor
    {
        private POIFSFileSystem fs;
        private HWPFDocument doc;

        /**
         * Create a new Word Extractor
         * @param is InputStream Containing the word file
         */
        public WordExtractor(Stream is1)
            : this(HWPFDocument.VerifyAndBuildPOIFS(is1))
        {

        }

        /**
         * Create a new Word Extractor
         * @param fs POIFSFileSystem Containing the word file
         */
        public WordExtractor(POIFSFileSystem fs)
            : this(new HWPFDocument(fs))
        {

            this.fs = fs;
        }
        public WordExtractor(DirectoryNode dir, POIFSFileSystem fs)
            : this(new HWPFDocument(dir, fs))
        {

            this.fs = fs;
        }

        /**
         * Create a new Word Extractor
         * @param doc The HWPFDocument to extract from
         */
        public WordExtractor(HWPFDocument doc)
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

                    ret = GetParagraphText(r);
                }
                catch (Exception)
                {
                    // Something's up with turning the text pieces into paragraphs
                    // Fall back to ripping out the text pieces
                    ret = new String[1];
                    ret[0] = this.TextFromPieces;
                }

                return ret;
            }
        }

        public String[] FootnoteText
        {
            get
            {
                Range r = doc.GetFootnoteRange();

                return GetParagraphText(r);
            }
        }
        public String[] MainTextboxText
        {
            get
            {
                Range r = doc.GetMainTextboxRange();

                return GetParagraphText(r);
            }
        }
        public String[] EndnoteText
        {
            get
            {
                Range r = doc.GetEndnoteRange();

                return GetParagraphText(r);
            }
        }

        public String[] CommentsText
        {
            get
            {
                Range r = doc.GetCommentsRange();

                return GetParagraphText(r);
            }
        }

        public static String[] GetParagraphText(Range r)
        {
            String[] ret;
            ret = new String[r.NumParagraphs];
            for (int i = 0; i < ret.Length; i++)
            {
                Paragraph p = r.GetParagraph(i);
                ret[i] = p.Text;

                // Fix the line ending
                if (ret[i].EndsWith("\r"))
                {
                    ret[i] = ret[i] + "\n";
                }
            }
            return ret;
        }

        /**
	 * Add the header/footer text, if it's not empty
	 */
        private void AppendHeaderFooter(String text, StringBuilder out1)
        {
            if (text == null || text.Length == 0)
                return;

            text = text.Replace('\r', '\n');
            if (!text.EndsWith("\n"))
            {
                out1.Append(text);
                out1.Append('\n');
                return;
            }
            if (text.EndsWith("\n\n"))
            {
                out1.Append(text.Substring(0, text.Length - 1));
                return;
            }
            out1.Append(text);
            return;
        }
        /**
         * Grab the text from the headers
         */
        public String HeaderText
        {
            get
            {
                HeaderStories hs = new HeaderStories(doc);

                StringBuilder ret = new StringBuilder();
                if (hs.FirstHeader != null)
                {
                    AppendHeaderFooter(hs.FirstHeader, ret);
                }
                if (hs.EvenHeader != null)
                {
                    AppendHeaderFooter(hs.EvenHeader, ret);
                }
                if (hs.OddHeader != null)
                {
                    AppendHeaderFooter(hs.OddHeader, ret);
                }

                return ret.ToString();
            }
        }
        /**
         * Grab the text from the footers
         */
        public String FooterText
        {
            get
            {
                HeaderStories hs = new HeaderStories(doc);

                StringBuilder ret = new StringBuilder();
                if (hs.FirstFooter != null)
                {
                    AppendHeaderFooter(hs.FirstFooter, ret);
                }
                if (hs.EvenFooter != null)
                {
                    AppendHeaderFooter(hs.EvenFooter, ret);
                }
                if (hs.OddFooter != null)
                {
                    AppendHeaderFooter(hs.OddFooter, ret);
                }

                return ret.ToString();
            }
        }

        /**
         * Grab the text out of the text pieces. Might also include various
         *  bits of crud, but will work in cases where the text piece -> paragraph
         *  mapping is broken. Fast too.
         */
        public String TextFromPieces
        {
            get
            {
                StringBuilder textBuf = new StringBuilder();

                foreach (TextPiece piece in doc.TextTable.TextPieces)
                {
                    String encoding = "Windows-1252";
                    if (piece.IsUnicode)
                    {
                        encoding = "UTF-16LE";
                    }
                    try
                    {
                        String text1 = Encoding.GetEncoding(encoding).GetString(piece.RawBytes);
                        textBuf.Append(text1);
                    }
                    catch (EncoderFallbackException)
                    {
                        throw new Exception("Standard Encoding " + encoding + " not found, JVM broken");
                    }
                }

                String text = textBuf.ToString();

                // Fix line endings (Note - won't get all of them
                text = text.Replace("\r\r\r", "\r\n\r\n\r\n");
                text = text.Replace("\r\r", "\r\n\r\n");

                if (text.EndsWith("\r"))
                {
                    text += "\n";
                }

                return text;
            }
        }

        /**
         * Grab the text, based on the paragraphs. Shouldn't include any crud,
         *  but slightly slower than GetTextFromPieces().
         */
        public override String Text
        {
            get
            {
                StringBuilder ret = new StringBuilder();

                ret.Append(HeaderText);

                List<String> text = new List<String>();
                text.AddRange(ParagraphText);
                text.AddRange(FootnoteText);
                text.AddRange(EndnoteText);

                foreach (String p in text)
                {
                    ret.Append(p);
                }

                ret.Append(FooterText);

                return ret.ToString();
            }
        }

        /**
         * Removes any fields (eg macros, page markers etc)
         *  from the string.
         */
        public static String StripFields(String text)
        {
            return Range.StripFields(text);
        }
    }
}