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
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using NPOI.POIFS.Common;
    using NPOI.HWPF.Model.IO;
    using System.Text;

    /**
     * The piece table for matching up character positions to bits of text. This
     * mostly works in bytes, but the TextPieces themselves work in characters. This
     * does the icky Convertion.
     *
     * @author Ryan Ackley
     */
    public class TextPieceTable : CharIndexTranslator
    {
        protected List<TextPiece> _textPieces = new List<TextPiece>();
        protected List<TextPiece> _textPiecesFCOrder = new List<TextPiece>();
        // int _multiple;
        int _cpMin;

        public TextPieceTable()
        {
        }

        public TextPieceTable(byte[] documentStream, byte[] tableStream, int offset, int size, int fcMin)
        {
            // get our plex of PieceDescriptors
            PlexOfCps pieceTable = new PlexOfCps(tableStream, offset, size, PieceDescriptor
                    .SizeInBytes);

            int length = pieceTable.Length;
            PieceDescriptor[] pieces = new PieceDescriptor[length];

            // iterate through piece descriptors raw bytes and create
            // PieceDescriptor objects
            for (int x = 0; x < length; x++)
            {
                GenericPropertyNode node = pieceTable.GetProperty(x);
                pieces[x] = new PieceDescriptor(node.Bytes, 0);
            }

            // Figure out the cp of the earliest text piece
            // Note that text pieces don't have to be stored in order!
            _cpMin = pieces[0].FilePosition - fcMin;
            for (int x = 0; x < pieces.Length; x++)
            {
                int start = pieces[x].FilePosition - fcMin;
                if (start < _cpMin)
                {
                    _cpMin = start;
                }
            }

            // using the PieceDescriptors, build our list of TextPieces.
            for (int x = 0; x < pieces.Length; x++)
            {
                int start = pieces[x].FilePosition;
                PropertyNode node = pieceTable.GetProperty(x);

                // Grab the start and end, which are in characters
                int nodeStartChars = node.Start;
                int nodeEndChars = node.End;

                // What's the relationship between bytes and characters?
                bool unicode = pieces[x].IsUnicode;
                int multiple = 1;
                if (unicode)
                {
                    multiple = 2;
                }

                // Figure out the Length, in bytes and chars
                int textSizeChars = (nodeEndChars - nodeStartChars);
                int textSizeBytes = textSizeChars * multiple;

                // Grab the data that Makes up the piece
                byte[] buf = new byte[textSizeBytes];
                Array.Copy(documentStream, start, buf, 0, textSizeBytes);

                // And now build the piece
                _textPieces.Add(new TextPiece(nodeStartChars, nodeEndChars, buf, pieces[x], node
                        .Start));
            }

            // In the interest of our sanity, now sort the text pieces
            // into order, if they're not already
            _textPieces.Sort();
            _textPiecesFCOrder = new List<TextPiece>(_textPieces);
            _textPiecesFCOrder.Sort(new FCComparator());
        }

        public int CpMin
        {
            get
            {
                return _cpMin;
            }
        }

        public List<TextPiece> TextPieces
        {
            get
            {
                return _textPieces;
            }
        }

        public void Add(TextPiece piece)
        {
            _textPieces.Add(piece);
            _textPiecesFCOrder.Add(piece);
            _textPieces.Sort();
            _textPiecesFCOrder.Sort(new FCComparator());
        }

        public byte[] WriteTo(HWPFStream docStream)
        {

            PlexOfCps textPlex = new PlexOfCps(PieceDescriptor.SizeInBytes);
            // int fcMin = docStream.Getoffset();

            int size = _textPieces.Count;
            for (int x = 0; x < size; x++)
            {
                TextPiece next = _textPieces[x];
                PieceDescriptor pd = next.PieceDescriptor;

                int offset = docStream.Offset;
                int mod = (offset % POIFSConstants.SMALLER_BIG_BLOCK_SIZE);
                if (mod != 0)
                {
                    mod = POIFSConstants.SMALLER_BIG_BLOCK_SIZE - mod;
                    byte[] buf = new byte[mod];
                    docStream.Write(buf);
                }

                // set the text piece position to the current docStream offset.
                pd.FilePosition = (docStream.Offset);

                // write the text to the docstream and save the piece descriptor to
                // the
                // plex which will be written later to the tableStream.
                docStream.Write(next.RawBytes);

                // The TextPiece is already in characters, which
                // Makes our life much easier
                int nodeStart = next.Start;
                int nodeEnd = next.End;
                textPlex.AddProperty(new GenericPropertyNode(nodeStart, nodeEnd, pd.ToByteArray()));
            }

            return textPlex.ToByteArray();

        }

        /**
         * Adjust all the text piece after inserting some text into one of them
         *
         * @param listIndex
         *            The TextPiece that had characters inserted into
         * @param length
         *            The number of characters inserted
         */
        public int AdjustForInsert(int listIndex, int length)
        {
            int size = _textPieces.Count;

            TextPiece tp = _textPieces[listIndex];

            // Update with the new end
            tp.End = (tp.End + length);

            // Now change all subsequent ones
            for (int x = listIndex + 1; x < size; x++)
            {
                tp = (TextPiece)_textPieces[x];
                tp.Start = (tp.Start + length);
                tp.End = (tp.End + length);
            }

            // All done
            return length;
        }

        public override bool Equals(Object o)
        {
            TextPieceTable tpt = (TextPieceTable)o;

            int size = tpt._textPieces.Count;
            if (size == _textPieces.Count)
            {
                for (int x = 0; x < size; x++)
                {
                    if (!tpt._textPieces[x].Equals(_textPieces[x]))
                    {
                        return false;
                    }
                }
                return true;
            }
            return false;
        }

        public int GetByteIndex(int charPos)
        {
            int byteCount = 0;
            foreach (TextPiece tp in _textPieces)
            {
                if (charPos >= tp.End)
                {
                    byteCount = tp.PieceDescriptor.FilePosition
                            + (tp.End - tp.Start)
                            * (tp.IsUnicode ? 2 : 1);

                    if (charPos == tp.End)
                        break;

                    continue;
                }
                if (charPos < tp.End)
                {
                    int left = charPos - tp.Start;
                    byteCount = tp.PieceDescriptor.FilePosition + left
                            * (tp.IsUnicode ? 2 : 1);
                    break;
                }
            }
            return byteCount;
        }

        public int GetCharIndex(int bytePos)
        {
            return GetCharIndex(bytePos, 0);
        }

        public int GetCharIndex(int bytePos, int startCP)
        {
            int charCount = 0;

            bytePos = LookIndexForward(bytePos);

            foreach (TextPiece tp in _textPieces)
            {
                int pieceStart = tp.PieceDescriptor.FilePosition;

                int bytesLength = tp.BytesLength;
                int pieceEnd = pieceStart + bytesLength;

                int toAdd;

                if (bytePos < pieceStart || bytePos > pieceEnd)
                {
                    toAdd = bytesLength;
                }
                else if (bytePos > pieceStart && bytePos < pieceEnd)
                {
                    toAdd = (bytePos - pieceStart);
                }
                else
                {
                    toAdd = bytesLength - (pieceEnd - bytePos);
                }

                if (tp.IsUnicode)
                {
                    charCount += toAdd / 2;
                }
                else
                {
                    charCount += toAdd;
                }

                if (bytePos >= pieceStart && bytePos <= pieceEnd && charCount >= startCP)
                {
                    break;
                }
            }

            return charCount;
        }

        public int LookIndexForward(int bytePos)
        {
            foreach (TextPiece tp in _textPiecesFCOrder)
            {
                int pieceStart = tp.PieceDescriptor.FilePosition;

                if (bytePos > pieceStart + tp.BytesLength)
                {
                    continue;
                }

                if (pieceStart > bytePos)
                {
                    bytePos = pieceStart;
                }

                break;
            }
            return bytePos;
        }

        public int LookIndexBackward(int bytePos)
        {
            int lastEnd = 0;

            foreach (TextPiece tp in _textPiecesFCOrder)
            {
                int pieceStart = tp.PieceDescriptor.FilePosition;

                if (bytePos > pieceStart + tp.BytesLength)
                {
                    lastEnd = pieceStart + tp.BytesLength;
                    continue;
                }

                if (pieceStart > bytePos)
                {
                    bytePos = lastEnd;
                }

                break;
            }

            return bytePos;
        }

        public virtual bool IsIndexInTable(int bytePos)
        {
            foreach (TextPiece tp in _textPiecesFCOrder)
            {
                int pieceStart = tp.PieceDescriptor.FilePosition;

                if (bytePos > pieceStart + tp.BytesLength)
                {
                    continue;
                }

                if (pieceStart > bytePos)
                {
                    return false;
                }

                return true;
            }

            return false;
        }

        public StringBuilder Text
        {
            get
            {
                long start = DateTime.Now.Ticks;

                // rebuild document paragraphs structure
                StringBuilder docText = new StringBuilder();
                foreach (TextPiece textPiece in _textPieces)
                {
                    String toAppend = textPiece.GetStringBuilder().ToString();
                    int toAppendLength = toAppend.Length;

                    //if ( toAppendLength != textPiece.getEnd() - textPiece.getStart() )
                    //{
                    //}
                    docText.Append(toAppend);
                }

                return docText;
            }
        }

        private class FCComparator : IComparer<TextPiece>
        {
            public int Compare(TextPiece textPiece, TextPiece textPiece1)
            {
                if (textPiece.PieceDescriptor.FilePosition > textPiece1.PieceDescriptor.FilePosition)
                {
                    return 1;
                }
                else if (textPiece.PieceDescriptor.FilePosition < textPiece1.PieceDescriptor.FilePosition)
                {
                    return -1;
                }
                else
                {
                    return 0;
                }
            }
        }
    }
}


