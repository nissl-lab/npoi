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
    using System.Text;

    /**
     * Lightweight representation of a text piece.
     * Works in the character domain, not the byte domain, so you
     *  need to have turned byte references into character
     *  references before Getting here.
     *
     * @author Ryan Ackley
     */

    public class TextPiece : PropertyNode
    {
        private bool _usesUnicode;

        private PieceDescriptor _pd;

        /**
         * @param start Beginning offset in main document stream, in characters.
         * @param end Ending offset in main document stream, in characters.
         * @param text The raw bytes of our text
         */
        public TextPiece(int start, int end, byte[] text, PieceDescriptor pd, int cpStart)
            : this(start, end, text, pd)
        {
            
        }
        public TextPiece(int start, int end, byte[] text, PieceDescriptor pd)
            : base(start, end, buildInitSB(text, pd))
         {
            _usesUnicode = pd.IsUnicode;
            _pd = pd;

            // Validate
            int textLength = ((StringBuilder)_buf).Length;
            if (end - start != textLength)
            {
                throw new InvalidOperationException("Told we're for characters " + start + " -> " + end + ", but actually covers " + textLength + " characters!");
            }
            if (end < start)
            {
                throw new InvalidOperationException("Told we're of negative size! start=" + start + " end=" + end);
            }
        }

        /**
         * Create the StringBuilder from the text and unicode flag
         */
        private static StringBuilder buildInitSB(byte[] text, PieceDescriptor pd)
        {
            String str;
            try
            {
                if (pd.IsUnicode)
                {
                    str = Encoding.GetEncoding("UTF-16LE").GetString(text);
                }
                else
                {
                    //str = Encoding.GetEncoding("CP1252").GetString(text);
                    str = Encoding.GetEncoding("Windows-1252").GetString(text);
                }
            }
            catch (EncoderFallbackException)
            {
                throw new Exception("Your Java is broken! It doesn't know about basic, required character encodings!");
            }
            return new StringBuilder(str);
        }

        /**
         * @return If this text piece is unicode
         */
        public bool IsUnicode
        {
            get
            {
                return _usesUnicode;
            }
        }

        public PieceDescriptor PieceDescriptor
        {
            get
            {
                return _pd;
            }
        }

        public StringBuilder GetStringBuilder()
        {
            return (StringBuilder)_buf;
        }

        public byte[] RawBytes
        {
            get
            {
                return Encoding.GetEncoding(_usesUnicode ? "UTF-16LE" : "Windows-1252").GetBytes(_buf.ToString());
            }
        }

        /**
         * Returns part of the string.
         * Works only in characters, not in bytes!
         * @param start Local start position, in characters
         * @param end Local end position, in characters
         */
        public String Substring(int start, int end)
        {
            StringBuilder buf = (StringBuilder)_buf;

            // Validate
            if (start < 0)
            {
                throw new IndexOutOfRangeException("Can't request a substring before 0 - asked for " + start);
            }
            if (end > buf.Length)
            {
                throw new IndexOutOfRangeException("Index " + end + " out of range 0 -> " + buf.Length);
            }
            if (end < start)
            {
                throw new IndexOutOfRangeException("Asked for text from " + start + " to " + end + ", which has an end before the start!");
            }
            return buf.ToString().Substring(start, end - start);
        }

        /**
         * Adjusts the internal string for deletinging
         *  some characters within this.
         * @param start The start position for the delete, in characters
         * @param length The number of characters to delete
         */
        public override void AdjustForDelete(int start, int Length)
        {
            int numChars = Length;

            int myStart = Start;
            int myEnd = End;
            int end = start + numChars;

            /* do we have to delete from this text piece? */
            if (start <= myEnd && end >= myStart)
            {

                /* find where the deleted area overlaps with this text piece */
                int overlapStart = Math.Max(myStart, start);
                int overlapEnd = Math.Min(myEnd, end);
                ((StringBuilder)_buf).Remove(overlapStart, overlapEnd - overlapStart);
            }

            // We need to invoke this even if text from this piece is not being
            // deleted because the adjustment must propagate to all subsequent
            // text pieces i.e., if text from tp[n] is being deleted, then
            // tp[n + 1], tp[n + 2], etc. will need to be adjusted.
            // The superclass is expected to use a separate sentry for this.
            base.AdjustForDelete(start, Length);
        }

        /**
         * Returns the Length, in characters
         */
        public virtual int CharacterLength
        {
            get
            {
                return (End - Start);
            }
        }
        /**
         * Returns the Length, in bytes
         */
        public virtual int BytesLength
        {
            get
            {
                return (End - Start) * (_usesUnicode ? 2 : 1);
            }
        }

        public override bool Equals(Object o)
        {
            if (LimitsAreEqual(o))
            {
                TextPiece tp = (TextPiece)o;
                return GetStringBuilder().ToString().Equals(tp.GetStringBuilder().ToString()) &&
                       tp._usesUnicode == _usesUnicode && _pd.Equals(tp._pd);
            }
            return false;
        }


        /**
         * Returns the character position we start at.
         */
        public virtual int GetCP()
        {
            return Start;
        }

        public override String ToString()
        {
            return "TextPiece from " + Start + " to " + End + " ("
                    + PieceDescriptor + ")";
        }
    }

}
