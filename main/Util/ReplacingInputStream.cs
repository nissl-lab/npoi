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
using System.Runtime.InteropServices;
using System.Text;

namespace NPOI.Util
{

    /// <summary>
    /// <para>
    /// Simple FilterInputStream that can replace occurrences of bytes with something else.
    /// </para>
    /// <para>
    /// This has been taken from inbot-utils. (MIT licensed)
    /// </para>
    /// </summary>
    /// @see <see href="https://github.com/Inbot/inbot-utils">inbot-utils</see>
    public class ReplacingInputStream : FilterInputStream
    {

        // while matching, this is where the bytes go.
        int[] buf;
        private int matchedIndex=0;
        private int unbufferIndex=0;
        private int replacedIndex=0;

        private  byte[] pattern;
        private  byte[] replacement;
        private State state=State.NOT_MATCHED;

        // simple state machine for keeping track of what we are doing
        private enum State
        {
            NOT_MATCHED,
            MATCHING,
            REPLACING,
            UNBUFFER
        }

        //private static  Charset UTF8 = Charset.forName("UTF-8");

        /// <summary>
        /// Replace occurrences of pattern in the input. Note: input is assumed to be UTF-8 encoded. If not the case use byte[] based pattern and replacement.
        /// </summary>
        /// <param name="in">input</param>
        /// <param name="pattern">pattern to replace.</param>
        /// <param name="replacement">the replacement or null</param>
        public ReplacingInputStream(InputStream in1, String pattern, String replacement)
            : this(in1, Encoding.UTF8.GetBytes(pattern), replacement==null ? null : Encoding.UTF8.GetBytes(replacement))
        {

        }

        /// <summary>
        /// <para>
        /// Replace occurrences of pattern in the input.
        /// </para>
        /// <para>
        /// If you want to normalize line endings DOS/MAC (\n\r | \r) to UNIX (\n), you can call the following:<br/>
        /// {@code new ReplacingInputStream(new ReplacingInputStream(is, "\n\r", "\n"), "\r", "\n")}
        /// </para>
        /// </summary>
        /// <param name="in">input</param>
        /// <param name="pattern">pattern to replace</param>
        /// <param name="replacement">the replacement or null</param>
        public ReplacingInputStream(InputStream in1, byte[] pattern, byte[] replacement)
            : base(in1)
        {
            ;
            if(pattern == null || pattern.Length == 0)
            {
                throw new ArgumentException("pattern length should be > 0");
            }
            this.pattern = pattern;
            this.replacement = replacement;
            // we will never match more than the pattern length
            buf = new int[pattern.Length];
        }
        public override int Read(byte[] b, int off, int len)
        {

            // copy of parent logic; we need to call our own read() instead of super.read(), which delegates instead of calling our read
            if(b == null)
            {
                throw new NullReferenceException();
            }
            else if(off < 0 || len < 0 || len > b.Length - off)
            {
                throw new IndexOutOfRangeException();
            }
            else if(len == 0)
            {
                return 0;
            }

            int c = Read();
            if(c == -1)
            {
                return -1;
            }
            b[off] = (byte) c;

            int i = 1;
            for(; i < len; i++)
            {
                c = Read();
                if(c == -1)
                {
                    break;
                }
                b[off + i] = (byte) c;
            }
            return i;

        }
        public override int Read(byte[] b)
        {

            // call our own read
            return Read(b, 0, b.Length);
        }
        public override int Read()
        {

            // use a simple state machine to figure out what we are doing
            int next;
            switch(state)
            {
                default:
                case State.NOT_MATCHED:
                    // we are not currently matching, replacing, or unbuffering
                    next=base.Read();
                    if(pattern[0] != next)
                    {
                        return next;
                    }

                    // clear whatever was there
                    Arrays.Fill(buf, 0);
                    // make sure we start at 0
                    matchedIndex=0;

                    buf[matchedIndex++]=next;
                    if(pattern.Length == 1)
                    {
                        // edge-case when the pattern length is 1 we go straight to replacing
                        state=State.REPLACING;
                        // reset replace counter
                        replacedIndex=0;
                    }
                    else
                    {
                        // pattern of length 1
                        state=State.MATCHING;
                    }
                    // recurse to continue matching
                    return Read();

                case State.MATCHING:
                    // the previous bytes matched part of the pattern
                    next=base.Read();
                    if(pattern[matchedIndex]==next)
                    {
                        buf[matchedIndex++]=next;
                        if(matchedIndex==pattern.Length)
                        {
                            // we've found a full match!
                            if(replacement==null || replacement.Length==0)
                            {
                                // the replacement is empty, go straight to NOT_MATCHED
                                state=State.NOT_MATCHED;
                                matchedIndex=0;
                            }
                            else
                            {
                                // start replacing
                                state=State.REPLACING;
                                replacedIndex=0;
                            }
                        }
                    }
                    else
                    {
                        // mismatch -> unbuffer
                        buf[matchedIndex++]=next;
                        state=State.UNBUFFER;
                        unbufferIndex=0;
                    }
                    return Read();

                case State.REPLACING:
                    // we've fully matched the pattern and are returning bytes from the replacement
                    next=replacement[replacedIndex++];
                    if(replacedIndex==replacement.Length)
                    {
                        state=State.NOT_MATCHED;
                        replacedIndex=0;
                    }
                    return next;

                case State.UNBUFFER:
                    // we partially matched the pattern before encountering a non matching byte
                    // we need to serve up the buffered bytes before we go back to NOT_MATCHED
                    next=buf[unbufferIndex++];
                    if(unbufferIndex==matchedIndex)
                    {
                        state=State.NOT_MATCHED;
                        matchedIndex=0;
                    }
                    return next;
            }
        }
        public override String ToString()
        {
            return state.ToString() + " " + matchedIndex + " " + replacedIndex + " " + unbufferIndex;
        }

    }
}

