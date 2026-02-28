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

namespace NPOI.HSLF.Record
{
    using System;
    using System.IO;
    using NPOI.Util;

    /**
     * A Container record that specifies information about animation information for a shape.
     *
     * @author Yegor Kozlov
     */
    public class AnimationInfo : RecordContainer
    {
        private byte[] _header;

        // Links to our more interesting children
        private AnimationInfoAtom animationAtom;

        /**
         * Set things up, and find our more interesting children
         */
        protected AnimationInfo(byte[] source, int start, int len)
        {
            // Grab the header
            _header = new byte[8];
            Array.Copy(source, start, _header, 0, 8);

            // Find our children
            _children = Record.FindChildRecords(source, start + 8, len - 8);
            FindInterestingChildren();
        }

        /**
         * Go through our child records, picking out the ones that are
         *  interesting, and saving those for use by the easy helper
         *  methods.
         */
        private void FindInterestingChildren()
        {

            // First child should be the ExMediaAtom
            if (_children[0] is AnimationInfoAtom)
            {
                animationAtom = (AnimationInfoAtom)_children[0];
            }
        }

        /**
         * Create a new AnimationInfo, with blank fields
         */
        public AnimationInfo()
        {
            // Setup our header block
            _header = new byte[8];
            _header[0] = 0x0f; // We are a Container record
            LittleEndian.PutShort(_header, 2, (short)RecordType);

            _children = new Record[1];
            _children[0] = animationAtom = new AnimationInfoAtom();
        }

        /**
         * We are of type 4103
         */
        public override long RecordType
        {
            get { return RecordTypes.AnimationInfo.typeID; }
        }

        /**
         * Write the contents of the record back, so it can be written
         *  to disk
         */
        public override void WriteOut(Stream out1)
        {
            WriteOut(_header[0], _header[1], RecordType, _children, out1);
        }

        /**
         * Returns the AnimationInfo
         */
        public AnimationInfoAtom GetAnimationInfoAtom()
        {
            return animationAtom;
        }

    }
}

