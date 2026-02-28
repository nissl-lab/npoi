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
    using System.Text;

    /**
     * An atom record that specifies the animation information for a shape.
     *
     * @author Yegor Kozlov
     */
    public class AnimationInfoAtom : RecordAtom
    {

        /**
         * whether the animation plays in the reverse direction
         */
        public static int Reverse = 1;
        /**
         * whether the animation starts automatically
         */
        public static int Automatic = 4;
        /**
         * whether the animation has an associated sound
         */
        public static int Sound = 16;
        /**
         * whether all playing sounds are stopped when this animation begins
         */
        public static int StopSound = 64;
        /**
         * whether an associated sound, media or action verb is activated when the shape is clicked.
         */
        public static int Play = 256;
        /**
         * specifies that the animation, while playing, stops other slide show actions.
         */
        public static int Synchronous = 1024;
        /**
         * whether the shape is hidden while the animation is not playing
         */
        public static int Hide = 4096;
        /**
         * whether the background of the shape is animated
         */
        public static int AnimateBg = 16384;

        /**
         * Record header.
         */
        private byte[] _header;

        /**
         * record data
         */
        private byte[] _recdata;

        /**
         * Constructs a brand new link related atom record.
         */
        public AnimationInfoAtom()
        {
            _recdata = new byte[28];

            _header = new byte[8];
            LittleEndian.PutShort(_header, 0, (short)0x01);
            LittleEndian.PutShort(_header, 2, (short)RecordType);
            LittleEndian.PutInt(_header, 4, _recdata.Length);
        }

        /**
         * Constructs the link related atom record from its
         *  source data.
         *
         * @param source the source data as a byte array.
         * @param start the start offset into the byte array.
         * @param len the length of the slice in the byte array.
         */
        public AnimationInfoAtom(byte[] source, int start, int len)
        {
            // Get the header
            _header = new byte[8];
            Array.Copy(source, start, _header, 0, 8);

            // Grab the record data
            _recdata = new byte[len - 8];
            Array.Copy(source, start + 8, _recdata, 0, len - 8);
        }

        /**
         * Gets the record type.
         * @return the record type.
         */
        public override long RecordType
        {
            get
            {
                return RecordTypes.AnimationInfoAtom.typeID;
            }
        }

        /**
         * Write the contents of the record back, so it can be written
         * to disk
         *
         * @param out the output stream to write to.
         * @throws java.io.IOException if an error occurs.
         */
        public override void WriteOut(Stream out1)
        {
            out1.Write(_header,(int)out1.Position,_header.Length);
            out1.Write(_recdata, (int)out1.Position, _recdata.Length);
        }

        /**
         * A rgb structure that specifies a color for the dim effect after the animation is complete.
         *
         * @return  color for the dim effect after the animation is complete
         */
        public int DimColor
        {
            get
            {
                return LittleEndian.GetInt(_recdata, 0);
            }
            set 
            {
                LittleEndian.PutInt(_recdata, 0, value);
            }
        }
        /**
         *  A bit mask specifying options for displaying headers and footers
         *
         * @return A bit mask specifying options for displaying headers and footers
         */
        public int Mask
        {
            get
            {
                return LittleEndian.GetInt(_recdata, 4);
            }
            set 
            {
                LittleEndian.PutInt(_recdata, 4,value);
            }
        }

        /**
         * @param bit the bit to check
         * @return whether the specified flag is set
         */
        public bool GetFlag(int bit)
        {
            return (Mask & bit) != 0;
        }

        /**
         * @param  bit the bit to set
         * @param  value whether the specified bit is set
         */
        public void SetFlag(int bit, bool value)
        {
            int mask = Mask;
            if (value) mask |= bit;
            else mask &= ~bit;
            Mask = (mask);
        }

        /**
         * A 4-byte unsigned integer that specifies a reference to a sound
         * in the SoundCollectionContainer record to locate the embedded audio
         *
         * @return  reference to a sound
         */
        public int SoundIdRef
        {
            get
            {
                return LittleEndian.GetInt(_recdata, 8);
            }
            set
            {
                LittleEndian.PutInt(_recdata, 8, value);
            }
        }

        /**
         * A signed integer that specifies the delay time, in milliseconds, before the animation starts to play.
         * If {@link #Automatic} is 0x1, this value MUST be greater than or equal to 0; otherwise, this field MUST be ignored.
         */
        public int DelayTime
        {
            get
            {
                return LittleEndian.GetInt(_recdata, 12);
            }
            set 
            {
                LittleEndian.PutInt(_recdata, 12, value);
            }
        }

        /**
         * A signed integer that specifies the order of the animation in the slide.
         * It MUST be greater than or equal to -2. The value -2 specifies that this animation follows the order of
         * the corresponding placeholder shape on the main master slide or title master slide.
         * The value -1 SHOULD NOT <105> be used.
         */
        public int OrderID
        {
            get
            {
                return LittleEndian.GetInt(_recdata, 16);
            }
            set 
            {
                LittleEndian.PutInt(_recdata, 16, value);
            }
        }

        /**
         * An unsigned integer that specifies the number of slides that this animation continues playing.
         * This field is utilized only in conjunction with media.
         * The value 0xFFFFFFFF specifies that the animation plays for one slide.
         */
        public int SlideCount
        {
            get
            {
                return LittleEndian.GetInt(_recdata, 18);
            }
            set 
            {
                LittleEndian.PutInt(_recdata, 18, value);
            }
        }

        public override String ToString()
        {
            StringBuilder buf = new StringBuilder();
            buf.Append("AnimationInfoAtom\n");
            buf.Append("\tDimColor: " + DimColor + "\n");
            int mask = Mask;
            buf.Append("\tMask: " + mask + ", 0x" + StringUtil.ToHexString(mask) + "\n");
            buf.Append("\t  Reverse: " + GetFlag(Reverse) + "\n");
            buf.Append("\t  Automatic: " + GetFlag(Automatic) + "\n");
            buf.Append("\t  Sound: " + GetFlag(Sound) + "\n");
            buf.Append("\t  StopSound: " + GetFlag(StopSound) + "\n");
            buf.Append("\t  Play: " + GetFlag(Play) + "\n");
            buf.Append("\t  Synchronous: " + GetFlag(Synchronous) + "\n");
            buf.Append("\t  Hide: " + GetFlag(Hide) + "\n");
            buf.Append("\t  AnimateBg: " + GetFlag(AnimateBg) + "\n");
            buf.Append("\tSoundIdRef: " + SoundIdRef + "\n");
            buf.Append("\tDelayTime: " + DelayTime + "\n");
            buf.Append("\tOrderID: " + OrderID + "\n");
            buf.Append("\tSlideCount: " + SlideCount + "\n");
            return buf.ToString();
        }

    }

}
