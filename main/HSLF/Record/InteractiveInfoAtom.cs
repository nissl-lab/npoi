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

using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSLF.Record
{
	public class InteractiveInfoAtom: RecordAtom
	{
		public enum Action
		{
			NONE,
			MACRO,
			RUN_PROGRAM,
			JUMP,
			HYPERLINK,
			OLE,
			MEDIA,
			CUSTOM_SHOW
		}

		public enum Jump
		{
			NONE,
			NEXT_SLIDE,
			PREVIOUS_SLIDE,
			FIRST_SLIDE,
			LAST_SLIDE,
			LAST_SLIDE_VIEWED,
			END_SHOW
		}

		public enum Link
		{
			NEXT_SLIDE,
			PREVIOUS_SLIDE,
			FIRST_SLIDE,
			LAST_SLIDE,
			CUSTOM_SHOW,
			SLIDE_NUMBER,
			URL,
			OTHER_PRESENTATION,
			OTHER_FILE,
			NULL
		}

		/**
		 * Action Table
		 */
		public static  byte ACTION_NONE = 0;
		public static  byte ACTION_MACRO = 1;
		public static  byte ACTION_RUNPROGRAM = 2;
		public static  byte ACTION_JUMP = 3;
		public static  byte ACTION_HYPERLINK = 4;
		public static  byte ACTION_OLE = 5;
		public static  byte ACTION_MEDIA = 6;
		public static  byte ACTION_CUSTOMSHOW = 7;

		/**
		 *  Jump Table
		 */
		public static  byte JUMP_NONE = 0;
		public static  byte JUMP_NEXTSLIDE = 1;
		public static  byte JUMP_PREVIOUSSLIDE = 2;
		public static  byte JUMP_FIRSTSLIDE = 3;
		public static  byte JUMP_LASTSLIDE = 4;
		public static  byte JUMP_LASTSLIDEVIEWED = 5;
		public static  byte JUMP_ENDSHOW = 6;

		/**
		 * Types of hyperlinks
		 */
		public const byte LINK_NextSlide = 0x00;
		public const byte LINK_PreviousSlide = 0x01;
		public const byte LINK_FirstSlide = 0x02;
		public const byte LINK_LastSlide = 0x03;
		public const byte LINK_CustomShow = 0x06;
		public const byte LINK_SlideNumber = 0x07;
		public const byte LINK_Url = 0x08;
		public const byte LINK_OtherPresentation = 0x09;
		public const byte LINK_OtherFile = 0x0A;
		public const byte LINK_NULL = (byte)0xFF;

		private static  int[] FLAGS_MASKS = {
		0x0001, 0x0002, 0x0004, 0x0008
	};

		private static  String[] FLAGS_NAMES = {
        "ANIMATED", "STOP_SOUND", "CUSTOM_SHOW_RETURN", "VISITED"
    };

    /**
     * Record header.
     */
    private  byte[] _header;

		/**
		 * Record data.
		 */
		private  byte[] _data;

		/**
		 * Constructs a brand new link related atom record.
		 */
		public InteractiveInfoAtom()
		{
			_header = new byte[8];
			_data = new byte[16];

			LittleEndian.PutShort(_header, 2, (short)GetRecordType());
			LittleEndian.PutInt(_header, 4, _data.Length);

			// It is fine for the other values to be zero
		}

		/**
		 * Constructs the link related atom record from its
		 *  source data.
		 *
		 * @param source the source data as a byte array.
		 * @param start the start offset into the byte array.
		 * @param len the length of the slice in the byte array.
		 */
		protected InteractiveInfoAtom(byte[] source, int start, int len)
		{
			// Get the header.
			_header = Arrays.CopyOfRange(source, start, start + 8);

			// Get the record data.
			_data = IOUtils.SafelyClone(source, start + 8, len - 8, GetMaxRecordLength());

			// Must be at least 16 bytes long
			if (_data.Length < 16)
			{
				throw new InvalidOperationException("The length of the data for a InteractiveInfoAtom must be at least 16 bytes, but was only " + _data.Length);
			}

			// First 4 bytes - no idea, normally 0
			// Second 4 bytes - the id of the link (from 1 onwards)
			// Third 4 bytes - no idea, normally 4
			// Fourth 4 bytes - no idea, normally 8
		}

		/**
		 * Gets the link number. You will normally look the
		 *  ExHyperlink with this number to get the details.
		 * @return the link number
		 */
		public int GetHyperlinkID()
		{
			return LittleEndian.GetInt(_data, 4);
		}

		/**
		 * Sets the persistent unique identifier of the link
		 *
		 * @param number the persistent unique identifier of the link
		 */
		public void SetHyperlinkID(int number)
		{
			LittleEndian.PutInt(_data, 4, number);
		}

		/**
		 * a reference to a sound in the sound collection.
		 */
		public int GetSoundRef()
		{
			return LittleEndian.GetInt(_data, 0);
		}
		/**
		 * a reference to a sound in the sound collection.
		 *
		 * @param val a reference to a sound in the sound collection
		 */
		public void SetSoundRef(int val)
		{
			LittleEndian.PutInt(_data, 0, val);
		}

		/**
		 * Hyperlink Action.
		 * <p>
		 * see {@code ACTION_*} constants for the list of actions
		 * </p>
		 *
		 * @return hyperlink action.
		 */
		public byte GetAction()
		{
			return _data[8];
		}

		/**
		 * Hyperlink Action
		 * <p>
		 * see {@code ACTION_*} constants for the list of actions
		 * </p>
		 *
		 * @param val hyperlink action.
		 */
		public void SetAction(byte val)
		{
			_data[8] = val;
		}

		/**
		 * Only valid when action == OLEAction. OLE verb to use, 0 = first verb, 1 = second verb, etc.
		 */
		public byte GetOleVerb()
		{
			return _data[9];
		}

		/**
		 * Only valid when action == OLEAction. OLE verb to use, 0 = first verb, 1 = second verb, etc.
		 */
		public void SetOleVerb(byte val)
		{
			_data[9] = val;
		}

		/**
		 * Jump
		 * <p>
		 * see {@code JUMP_*} constants for the list of actions
		 * </p>
		 *
		 * @return jump
		 */
		public byte GetJump()
		{
			return _data[10];
		}

		/**
		 * Jump
		 * <p>
		 * see {@code JUMP_*} constants for the list of actions
		 * </p>
		 *
		 * @param val jump
		 */
		public void SetJump(byte val)
		{
			_data[10] = val;
		}

		/**
		 * Flags
		 * <ul>
		 * <li> Bit 1: Animated. If 1, then button is animated
		 * <li> Bit 2: Stop sound. If 1, then stop current sound when button is pressed.
		 * <li> Bit 3: CustomShowReturn. If 1, and this is a jump to custom show,
		 *   then return to this slide after custom show.
		 * </ul>
		 */
		public byte GetFlags()
		{
			return _data[11];
		}

		/**
		 * Flags
		 * <ul>
		 * <li> Bit 1: Animated. If 1, then button is animated
		 * <li> Bit 2: Stop sound. If 1, then stop current sound when button is pressed.
		 * <li> Bit 3: CustomShowReturn. If 1, and this is a jump to custom show,
		 *   then return to this slide after custom show.
		 * </ul>
		 */
		public void SetFlags(byte val)
		{
			_data[11] = val;
		}

		/**
		 * hyperlink type
		 *
		 * @return hyperlink type
		 */
		public byte GetHyperlinkType()
		{
			return _data[12];
		}

		/**
		 * hyperlink type
		 *
		 * @param val hyperlink type
		 */
		public void SetHyperlinkType(byte val)
		{
			_data[12] = val;
		}

		/**
		 * Gets the record type.
		 * @return the record type.
		 */
		//@Override
	public override long GetRecordType() { return RecordTypes.InteractiveInfoAtom.typeID; }

		/**
		 * Write the contents of the record back, so it can be written
		 * to disk
		 *
		 * @param out the output stream to write to.
		 * @throws IOException if an error occurs.
		 */
		//@Override
	public override void WriteOut(OutputStream _out)
		{
        _out.Write(_header);
        _out.Write(_data);
	}

	//@Override
	public override IDictionary<string, Func<T>> GetGenericProperties<T>()
	{
		return (IDictionary<string, Func<T>>) GenericRecordUtil.GetGenericProperties(
			"hyperlinkID", ()=>GetHyperlinkID(),
			"soundRef", ()=>GetSoundRef(),
			"action", ()=> GenericRecordUtil.SafeEnum<Action>((Action[])Enum.GetValues(typeof(Action)), GetAction, Action.NONE)(),
			"jump", ()=>GenericRecordUtil.SafeEnum<Jump>((Jump[])Enum.GetValues(typeof(Jump)), GetJump, Jump.NONE)(),
			"hyperlinkType", ()=>GenericRecordUtil.SafeEnum((Link[])Enum.GetValues(typeof(Link)), GetHyperlinkType, Link.NULL)(),
			"flags", GenericRecordUtil.GetBitsAsString(GetFlags, FLAGS_MASKS, FLAGS_NAMES)
		);
	}
}
}
