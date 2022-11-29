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

namespace NPOI.HSLF.Record
{
	/**
 * A Document Atom (type 1001). Holds misc information on the PowerPoint
 * document, lots of them size and scale related.
 */
	public class DocumentAtom : RecordAtom
	{
		/**
     * Holds the different Slide Size values
     */
		public enum SlideSize
		{
			/** Slide size ratio is consistent with a computer screen. */
			ON_SCREEN,
			/** Slide size ratio is consistent with letter paper. */
			LETTER_SIZED_PAPER,
			/** Slide size ratio is consistent with A4 paper. */
			A4_SIZED_PAPER,
			/** Slide size ratio is consistent with 35mm photo slides. */
			ON_35MM,
			/** Slide size ratio is consistent with overhead projector slides. */
			OVERHEAD,
			/** Slide size ratio is consistent with a banner. */
			BANNER,
			/**
			 * Slide size ratio that is not consistent with any of the other specified slide sizes in
			 * this enumeration.
			 */
			CUSTOM
		}


		private byte[] _header = new byte[8];
		private static long _type = RecordTypes.DocumentAtom.typeID;

		private long slideSizeX; // PointAtom, assume 1st 4 bytes = X
		private long slideSizeY; // PointAtom, assume 2nd 4 bytes = Y
		private long notesSizeX; // PointAtom, assume 1st 4 bytes = X
		private long notesSizeY; // PointAtom, assume 2nd 4 bytes = Y
		private long serverZoomFrom; // RatioAtom, assume 1st 4 bytes = from
		private long serverZoomTo;   // RatioAtom, assume 2nd 4 bytes = to

		private long notesMasterPersist; // ref to NotesMaster, 0 if none
		private long handoutMasterPersist; // ref to HandoutMaster, 0 if none

		private int firstSlideNum;
		private int slideSizeType; // see DocumentAtom.SlideSize

		private byte saveWithFonts;
		private byte omitTitlePlace;
		private byte rightToLeft;
		private byte showComments;

		private byte[] reserved;


		public long GetSlideSizeX() { return slideSizeX; }
		public long GetSlideSizeY() { return slideSizeY; }
		public long GetNotesSizeX() { return notesSizeX; }
		public long GetNotesSizeY() { return notesSizeY; }
		public void SetSlideSizeX(long x) { slideSizeX = x; }
		public void SetSlideSizeY(long y) { slideSizeY = y; }
		public void SetNotesSizeX(long x) { notesSizeX = x; }
		public void SetNotesSizeY(long y) { notesSizeY = y; }

		public long GetServerZoomFrom() { return serverZoomFrom; }
		public long GetServerZoomTo() { return serverZoomTo; }
		public void GetServerZoomFrom(long zoom) { serverZoomFrom = zoom; }
		public void SetServerZoomTo(long zoom) { serverZoomTo = zoom; }

		/** Returns a reference to the NotesMaster, or 0 if none */
		public long GetNotesMasterPersist() { return notesMasterPersist; }
		/** Returns a reference to the HandoutMaster, or 0 if none */
		public long GetHandoutMasterPersist() { return handoutMasterPersist; }

		public long GetFirstSlideNum() { return firstSlideNum; }

		/**
		 * The Size of the Document's slides, {@link DocumentAtom.SlideSize} for values.
		 */
		public SlideSize GetSlideSizeType() { return (SlideSize)slideSizeType; }

		//	/**
		//	 * The Size of the Document's slides, {@link DocumentAtom.SlideSize} for values.
		//	 * @deprecated replaced by {@link #getSlideSizeType()}
		//	 */
		//	@Deprecated
		//	@Removal(version = "6.0.0")

		public SlideSize GetSlideSizeTypeEnum()
			{
				return (SlideSize)slideSizeType;
			}

		public void SetSlideSize(SlideSize size)
		{
			slideSizeType = (int)size;
		}

		/** Was the document saved with True Type fonts embedded? */
		public bool GetSaveWithFonts()
		{
			return saveWithFonts != 0;
		}

		/** Set the font embedding state */
		public void SetSaveWithFonts(bool saveWithFonts)
		{
			this.saveWithFonts = (byte)(saveWithFonts ? 1 : 0);
		}

		/** Have the placeholders on the title slide been omitted? */
		public bool GetOmitTitlePlace()
		{
			return omitTitlePlace != 0;
		}

		/** Is this a Bi-Directional PPT Doc? */
		public bool GetRightToLeft()
		{
			return rightToLeft != 0;
		}

		/** Are comment shapes visible? */
		public bool GetShowComments()
		{
			return showComments != 0;
		}


		/* *************** record code follows ********************** */

		/**
		 * For the Document Atom
		 */
		/* package */
		public DocumentAtom(byte[] source, int start, int len)
		{
			int maxLen = Math.Max(len, 48);
			LittleEndianByteArrayInputStream leis =
				new LittleEndianByteArrayInputStream(source, start, maxLen);

			// Get the header
			leis.ReadFully(_header);

			// Get the sizes and zoom ratios
			slideSizeX = leis.ReadInt();
			slideSizeY = leis.ReadInt();
			notesSizeX = leis.ReadInt();
			notesSizeY = leis.ReadInt();
			serverZoomFrom = leis.ReadInt();
			serverZoomTo = leis.ReadInt();

			// Get the master persists
			notesMasterPersist = leis.ReadInt();
			handoutMasterPersist = leis.ReadInt();

			// Get the ID of the first slide
			firstSlideNum = leis.ReadShort();

			// Get the slide size type
			slideSizeType = leis.ReadShort();

			// Get the bools as bytes
			saveWithFonts = (byte)leis.ReadByte();
			omitTitlePlace = (byte)leis.ReadByte();
			rightToLeft = (byte)leis.ReadByte();
			showComments = (byte)leis.ReadByte();

			// If there's any other bits of data, keep them about
			reserved = IOUtils.SafelyAllocate(maxLen - 48L, GetMaxRecordLength());
			leis.ReadFully(reserved);
		}

		/**
		 * We are of type 1001
		 */
		//@Override
		public override long GetRecordType() { return _type; }

		/**
		 * Write the contents of the record back, so it can be written
		 *  to disk
		 */
		//@Override
		public override void WriteOut(OutputStream _out)
		{
			// Header
			_out.Write(_header);

			// The sizes and zoom ratios
			WriteLittleEndian((int)slideSizeX, _out);
			WriteLittleEndian((int)slideSizeY, _out);
			WriteLittleEndian((int)notesSizeX, _out);
			WriteLittleEndian((int)notesSizeY, _out);
			WriteLittleEndian((int)serverZoomFrom, _out);
			WriteLittleEndian((int)serverZoomTo, _out);

			// The master persists
			WriteLittleEndian((int)notesMasterPersist, _out);
			WriteLittleEndian((int)handoutMasterPersist, _out);

			// The ID of the first slide
			WriteLittleEndian((short)firstSlideNum, _out);

			// The slide size type
			WriteLittleEndian((short)slideSizeType, _out);

			// The bools as bytes
			_out.Write(saveWithFonts);
			_out.Write(omitTitlePlace);
			_out.Write(rightToLeft);
			_out.Write(showComments);

			// Reserved data
			_out.Write(reserved);
		}

		//@Override
		public override IDictionary<string, Func<T>> GetGenericProperties<T>()
		{
			IDictionary<string, Func<object>> m = new Dictionary<string, Func<object>>() { };
			m.Add("slideSizeX", () => GetSlideSizeX());
			m.Add("slideSizeY", () => GetSlideSizeY());
			m.Add("notesSizeX", () => GetNotesSizeX());
			m.Add("notesSizeY", () => GetNotesSizeY());
			m.Add("serverZoomFrom", () => GetServerZoomFrom());
			m.Add("serverZoomTo", () => GetServerZoomTo());
			m.Add("notesMasterPersist", () => GetNotesMasterPersist());
			m.Add("handoutMasterPersist", () => GetHandoutMasterPersist());
			m.Add("firstSlideNum", () => GetFirstSlideNum());
			m.Add("slideSize", () => GetSlideSizeTypeEnum());
			m.Add("saveWithFonts", () => GetSaveWithFonts());
			m.Add("omitTitlePlace", () => GetOmitTitlePlace());
			m.Add("rightToLeft", () => GetRightToLeft());
			m.Add("showComments", () => GetShowComments());

			return (IDictionary<string, Func<T>>)m;
		}
	}
}