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
using System.Collections.Generic;
using System;
using System.IO;

namespace NPOI.HSLF.Record
{
    public class TextBytesAtom: RecordAtom
    {
		public static long _type = RecordTypes.TextBytesAtom.typeID;

		private byte[] _header;

		/** The bytes that make up the text */
		private byte[] _text;

		/** Grabs the text. Uses the default codepage */
		public string GetText()
		{
			return StringUtil.GetFromCompressedUnicode(_text, 0, _text.Length);
		}

		/** Updates the text in the Atom. Must be 8 bit ascii */
		public void SetText(byte[] b)
		{
			// Set the text
			_text = (byte[])b.Clone();

			// Update the size (header bytes 5-8)
			LittleEndian.PutInt(_header, 4, _text.Length);
		}

		/* *************** record code follows ********************** */

		/**
		 * For the TextBytes Atom
		 */
		protected TextBytesAtom(byte[] source, int start, int len)
		{
			// Sanity Checking
			if (len < 8) { len = 8; }

			// Get the header
			_header = Arrays.CopyOfRange(source, start, start + 8);

			// Grab the text
			_text = IOUtils.SafelyClone(source, start + 8, len - 8, GetMaxRecordLength());
		}

		/**
		 * Create an empty TextBytes Atom
		 */
		public TextBytesAtom()
		{
			_header = new byte[8];
			LittleEndian.PutUShort(_header, 0, 0);
			LittleEndian.PutUShort(_header, 2, (int)_type);
			LittleEndian.PutInt(_header, 4, 0);

			_text = new byte[] { };
		}

		/**
		 * We are of type 4008
		 */
		//@Override
		public override long GetRecordType() { return _type; }

		/**
		 * Write the contents of the record back, so it can be written
		 *  to disk
		 */
		//@Override
		public override void WriteOut(BinaryWriter _out)
		{
			// Header - size or type unchanged
			_out.Write(_header);

			// Write out our text
			_out.Write(_text);
		}

	/**
     * dump debug info; use getText() to return a string
     * representation of the atom
     */
	//@Override
	public override string ToString()
	{
		return GenericRecordJsonWriter.marshal(this);
	}

	//@Override
	public override IDictionary<string, Func<object>> GetGenericProperties()
	{
		return GenericRecordUtil.GetGenericProperties(
			"text", this.GetText
		);
	}
}
}