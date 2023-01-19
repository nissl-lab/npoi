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

using NPOI.HSLF.Exceptions;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.HSLF.Record
{
	public class TextSpecInfoAtom: RecordAtom
	{
		private static long _type = RecordTypes.TextSpecInfoAtom.typeID;

		/**
		 * Record header.
		 */
		private byte[] _header;

		/**
		 * Record data.
		 */
		private byte[] _data;

		/**
		 * Constructs an empty atom, with a default run of size 1
		 */
		public TextSpecInfoAtom()
		{
			_header = new byte[8];
			LittleEndian.PutUInt(_header, 4, _type);
			Reset(1);
		}

		/**
		 * Constructs the link related atom record from its
		 *  source data.
		 *
		 * @param source the source data as a byte array.
		 * @param start the start offset into the byte array.
		 * @param len the length of the slice in the byte array.
		 */
		public TextSpecInfoAtom(byte[] source, int start, int len)
		{
			// Get the header.
			_header = Arrays.CopyOfRange(source, start, start + 8);

			// Get the record data.
			_data = IOUtils.SafelyClone(source, start + 8, len - 8, GetMaxRecordLength());
		}
		/**
		 * Gets the record type.
		 * @return the record type.
		 */
		//@Override
		public override long GetRecordType() { return _type; }

		/**
		 * Write the contents of the record back, so it can be written
		 * to disk
		 *
		 * @param out the output stream to write to.
		 * @throws java.io.IOException if an error occurs.
		 */
		//@Override
		public override void WriteOut(OutputStream _out)
		{
			_out.Write(_header);
			_out.Write(_data);
		}

		/**
		 * Update the text length
		 *
		 * @param size the text length
		 */
		public void SetTextSize(int size)
		{
			LittleEndian.PutInt(_data, 0, size);
		}

		/**
		 * Reset the content to one info run with the default values
		 * @param size  the site of parent text
		 */
		public void Reset(int size)
		{
			TextSpecInfoRun sir = new TextSpecInfoRun(size);
			using (MemoryStream bos = new MemoryStream())
			{
				try
				{
					sir.WriteOut((OutputStream)bos);
				}
				catch (IOException e)
				{
					throw new HSLFException(e);
				}
				_data = bos.ToArray();

			}
			// Update the size (header bytes 5-8)
			LittleEndian.PutInt(_header, 4, _data.Length);
		}

	/**
     * Adapts the size by enlarging the last {@link TextSpecInfoRun}
     * or chopping the runs to the given length
     */
	public void SetParentSize(int size)
	{
			//assert(size > 0);
			try
			{
				using (MemoryStream bos = new MemoryStream())
				{
					TextSpecInfoRun[] runs = GetTextSpecInfoRuns();
					int remaining = size;
					int idx = 0;
					foreach (TextSpecInfoRun run in runs)
					{
						int len = run.GetLength();
						if (len > remaining || idx == runs.Length - 1)
						{
							run.SetLength(len = remaining);
						}
						remaining -= len;
						run.WriteOut((OutputStream)bos);
						idx++;
					}

					_data = bos.ToArray();

					// Update the size (header bytes 5-8)
					LittleEndian.PutInt(_header, 4, _data.Length);
				}
			}
			catch (IOException e)
			{
				throw new HSLFException(e);
			}
	}

	/**
     * Get the number of characters covered by this records
     *
     * @return the number of characters covered by this records
     */
	public int GetCharactersCovered()
	{
		int covered = 0;
		foreach (TextSpecInfoRun r in GetTextSpecInfoRuns())
		{
			covered += r.GetLength();
		}
		return covered;
	}

	public TextSpecInfoRun[] GetTextSpecInfoRuns()
	{
		LittleEndianByteArrayInputStream bis = new LittleEndianByteArrayInputStream(_data); // NOSONAR
			List<TextSpecInfoRun> lst = new List<TextSpecInfoRun>();
		while (bis.GetReadIndex() < _data.Length)
		{
			lst.Add(new TextSpecInfoRun(bis));
		}
		return lst.ToArray();
	}

	//@Override
	public override IDictionary<string, Func<T>> GetGenericProperties<T>()
	{
		return (IDictionary<string, Func<T>>)GenericRecordUtil.GetGenericProperties(
			"charactersCovered", ()=> GetCharactersCovered(),
			"textSpecInfoRuns", GetTextSpecInfoRuns
		);
	}
}
}
