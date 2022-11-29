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
using System.IO;

namespace NPOI.HSLF.Record
{
	/**
 * The atom that holds the seed info used by a ExObjList
 */
	public class ExObjListAtom : RecordAtom
	{
		/**
     * Record header.
     */
		private byte[] _header;

		/**
		 * Record data.
		 */
		private byte[] _data;

		/**
		 * Constructs a brand new link related atom record.
		 */
		public ExObjListAtom()
		{
			_header = new byte[8];
			_data = new byte[4];

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
		protected ExObjListAtom(byte[] source, int start, int len)
		{
			// Get the header.
			_header = Arrays.CopyOfRange(source, start, start + 8);

			// Get the record data.
			_data = IOUtils.SafelyClone(source, start + 8, len - 8, GetMaxRecordLength());

			// Must be at least 4 bytes long
			if (_data.Length < 4)
			{
				throw new InvalidOperationException("The length of the data for a ExObjListAtom must be at least 4 bytes, but was only " + _data.Length);
			}
		}

		/**
		 * Gets the object ID seed, which will be used as the unique
		 *  OLE identifier for the next OLE object added
		 * @return the object ID seed
		 */
		public long GetObjectIDSeed()
		{
			return LittleEndian.GetUInt(_data, 0);
		}

		/**
		 * Sets the object ID seed
		 * @param seed the new ID seed
		 */
		public void SetObjectIDSeed(int seed)
		{
			LittleEndian.PutInt(_data, 0, seed);
		}

		/**
		 * Gets the record type.
		 * @return the record type.
		 */
		public override long GetRecordType() { return RecordTypes.ExObjListAtom.typeID; }

		/**
		 * Write the contents of the record back, so it can be written
		 * to disk
		 *
		 * @param out the output stream to write to.
		 * @throws IOException if an error occurs.
		 */
		public override void WriteOut(OutputStream _out)
		{
			_out.Write(_header);
			_out.Write(_data);
		}

		//@Override
		public override IDictionary<string, Func<T>> GetGenericProperties<T>()
		{
			return (IDictionary<string, Func<T>>)GenericRecordUtil.GetGenericProperties<long>(
				"objectIDSeed", this.GetObjectIDSeed
			);
		}
	}
}