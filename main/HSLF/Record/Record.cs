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

using NPOI.Common.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.Util;
using System.Collections.Generic;
using System;
using NPOI.HSLF.Exceptions;
using System.Linq;
using System.IO;

namespace NPOI.HSLF.Record
{
	/**
	 * This abstract class represents a record in the PowerPoint document.
	 * Record classes should extend with RecordContainer or RecordAtom, which
	 *  extend this in turn.
	 */
	public abstract class Record: GenericRecord
	{
		/**
     * Is this record type an Atom record (only has data),
     *  or is it a non-Atom record (has other records)?
     */
		public abstract bool IsAnAtom();

		/**
		 * Returns the type (held as a little endian in bytes 3 and 4)
		 *  that this class handles
		 */
		public abstract long GetRecordType();

		/**
		 * Fetch all the child records of this record
		 * If this record is an atom, will return null
		 * If this record is a non-atom, but has no children, will return
		 *  an empty array
		 */
		public abstract Record[] GetChildRecords();

		/**
		 * Have the contents printer out into an OutputStream, used when
		 *  writing a file back out to disk
		 * (Normally, atom classes will keep their bytes around, but
		 *  non atom classes will just request the bytes from their
		 *  children, then chuck on their header and return)
		 */
		public abstract void WriteOut(OutputStream o);

		//@Override
		public RecordTypes GetGenericRecordType()
		{
			return RecordTypes.ForTypeID((int)GetRecordType());
		}

		//@Override
		public IList<GenericRecord> GetGenericChildren()
		{
			Record[] recs = GetChildRecords();
			return (IList<GenericRecord>)((recs == null) ? null : Arrays.AsList(recs));
		}

		/**
		 * When writing out, write out a signed int (32bit) in Little Endian format
		 */
		public static void WriteLittleEndian(int i, OutputStream o)
		{
			byte[] bi = new byte[4];
			LittleEndian.PutInt(bi,0,i);
			o.Write(bi);
		}

		/**
		 * When writing out, write out a signed short (16bit) in Little Endian format
		 */
		public static void WriteLittleEndian(short s, OutputStream o)
		{
			byte[]
			bs = new byte[2];
			LittleEndian.PutShort(bs,0,s);
			o.Write(bs);
		}

		/**
		 * Build and return the Record at the given offset.
		 * Note - does less error checking and handling than findChildRecords
		 * @param b The byte array to build from
		 * @param offset The offset to build at
		 */
		public static Record BuildRecordAtOffset(byte[] b, int offset)
		{
			long type = LittleEndian.GetUShort(b, offset + 2);
			long rlen = LittleEndian.GetUInt(b, offset + 4);

			// Sanity check the length
			int rleni = (int)rlen;
			if (rleni < 0) { rleni = 0; }

			return CreateRecordForType(type, b, offset, 8 + rleni);
		}

		/**
		 * Default method for finding child records of a container record
		 */
		public static Record[] FindChildRecords(byte[] b, int start, int len)
		{
			List<Record> children = new List<Record>(5);

			// Jump our little way along, creating records as we go
			int pos = start;
			while (pos <= (start + len - 8))
			{
				long type = LittleEndian.GetUShort(b, pos + 2);
				long rlen = LittleEndian.GetUInt(b, pos + 4);

				// Sanity check the length
				int rleni = (int)rlen;
				if (rleni < 0) { rleni = 0; }

				// Abort if first record is of type 0000 and length FFFF,
				//  as that's a sign of a screwed up record
				if (pos == 0 && type == 0L && rleni == 0xffff)
				{
					throw new CorruptPowerPointFileException("Corrupt document - starts with record of type 0000 and length 0xFFFF");
				}

				Record r = CreateRecordForType(type, b, pos, 8 + rleni);
				if (r != null)
				{
					children.Add(r);
				}
				pos += 8;
				pos += rleni;
			}

			// Turn the vector into an array, and return
			return children.ToArray();
		}

		/**
		 * For a given type (little endian bytes 3 and 4 in record header),
		 *  byte array, start position and length:
		 *  will return a Record object that will handle that record
		 *
		 * Remember that while PPT stores the record lengths as 8 bytes short
		 *  (not including the size of the header), this code assumes you're
		 *  passing in corrected lengths
		 */
		public static Record CreateRecordForType(long type, byte[] b, int start, int len)
		{
			// We use the RecordTypes class to provide us with the right
			//  class to use for a given type
			// A spot of reflection gets us the (byte[],int,int) constructor
			// From there, we instanciate the class
			// Any special record handling occurs once we have the class
			RecordTypes recordType = RecordTypes.ForTypeID((short)type);
			RecordConstructor<Record> c = recordType.RecordConstructor;
			if (c == null)
			{
				// How odd. RecordTypes normally substitutes in
				//  a default handler class if it has heard of the record
				//  type but there's no support for it. Explicitly request
				//  that now
				//LOG.atDebug().log(()-> new StringFormattedMessage("Known but unhandled record type %d (0x%04x) at offset %d", type, type, start));
				c = RecordTypes.UnknownRecordPlaceholder.RecordConstructor;
			}
			else if (recordType == RecordTypes.UnknownRecordPlaceholder)
			{
				//LOG.atDebug().log(()-> new StringFormattedMessage("Unknown placeholder type %d (0x%04x) at offset %d", type, type, start));
			}

			Record toReturn;
			try
			{
				toReturn = c(b, start, len);
			}
			catch (RuntimeException e)
			{
				// Handle case of a corrupt last record, whose claimed length
				//  would take us passed the end of the file
				if (start + len > b.Length)
				{
					//LOG.atWarn().log("Warning: Skipping record of type {} at position {} which claims to be longer than the file! ({} vs {})", type, box(start), box(len), box(b.length - start));
					return null;
				}

				throw new HSLFException("Couldn't instantiate the class for type with id " + type + " on class " + c + " : {0}", e);
			}

			// Handling for special kinds of records follow

			// If it's a position aware record, tell it where it is
			if (toReturn is PositionDependentRecord) {
			PositionDependentRecord pdr = (PositionDependentRecord)toReturn;
			pdr.SetLastOnDiskOffset(start);
		}

		// Return the created record
		return toReturn;
		}

		public abstract IDictionary<string, Func<object>> GetGenericProperties();
	}
}