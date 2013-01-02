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
    using NPOI.HSLF.Exceptions;
    using System.Reflection;
    using System.Collections.Generic;

    /**
     * This abstract class represents a record in the PowerPoint document.
     * Record classes should extend with RecordContainer or RecordAtom, which
     *  extend this in turn.
     *
     * @author Nick Burch
     */

    public abstract class Record
    {
        // For logging
        protected POILogger logger = POILogFactory.GetLogger(typeof(Record));

        /**
         * Is this record type an Atom record (only has data),
         *  or is it a non-Atom record (has other records)?
         */
        public abstract bool IsAnAtom { get; }

        /**
         * Returns the type (held as a little endian in bytes 3 and 4)
         *  that this class handles
         */
        public abstract long RecordType { get; }

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
        public abstract void WriteOut(Stream o);

        /**
         * When writing out, write out a signed int (32bit) in Little Endian format
         */
        public static void WriteLittleEndian(int i, Stream o)
        {
            byte[] bi = new byte[4];
            LittleEndian.PutInt(bi, i);
            o.Write(bi, (int)o.Position, bi.Length);
        }
        /**
         * When writing out, write out a signed short (16bit) in Little Endian format
         */
        public static void WriteLittleEndian(short s, Stream o)
        {
            byte[] bs = new byte[2];
            LittleEndian.PutShort(bs, s);
            o.Write(bs,(int)o.Position,bs.Length);
        }

        /**
         * Build and return the Record at the given OffSet.
         * Note - does less error Checking and handling than FindChildRecords
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
         * Default method for Finding child records of a Container record
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
                else
                {
                    // Record was horribly corrupt
                }
                pos += 8;
                pos += rleni;
            }

            // Turn the vector into an array, and return
            Record[] cRecords = new Record[children.Count];
            for (int i = 0; i < children.Count; i++)
            {
                cRecords[i] = (Record)children[i];
            }
            return cRecords;
        }

        /**
         * For a given type (little endian bytes 3 and 4 in record header),
         *  byte array, start position and Length:
         *  will return a Record object that will handle that record
         *
         * Remember that while PPT stores the record Lengths as 8 bytes short
         *  (not including the size of the header), this code assumes you're
         *  passing in corrected Lengths
         */
        public static Record CreateRecordForType(long type, byte[] b, int start, int len)
        {
            Record toReturn = null;

            // Handle case of a corrupt last record, whose claimed length
            //  would take us passed the end of the file
            if (start + len > b.Length)
            {
                Console.Error.WriteLine("Warning: Skipping record of type " + type + " at position " + start + " which claims to be longer than the file! (" + len + " vs " + (b.Length - start) + ")");
                return null;
            }

            // We use the RecordTypes class to provide us with the right
            //  class to use for a given type
            // A spot of reflection Gets us the (byte[],int,int) constructor
            // From there, we instanciate the class
            // Any special record handling occurs once we have the class
            Type c = null;
            try
            {
                c = RecordTypes.RecordHandlingClass((int)type);
                if (c == null)
                {
                    // How odd. RecordTypes normally subsitutes in
                    //  a default handler class if it has heard of the record
                    //  type but there's no support for it. Explicitly request
                    //  that now
                    c = RecordTypes.RecordHandlingClass(RecordTypes.Unknown.typeID);
                }

                // Grab the right constructor
                ConstructorInfo con = c.GetConstructor(new Type[] { typeof(byte[]), typeof(Int32), typeof(Int32) });
                // Instantiate
                toReturn = (Record)(con.Invoke(new Object[] { b, start, len }));
            }
            catch (TargetInvocationException ite)
            {
                throw new Exception("Couldn't instantiate the class for type with id " + type + " on class " + c + " : " + ite + "\nCause was : " + ite.Message, ite);
            }
            catch (MethodAccessException iae)
            {
                throw new Exception("Couldn't access the constructor for type with id " + type + " on class " + c + " : " + iae, iae);
            }
            catch (NotSupportedException nsme)
            {
                throw new Exception("Couldn't access the constructor for type with id " + type + " on class " + c + " : " + nsme, nsme);
            }

            // Handling for special kinds of records follow

            // If it's a position aware record, tell it where it is
            if (toReturn is PositionDependentRecord)
            {
                PositionDependentRecord pdr = (PositionDependentRecord)toReturn;
                pdr.LastOnDiskOffset = (start);
            }

            // Return the Created record
            return toReturn;
        }
    }
}

