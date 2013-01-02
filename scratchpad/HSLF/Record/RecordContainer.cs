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
    using NPOI.HSLF.Util;
    using System.Collections.Generic;

    /**
     * Abstract class which all Container records will extend. Providers
     *  helpful methods for writing child records out to disk
     *
     * @author Nick Burch
     */

    public abstract class RecordContainer : Record
    {
        protected Record[] _children;
        //private Boolean changingChildRecordsLock = true;

        /**
         * Return any children
         */
        public override Record[] GetChildRecords() { return _children; }

        /**
         * We're not an atom
         */
        public override bool IsAnAtom
        {
            get { return false; }
        }


        /* ===============================================================
         *                   Internal Move Helpers
         * ===============================================================
         */

        /**
         * Finds the location of the given child record
         */
        private int FindChildLocation(Record child)
        {
            // Synchronized as we don't want things changing
            //  as we're doing our search
            lock (this)
            {
                for (int i = 0; i < _children.Length; i++)
                {
                    if (_children[i].Equals(child))
                    {
                        return i;
                    }
                }
            }
            return -1;
        }

        /**
         * Adds a child record, at the very end.
         * @param newChild The child record to add
         */
        private void AppendChild(Record newChild)
        {
            lock (this)
            {
                // Copy over, and pop the child in at the end
                Record[] nc = new Record[(_children.Length + 1)];
                Array.Copy(_children, 0, nc, 0, _children.Length);
                // Switch the arrays
                nc[_children.Length] = newChild;
                _children = nc;
            }
        }

        /**
         * Adds the given new Child Record at the given location,
         *  shuffling everything from there on down by one
         * @param newChild
         * @param position
         */
        private void AddChildAt(Record newChild, int position)
        {
            lock (this)
            {
                // Firstly, have the child Added in at the end
                AppendChild(newChild);

                // Now, have them moved to the right place
                MoveChildRecords((_children.Length - 1), position, 1);
            }
        }

        /**
         * Moves <i>number</i> child records from <i>oldLoc</i>
         *  to <i>newLoc</i>. Caller must have the changingChildRecordsLock
         * @param oldLoc the current location of the records to move
         * @param newLoc the new location for the records
         * @param number the number of records to move
         */
        private void MoveChildRecords(int oldLoc, int newLoc, int number)
        {
            if (oldLoc == newLoc) { return; }
            if (number == 0) { return; }

            // Check that we're not asked to move too many
            if (oldLoc + number > _children.Length)
            {
                throw new ArgumentException("Asked to move more records than there are!");
            }

            // Do the move
            Arrays.ArrayMoveWithin(_children, oldLoc, newLoc, number);
        }


        /**
         * Finds the first child record of the given type,
         *  or null if none of the child records are of the
         *  given type. Does not descend.
         */
        public Record FindFirstOfType(long type)
        {
            for (int i = 0; i < _children.Length; i++)
            {
                if (_children[i].RecordType == type)
                {
                    return _children[i];
                }
            }
            return null;
        }

        /**
         * Remove a child record from this record Container
         *
         * @param ch the child to remove
         * @return the Removed record
         */
        public Record RemoveChild(Record ch)
        {
            Record rm = null;
            List<Record> lst = new List<Record>();
            foreach (Record r in _children)
            {
                if (r != ch) lst.Add(r);
                else rm = r;
            }
            _children = lst.ToArray();
            return rm;
        }

        /* ===============================================================
         *                   External Move Methods
         * ===============================================================
         */

        /**
         * Add a new child record onto a record's list of children.
         */
        public void AppendChildRecord(Record newChild)
        {
            lock (this)
            {
                AppendChild(newChild);
            }
        }

        /**
         * Adds the given Child Record after the supplied record
         * @param newChild
         * @param after
         */
        public void AddChildAfter(Record newChild, Record after)
        {
            lock (this)
            {
                // Decide where we're going to put it
                int loc = FindChildLocation(after);
                if (loc == -1)
                {
                    throw new ArgumentException("Asked to add a new child after another record, but that record wasn't one of our children!");
                }

                // Add one place after the supplied record
                AddChildAt(newChild, loc + 1);
            }
        }

        /**
         * Adds the given Child Record before the supplied record
         * @param newChild
         * @param before
         */
        public void AddChildBefore(Record newChild, Record before)
        {
            lock (this)
            {
                // Decide where we're going to put it
                int loc = FindChildLocation(before);
                if (loc == -1)
                {
                    throw new ArgumentException("Asked to add a new child before another record, but that record wasn't one of our children!");
                }

                // Add at the place of the supplied record
                AddChildAt(newChild, loc);
            }
        }

        /**
         * Moves the given Child Record to before the supplied record
         */
        public void MoveChildBefore(Record child, Record before)
        {
            MoveChildrenBefore(child, 1, before);
        }

        /**
         * Moves the given Child Records to before the supplied record
         */
        public void MoveChildrenBefore(Record firstChild, int number, Record before)
        {
            if (number < 1) { return; }

            lock (this)
            {
                // Decide where we're going to put them
                int newLoc = FindChildLocation(before);
                if (newLoc == -1)
                {
                    throw new ArgumentException("Asked to move children before another record, but that record wasn't one of our children!");
                }

                // Figure out where they are now
                int oldLoc = FindChildLocation(firstChild);
                if (oldLoc == -1)
                {
                    throw new ArgumentException("Asked to move a record that wasn't a child!");
                }

                // Actually move
                MoveChildRecords(oldLoc, newLoc, number);
            }
        }

        /**
         * Moves the given Child Records to after the supplied record
         */
        public void MoveChildrenAfter(Record firstChild, int number, Record after)
        {
            if (number < 1) { return; }

            lock (this)
            {
                // Decide where we're going to put them
                int newLoc = FindChildLocation(after);
                if (newLoc == -1)
                {
                    throw new ArgumentException("Asked to move children before another record, but that record wasn't one of our children!");
                }
                // We actually want after this though
                newLoc++;

                // Figure out where they are now
                int oldLoc = FindChildLocation(firstChild);
                if (oldLoc == -1)
                {
                    throw new ArgumentException("Asked to move a record that wasn't a child!");
                }

                // Actually move
                MoveChildRecords(oldLoc, newLoc, number);
            }
        }

        /**
         * Set child records.
         *
         * @param records   the new child records
         */
        public void SetChildRecord(Record[] records)
        {
            this._children = records;
        }

        /* ===============================================================
         *                 External Serialisation Methods
         * ===============================================================
         */

        /**
         * Write out our header, and our children.
         * @param headerA the first byte of the header
         * @param headerB the second byte of the header
         * @param type the record type
         * @param children our child records
         * @param out the stream to write to
         */
        public void WriteOut(byte headerA, byte headerB, long type, Record[] children, Stream out1)
        {
            // If we have a mutable output stream, take advantage of that
            if (out1 is MutableMemoryStream)
            {
                MutableMemoryStream mout =
                    (MutableMemoryStream)out1;

                // Grab current size
                int oldSize = mout.GetBytesWritten();

                // Write out our header, less the size
                mout.Write(new byte[] { headerA, headerB });
                byte[] typeB = new byte[2];
                LittleEndian.PutShort(typeB, (short)type);
                mout.Write(typeB);
                mout.Write(new byte[4]);

                // Write out the children
                for (int i = 0; i < children.Length; i++)
                {
                    children[i].WriteOut(mout);
                }

                // Update our header with the size
                // Don't forget to knock 8 more off, since we don't include the
                //  header in the size
                int length = mout.GetBytesWritten() - oldSize - 8;
                byte[] size = new byte[4];
                LittleEndian.PutInt(size, 0, length);
                mout.OverWrite(size, oldSize + 4);
            }
            else
            {
                // Going to have to do it a slower way, because we have
                // to update the length come the end

                // Create a MemoryStream to hold everything in
                MemoryStream baos = new MemoryStream();

                // Write out our header, less the size
                baos.Write(new byte[] { headerA, headerB }, 0, 2);
                byte[] typeB = new byte[2];
                LittleEndian.PutShort(typeB, (short)type);
                baos.Write(typeB, 2, 2);
                baos.Write(new byte[] { 0, 0, 0, 0 }, 4, 4);

                // Write out our children
                for (int i = 0; i < children.Length; i++)
                {
                    children[i].WriteOut(baos);
                }

                // Grab the bytes back
                byte[] toWrite = baos.ToArray();

                // Update our header with the size
                // Don't forget to knock 8 more off, since we don't include the
                //  header in the size
                LittleEndian.PutInt(toWrite, 4, (toWrite.Length - 8));

                // Write out the bytes
                out1.Write(toWrite, (int)out1.Position, toWrite.Length);
            }
        }
    }

}