/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */
namespace NPOI.HSSF.EventUserModel
{
    using System.Collections;
    using NPOI.HSSF.Record;
    using NPOI.Util;


    /// <summary>
    /// A stream based way to Get at complete records, with
    /// as low a memory footprint as possible.
    /// This handles Reading from a RecordInputStream, turning
    /// the data into full records, Processing continue records
    /// etc.
    /// Most users should use HSSFEventFactory 
    /// HSSFListener and have new records pushed to
    /// them, but this does allow for a "pull" style of coding. 
    /// </summary>
    public class HSSFRecordStream
    {
        private RecordInputStream in1;

        /** Have we run out of records on the stream? */
        private bool hitEOS = false;
        /** Have we returned all the records there are? */
        private bool complete = false;

        /**
         * Sometimes we end up with a bunch of
         *  records. When we do, these should
         *  be returned before the next normal
         *  record Processing occurs (i.e. before
         *  we Check for continue records and
         *  return rec)
         */
        private ArrayList bonusRecords = null;

        /** 
         * The next record to return, which may need to have its
         *  continue records passed to it before we do
         */
        private Record rec = null;
        /**
         * The most recent record that we gave to the user
         */
        private Record lastRec = null;
        /**
         * The most recent DrawingRecord seen
         */
        private DrawingRecord lastDrawingRecord = new DrawingRecord();

        public HSSFRecordStream(RecordInputStream inp)
        {
            this.in1 = inp;
        }

        /// <summary>
        /// Returns the next (complete) record from the
        /// stream, or null if there are no more.
        /// </summary>
        /// <returns></returns>
        public Record NextRecord()
        {
            Record r = null;

            // Loop Until we Get something
            while (r == null && !complete)
            {
                // Are there any bonus records that we need to
                //  return?
                r = GetBonusRecord();

                // If not, ask for the next real record
                if (r == null)
                {
                    r = GetNextRecord();
                }
            }

            // All done
            return r;
        }

        /// <summary>
        /// If there are any "bonus" records, that should
        /// be returned before Processing new ones,
        /// grabs the next and returns it.
        /// If not, returns null;
        /// </summary>
        /// <returns></returns>
        private Record GetBonusRecord()
        {
            if (bonusRecords != null)
            {
                Record r = (Record)bonusRecords[0];
                bonusRecords.RemoveAt(0);

                if (bonusRecords.Count == 0)
                {
                    bonusRecords = null;
                }
                return r;
            }
            return null;
        }

        /// <summary>
        /// Returns the next available record, or null if
        /// this pass didn't return a record that's
        /// suitable for returning (eg was a continue record).
        /// </summary>
        /// <returns></returns>
        private Record GetNextRecord()
        {
            Record toReturn = null;

            if (in1.HasNextRecord)
            {
                // Grab our next record
                in1.NextRecord();
                short sid = in1.Sid;

                //
                // for some reasons we have to make the workbook to be at least 4096 bytes
                // but if we have such workbook we Fill the end of it with zeros (many zeros)
                //
                // it Is not good:
                // if the Length( all zero records ) % 4 = 1
                // e.g.: any zero record would be Readed as  4 bytes at once ( 2 - id and 2 - size ).
                // And the last 1 byte will be Readed WRONG ( the id must be 2 bytes )
                //
                // So we should better to Check if the sid Is zero and not to Read more data
                // The zero sid shows us that rest of the stream data Is a fake to make workbook 
                // certain size
                //
                if (sid == 0)
                    return null;


                // If we had a last record, and this one
                //  Isn't a continue record, then pass
                //  it on to the listener
                if ((rec != null) && (sid != ContinueRecord.sid))
                {
                    // This last record ought to be returned
                    toReturn = rec;
                }

                // If this record Isn't a continue record,
                //  then build it up
                if (sid != ContinueRecord.sid)
                {
                    //Console.WriteLine("creating "+sid);
                    Record[] recs = RecordFactory.CreateRecord(in1);

                    // We know that the multiple record situations
                    //  don't contain continue records, so just
                    //  pass those on to the listener now
                    if (recs.Length > 1)
                    {
                        bonusRecords = new ArrayList(recs.Length - 1);
                        for (int k = 0; k < (recs.Length - 1); k++)
                        {
                            bonusRecords.Add(recs[k]);
                        }
                    }

                    // Regardless of the number we Created, always hold
                    //  onto the last record to be Processed on the next
                    //  loop, in case it has any continue records
                    rec = recs[recs.Length - 1];
                    // Don't return it just yet though, as we probably have
                    //  a record from the last round to return
                }
                else
                {
                    // Normally, ContinueRecords are handled internally
                    // However, in a few cases, there Is a gap between a record at
                    //  its Continue, so we have to handle them specially
                    // This logic Is much like in RecordFactory.CreateRecords()
                    Record[] recs = RecordFactory.CreateRecord(in1);
                    ContinueRecord crec = (ContinueRecord)recs[0];
                    if ((lastRec is ObjRecord) || (lastRec is TextObjectRecord))
                    {
                        // You can have Obj records between a DrawingRecord
                        //  and its continue!
                        lastDrawingRecord.ProcessContinueRecord(crec.Data);
                        // Trigger them on the drawing record, now it's complete
                        rec = lastDrawingRecord;
                    }
                    else if ((lastRec is DrawingGroupRecord))
                    {
                        ((DrawingGroupRecord)lastRec).ProcessContinueRecord(crec.Data);
                        // Trigger them on the drawing record, now it's complete
                        rec = lastRec;
                    }
                    else
                    {
                        if (rec is UnknownRecord)
                        {
                            ;//silently skip records we don't know about
                        }
                        else
                        {
                            throw new RecordFormatException("Records should handle ContinueRecord internally. Should not see this exception");
                        }
                    }
                }

                // Update our tracking of the last record
                lastRec = rec;
                if (rec is DrawingRecord)
                {
                    lastDrawingRecord = (DrawingRecord)rec;
                }
            }
            else
            {
                // No more records
                hitEOS = true;
            }

            // If we've hit the end-of-stream, then
            //  finish off the last record and be done
            if (hitEOS)
            {
                complete = true;

                // Return the last record if there was
                //  one, otherwise null
                if (rec != null)
                {
                    toReturn = rec;
                    rec = null;
                }
            }

            return toReturn;
        }
    }
}