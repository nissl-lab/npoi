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
namespace NPOI.HSSF.Record
{
    using System;
    using System.Collections.Generic;
    using NPOI;
    using System.IO;
    using NPOI.HSSF.Record.Crypto;
    using NPOI.Util;
    using NPOI.HSSF.Record.Chart;
    /**
     * A stream based way to get at complete records, with
     * as low a memory footprint as possible.
     * This handles Reading from a RecordInputStream, turning
     * the data into full records, processing continue records
     * etc.
     * Most users should use {@link HSSFEventFactory} /
     * {@link HSSFListener} and have new records pushed to
     * them, but this does allow for a "pull" style of coding.
     */
    public class RecordFactoryInputStream
    {

        /**
         * Keeps track of the sizes of the Initial records up to and including {@link FilePassRecord}
         * Needed for protected files because each byte is encrypted with respect to its absolute
         * position from the start of the stream.
         */
        private class StreamEncryptionInfo
        {
            private int _InitialRecordsSize;
            private FilePassRecord _filePassRec;
            private Record _lastRecord;
            private bool _hasBOFRecord;

            public StreamEncryptionInfo(RecordInputStream rs, List<Record> outputRecs)
            {
                Record rec;
                rs.NextRecord();
                int recSize = 4 + rs.Remaining;
                rec = RecordFactory.CreateSingleRecord(rs);
                outputRecs.Add(rec);
                FilePassRecord fpr = null;
                if (rec is BOFRecord)
                {
                    _hasBOFRecord = true;
                    // Fetch the next record, and see if it indicates whether
                    //  the document is encrypted or not
                    if (rs.HasNextRecord)
                    {
                        rs.NextRecord();
                        rec = RecordFactory.CreateSingleRecord(rs);
                        recSize += rec.RecordSize;
                        outputRecs.Add(rec);
                        // Encrypted is normally BOF then FILEPASS
					    // May sometimes be BOF, WRITEPROTECT, FILEPASS
                        if (rec is WriteProtectRecord && rs.HasNextRecord)
                        {
                            rs.NextRecord();
                            rec = RecordFactory.CreateSingleRecord(rs);
                            recSize += rec.RecordSize;
                            outputRecs.Add(rec);
                        }
                        // If it's a FILEPASS, track it specifically but
                        //  don't include it in the main stream
                        if (rec is FilePassRecord)
                        {
                            fpr = (FilePassRecord)rec;
                            outputRecs.RemoveAt(outputRecs.Count - 1);
                            // TODO - add fpr not Added to outPutRecs
                            rec = outputRecs[0];
                        }
                        else
                        {
                            // workbook not encrypted (typical case)
                            if (rec is EOFRecord)
                            {
                                // A workbook stream is never empty, so crash instead
                                // of trying to keep track of nesting level
                                throw new InvalidOperationException("Nothing between BOF and EOF");
                            }
                        }
                    }
                }
                else
                {
                    // Invalid in a normal workbook stream.
                    // However, some test cases work on sub-sections of
                    // the workbook stream that do not begin with BOF
                    _hasBOFRecord = false;
                }
                _InitialRecordsSize = recSize;
                _filePassRec = fpr;
                _lastRecord = rec;
            }

            public RecordInputStream CreateDecryptingStream(Stream original)
            {
                FilePassRecord fpr = _filePassRec;
                String userPassword = Biff8EncryptionKey.CurrentUserPassword;

                Biff8EncryptionKey key;
                if (userPassword == null)
                {
                    key = Biff8EncryptionKey.Create(fpr.DocId);
                }
                else
                {
                    key = Biff8EncryptionKey.Create(userPassword, fpr.DocId);
                }
                if (!key.Validate(fpr.SaltData, fpr.SaltHash))
                {
                    throw new EncryptedDocumentException(
                            (userPassword == null ? "Default" : "Supplied")
                            + " password is invalid for docId/saltData/saltHash");
                }
                return new RecordInputStream(original, key, _InitialRecordsSize);
            }

            public bool HasEncryption
            {
                get
                {
                    return _filePassRec != null;
                }
            }

            /**
             * @return last record scanned while looking for encryption info.
             * This will typically be the first or second record Read. Possibly <code>null</code>
             * if stream was empty
             */
            public Record LastRecord
            {
                get
                {
                    return _lastRecord;
                }
            }

            /**
             * <c>false</c> in some test cases
             */
            public bool HasBOFRecord
            {
                get
                {
                    return _hasBOFRecord;
                }
            }
        }


        private RecordInputStream _recStream;
        private bool _shouldIncludeContinueRecords;

        /**
         * Temporarily stores a group of {@link Record}s, for future return by {@link #nextRecord()}.
         * This is used at the start of the workbook stream, and also when the most recently read
         * underlying record is a {@link MulRKRecord}
         */
        private Record[] _unreadRecordBuffer;

        /**
         * used to help iterating over the unread records
         */
        private int _unreadRecordIndex = -1;

        /**
         * The most recent record that we gave to the user
         */
        private Record _lastRecord = null;
        /**
         * The most recent DrawingRecord seen
         */
        private DrawingRecord _lastDrawingRecord = new DrawingRecord();

        private int _bofDepth;

        private bool _lastRecordWasEOFLevelZero;


        /**
         * @param shouldIncludeContinueRecords caller can pass <c>false</c> if loose
         * {@link ContinueRecord}s should be skipped (this is sometimes useful in event based
         * processing).
         */
        public RecordFactoryInputStream(Stream in1, bool shouldIncludeContinueRecords)
        {
            RecordInputStream rs = new RecordInputStream(in1);
            List<Record> records = new List<Record>();
            StreamEncryptionInfo sei = new StreamEncryptionInfo(rs, records);
            if (sei.HasEncryption)
            {
                rs = sei.CreateDecryptingStream(in1);
            }
            else
            {
                // typical case - non-encrypted stream
            }

            if (records.Count != 0)
            {
                _unreadRecordBuffer = new Record[records.Count];
                _unreadRecordBuffer = records.ToArray();
                _unreadRecordIndex = 0;
            }
            _recStream = rs;
            _shouldIncludeContinueRecords = shouldIncludeContinueRecords;
            _lastRecord = sei.LastRecord;

            /*
            * How to recognise end of stream?
            * In the best case, the underlying input stream (in) ends just after the last EOF record
            * Usually however, the stream is pAdded with an arbitrary byte count.  Excel and most apps
            * reliably use zeros for pAdding and if this were always the case, this code could just
            * skip all the (zero sized) records with sid==0.  However, bug 46987 Shows a file with
            * non-zero pAdding that is read OK by Excel (Excel also fixes the pAdding).
            *
            * So to properly detect the workbook end of stream, this code has to identify the last
            * EOF record.  This is not so easy because the worbook bof+eof pair do not bracket the
            * whole stream.  The worksheets follow the workbook, but it is not easy to tell how many
            * sheet sub-streams should be present.  Hence we are looking for an EOF record that is not
            * immediately followed by a BOF record.  One extra complication is that bof+eof sub-
            * streams can be nested within worksheet streams and it's not clear in these cases what
            * record might follow any EOF record.  So we also need to keep track of the bof/eof
            * nesting level.
            */
            _bofDepth = sei.HasBOFRecord ? 1 : 0;
            _lastRecordWasEOFLevelZero = false;
        }

        /**
         * Returns the next (complete) record from the
         * stream, or null if there are no more.
         */
        public Record NextRecord()
        {
            Record r;
            r = GetNextUnreadRecord();
            if (r != null)
            {
                // found an unread record
                return r;
            }
            while (true)
            {
                if (!_recStream.HasNextRecord)
                {
                    // recStream is exhausted;
                    return null;
                }

                // step underlying RecordInputStream to the next record
                _recStream.NextRecord();

                if (_lastRecordWasEOFLevelZero)
                {
                    // Potential place for ending the workbook stream
                    // Check that the next record is not BOFRecord(0x0809)
                    // Normally the input stream Contains only zero pAdding after the last EOFRecord,
                    // but bug 46987 and 48068 suggests that the padding may be garbage.
                    // This code relies on the pAdding bytes not starting with BOFRecord.sid
                    if (_recStream.Sid != BOFRecord.sid)
                    {
                        return null;
                    }
                    // else - another sheet substream starting here
                }

                r = ReadNextRecord();
                if (r == null)
                {
                    // some record types may get skipped (e.g. DBCellRecord and ContinueRecord)
                    continue;
                }
                return r;
            }
        }

        /**
         * @return the next {@link Record} from the multiple record group as expanded from
         * a recently read {@link MulRKRecord}. <code>null</code> if not present.
         */
        private Record GetNextUnreadRecord()
        {
            if (_unreadRecordBuffer != null)
            {
                int ix = _unreadRecordIndex;
                if (ix < _unreadRecordBuffer.Length)
                {
                    Record result = _unreadRecordBuffer[ix];
                    _unreadRecordIndex = ix + 1;
                    return result;
                }
                _unreadRecordIndex = -1;
                _unreadRecordBuffer = null;
            }
            return null;
        }

        /**
         * @return the next available record, or <code>null</code> if
         * this pass didn't return a record that's
         * suitable for returning (eg was a continue record).
         */
        private Record ReadNextRecord()
        {

            Record record = RecordFactory.CreateSingleRecord(_recStream);
            _lastRecordWasEOFLevelZero = false;

            if (record is BOFRecord)
            {
                _bofDepth++;
                return record;
            }

            if (record is EOFRecord)
            {
                _bofDepth--;
                if (_bofDepth < 1)
                {
                    _lastRecordWasEOFLevelZero = true;
                }

                return record;
            }

            if (record is DBCellRecord)
            {
                // Not needed by POI.  Regenerated from scratch by POI when spreadsheet is written
                return null;
            }

            if (record is RKRecord)
            {
                return RecordFactory.ConvertToNumberRecord((RKRecord)record);
            }

            if (record is MulRKRecord)
            {
                Record[] records = RecordFactory.ConvertRKRecords((MulRKRecord)record);

                _unreadRecordBuffer = records;
                _unreadRecordIndex = 1;
                return records[0];
            }

            if (record.Sid == DrawingGroupRecord.sid
                    && _lastRecord is DrawingGroupRecord)
            {
                DrawingGroupRecord lastDGRecord = (DrawingGroupRecord)_lastRecord;
                lastDGRecord.Join((AbstractEscherHolderRecord)record);
                return null;
            }
            if (record.Sid == ContinueRecord.sid)
            {
                ContinueRecord contRec = (ContinueRecord)record;

                if (_lastRecord is ObjRecord || _lastRecord is TextObjectRecord)
                {
                    // Drawing records have a very strange continue behaviour.
                    //There can actually be OBJ records mixed between the continues.
                    _lastDrawingRecord.ProcessContinueRecord(contRec.Data);
                    //we must remember the position of the continue record.
                    //in the serialization procedure the original structure of records must be preserved
                    if (_shouldIncludeContinueRecords)
                    {
                        return record;
                    }
                    return null;
                }
                if (_lastRecord is DrawingGroupRecord)
                {
                    ((DrawingGroupRecord)_lastRecord).ProcessContinueRecord(contRec.Data);
                    return null;
                }
                if (_lastRecord is DrawingRecord)
                {
                    //((DrawingRecord)_lastRecord).ProcessContinueRecord(contRec.Data);
                    return contRec;
                }
                if (_lastRecord is CrtMlFrtRecord)
                {
                    return record;
                }
                if (_lastRecord is UnknownRecord)
                {
                    //Gracefully handle records that we don't know about,
                    //that happen to be continued
                    return record;
                }
                if (_lastRecord is EOFRecord)
                {
                    // This is really odd, but excel still sometimes
                    //  outPuts a file like this all the same
                    return record;
                }
                //if (_lastRecord is StringRecord)
                //{
                //    ((StringRecord)_lastRecord).ProcessContinueRecord(contRec.Data);
                //    return null;
                //}
                throw new RecordFormatException("Unhandled Continue Record");
            }
            _lastRecord = record;
            if (record is DrawingRecord)
            {
                _lastDrawingRecord = (DrawingRecord)record;
            }
            return record;
        }
    }
}
