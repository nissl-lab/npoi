
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

namespace NPOI.DDF
{
    using System;
    using System.Text;
    using System.Collections;
    using NPOI.Util;
    using System.Collections.Generic;


    /// <summary>
    /// Escher container records store other escher records as children.
    /// The container records themselves never store any information beyond
    /// the standard header used by all escher records.  This one record is
    /// used to represent many different types of records.
    /// @author Glen Stampoultzis
    /// </summary>
    public class EscherContainerRecord : EscherRecord
    {
        public const short DGG_CONTAINER = unchecked((short)0xF000);
        public const short BSTORE_CONTAINER = unchecked((short)0xF001);
        public const short DG_CONTAINER = unchecked((short)0xF002);
        public const short SPGR_CONTAINER = unchecked((short)0xF003);
        public const short SP_CONTAINER = unchecked((short)0xF004);
        public const short SOLVER_CONTAINER = unchecked((short)0xF005);
        private static POILogger log = POILogFactory.GetLogger(typeof(EscherContainerRecord));

        /**
         * in case if document contains any charts we have such document structure:
         * BOF
         * ...
         * DrawingRecord
         * ...
         * ObjRecord|TxtObjRecord
         * ...
         * EOF
         * ...
         * BOF(Chart begin)
         * ...
         * DrawingRecord
         * ...
         * ObjRecord|TxtObjRecord
         * ...
         * EOF
         * So, when we call EscherAggregate.createAggregate() we have not all needed data.
         * When we got warning "WARNING: " + bytesRemaining + " bytes remaining but no space left"
         * we should save value of bytesRemaining
         * and add it to container size when we serialize it
         */
        private int _remainingLength;
        private List<EscherRecord> _childRecords = new List<EscherRecord>();

        /// <summary>
        /// The contract of this method is to deSerialize an escher record including
        /// it's children.
        /// </summary>
        /// <param name="data">The byte array containing the Serialized escher
        /// records.</param>
        /// <param name="offset">The offset into the byte array.</param>
        /// <param name="recordFactory">A factory for creating new escher records</param>
        /// <returns>The number of bytes written.</returns>        
        public override int FillFields(byte[] data, int offset, IEscherRecordFactory recordFactory)
        {
            int bytesRemaining = ReadHeader(data, offset);
            int bytesWritten = 8;
            offset += 8;
            while (bytesRemaining > 0 && offset < data.Length)
            {
                EscherRecord child = recordFactory.CreateRecord(data, offset);
                int childBytesWritten = child.FillFields(data, offset, recordFactory);
                bytesWritten += childBytesWritten;
                offset += childBytesWritten;
                bytesRemaining -= childBytesWritten;
                AddChildRecord(child);
                if (offset >= data.Length && bytesRemaining > 0)
                {
                    _remainingLength = bytesRemaining;
                    log.Log(POILogger.WARN, "Not enough Escher data: " + bytesRemaining + " bytes remaining but no space left");
                }
            }
            return bytesWritten;
        }

        /// <summary>
        /// Serializes to an existing byte array without serialization listener.
        /// This is done by delegating to Serialize(int, byte[], EscherSerializationListener).
        /// </summary>
        /// <param name="offset">the offset within the data byte array.</param>
        /// <param name="data"> the data array to Serialize to.</param>
        /// <param name="listener">a listener for begin and end serialization events.</param>
        /// <returns>The number of bytes written.</returns>
        public override int Serialize(int offset, byte[] data, EscherSerializationListener listener)
        {
            listener.BeforeRecordSerialize(offset, RecordId, this);

            LittleEndian.PutShort(data, offset, Options);
            LittleEndian.PutShort(data, offset + 2, RecordId);
            int remainingBytes = 0;
            for (IEnumerator iterator = ChildRecords.GetEnumerator(); iterator.MoveNext(); )
            {
                EscherRecord r = (EscherRecord)iterator.Current;
                remainingBytes += r.RecordSize;
            }

            remainingBytes += _remainingLength;

            LittleEndian.PutInt(data, offset + 4, remainingBytes);
            int pos = offset + 8;
            for (IEnumerator iterator = ChildRecords.GetEnumerator(); iterator.MoveNext(); )
            {
                EscherRecord r = (EscherRecord)iterator.Current;
                pos += r.Serialize(pos, data, listener);
            }

            listener.AfterRecordSerialize(pos, RecordId, pos - offset, this);
            return pos - offset;
        }

        /// <summary>
        /// Subclasses should effeciently return the number of bytes required to
        /// Serialize the record.
        /// </summary>
        /// <value>number of bytes</value>
        public override int RecordSize
        {
            get
            {
                int childRecordsSize = 0;
                for (IEnumerator iterator = ChildRecords.GetEnumerator(); iterator.MoveNext(); )
                {
                    EscherRecord r = (EscherRecord)iterator.Current;
                    childRecordsSize += r.RecordSize;
                }
                return 8 + childRecordsSize;
            }
        }

        /// <summary>
        /// Do any of our (top level) children have the
        /// given recordId?
        /// </summary>
        /// <param name="recordId">The record id.</param>
        /// <returns>
        /// 	<c>true</c> if [has child of type] [the specified record id]; otherwise, <c>false</c>.
        /// </returns>
        public bool HasChildOfType(short recordId)
        {
            for (IEnumerator iterator = ChildRecords.GetEnumerator(); iterator.MoveNext(); )
            {
                EscherRecord r = (EscherRecord)iterator.Current;
                if (r.RecordId == recordId)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Returns a list of all the child (escher) records
        /// of the container.
        /// </summary>
        /// <value></value>
        public override List<EscherRecord> ChildRecords
        {
            get { return new List<EscherRecord>(_childRecords); }
            set
            {
                if (value == _childRecords)
                {
                    throw new InvalidOperationException("Child records private data member has escaped");
                }
                _childRecords.Clear();
                _childRecords.AddRange(value);
            }
        }

        public bool RemoveChildRecord(EscherRecord toBeRemoved)
        {
            return _childRecords.Remove(toBeRemoved);
        }
        public List<EscherRecord>.Enumerator GetChildIterator()
        {
            return _childRecords.GetEnumerator();
        }
        /// <summary>
        /// Returns all of our children which are also
        /// EscherContainers (may be 0, 1, or vary rarely
        /// 2 or 3)
        /// </summary>
        /// <value>The child containers.</value>
        public IList<EscherContainerRecord> ChildContainers
        {
            get
            {
                IList<EscherContainerRecord> containers = new List<EscherContainerRecord>();
                for (IEnumerator iterator = ChildRecords.GetEnumerator(); iterator.MoveNext(); )
                {
                    EscherRecord r = (EscherRecord)iterator.Current;
                    if (r is EscherContainerRecord)
                    {
                        containers.Add((EscherContainerRecord)r);
                    }
                }
                return containers;
            }
        }


        /// <summary>
        /// Subclasses should return the short name for this escher record.
        /// </summary>
        /// <value></value>
        public override String RecordName
        {
            get
            {
                switch ((short)RecordId)
                {
                    case DGG_CONTAINER:
                        return "DggContainer";
                    case BSTORE_CONTAINER:
                        return "BStoreContainer";
                    case DG_CONTAINER:
                        return "DgContainer";
                    case SPGR_CONTAINER:
                        return "SpgrContainer";
                    case SP_CONTAINER:
                        return "SpContainer";
                    case SOLVER_CONTAINER:
                        return "SolverContainer";
                    default:
                        return "Container 0x" + HexDump.ToHex(RecordId);
                }
            }
        }

        /// <summary>
        /// The display methods allows escher variables to print the record names
        /// according to their hierarchy.
        /// </summary>
        /// <param name="indent">The current indent level.</param> 
        public override void Display(int indent)
        {
            base.Display(indent);
            for (IEnumerator iterator = _childRecords.GetEnumerator(); iterator.MoveNext(); )
            {
                EscherRecord escherRecord = (EscherRecord)iterator.Current;
                escherRecord.Display(indent + 1);
            }
        }

        /// <summary>
        /// Adds the child record.
        /// </summary>
        /// <param name="record">The record.</param>
        public void AddChildRecord(EscherRecord record)
        {
            this._childRecords.Add(record);
        }

        public void AddChildBefore(EscherRecord record, int insertBeforeRecordId)
        {
            for (int i = 0; i < _childRecords.Count; i++)
            {
                EscherRecord rec = _childRecords[(i)];
                if (rec.RecordId == insertBeforeRecordId)
                {
                    _childRecords.Insert(i++, record);
                    // TODO - keep looping? Do we expect multiple matches?
                }
            }
        }
        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override string ToString()
        {
            String nl = Environment.NewLine;

            StringBuilder children = new StringBuilder();
            if (ChildRecords.Count > 0)
            {
                children.Append("  children: " + nl);

                int count = 0;
                for (IEnumerator iterator = ChildRecords.GetEnumerator(); iterator.MoveNext(); )
                {

                    EscherRecord record = (EscherRecord)iterator.Current;
                    children.Append("    Child " + count + ":" + nl);

                    String childResult = (record).ToString();
                    childResult = childResult.Replace("\n", "\n    ");
                    children.Append("    ");
                    children.Append(childResult);
                    children.Append(nl);

                    //if (record is EscherContainerRecord)
                    //{
                    //    EscherContainerRecord ecr = (EscherContainerRecord)record;
                    //    children.Append(ecr.ToString());
                    //}
                    //else
                    //{
                    //    children.Append(record.ToString());
                    //}
                    count++;
                }
            }

            return
                this.GetType().Name + " (" + RecordName + "):" + nl +
                "  isContainer: " + IsContainerRecord + nl +
                "  version: 0x" + HexDump.ToHex(Version) + nl +
                "  instance: 0x" + HexDump.ToHex(Instance) + nl +
                "  recordId: 0x" + HexDump.ToHex(RecordId) + nl +
                "  numchildren: " + ChildRecords.Count + nl +
                children.ToString();
        }

        public override String ToXml(String tab)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(tab).Append(FormatXmlRecordHeader(RecordName, HexDump.ToHex(RecordId), HexDump.ToHex(Version), HexDump.ToHex(Instance)));
            for (IEnumerator<EscherRecord> iterator = _childRecords.GetEnumerator(); iterator.MoveNext(); )
            {
                EscherRecord record = iterator.Current;
                builder.Append(record.ToXml(tab + "\t"));
            }
            builder.Append(tab).Append("</").Append(RecordName).Append(">\n");
            return builder.ToString();
        }
        /// <summary>
        /// Gets the child by id.
        /// </summary>
        /// <param name="recordId">The record id.</param>
        /// <returns></returns>
        public EscherRecord GetChildById(short recordId)
        {
            for (IEnumerator iterator = _childRecords.GetEnumerator(); iterator.MoveNext(); )
            {
                EscherRecord escherRecord = (EscherRecord)iterator.Current;
                if (escherRecord.RecordId == recordId)
                    return escherRecord;
            }
            return null;
        }


        /// <summary>
        /// Recursively find records with the specified record ID
        /// </summary>
        /// <param name="recordId"></param>
        /// <param name="out1">list to store found records</param>
        public void GetRecordsById(short recordId, ref ArrayList out1)
        {
            for (IEnumerator it = ChildRecords.GetEnumerator(); it.MoveNext(); )
            {
                Object er = it.Current;
                EscherRecord r = (EscherRecord)er;
                if (r is EscherContainerRecord)
                {
                    EscherContainerRecord c = (EscherContainerRecord)r;
                    c.GetRecordsById(recordId, ref out1);
                }
                else if (r.RecordId == recordId)
                {
                    out1.Add(er);
                }
            }
        }
    }
}