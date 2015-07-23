
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
    /// This record defines the drawing groups used for a particular sheet.
    /// </summary>
    public class EscherDggRecord : EscherRecord
    {
        public const short RECORD_ID = unchecked((short)0xF006);
        public const String RECORD_DESCRIPTION = "MsofbtDgg";

        private int field_1_shapeIdMax;
        //    private int field_2_numIdClusters;      // for some reason the number of clusters is actually the real number + 1
        private int field_3_numShapesSaved;
        private int field_4_drawingsSaved;
        private FileIdCluster[] field_5_fileIdClusters;
        private int maxDgId;

        public class FileIdCluster
        {
            public FileIdCluster(int drawingGroupId, int numShapeIdsUsed)
            {
                this.field_1_drawingGroupId = drawingGroupId;
                this.field_2_numShapeIdsUsed = numShapeIdsUsed;
            }

            private int field_1_drawingGroupId;
            private int field_2_numShapeIdsUsed;

            public int DrawingGroupId
            {
                get { return field_1_drawingGroupId; }
            }

            public int NumShapeIdsUsed
            {
                get { return field_2_numShapeIdsUsed; }
            }

            public void IncrementShapeId()
            {
                this.field_2_numShapeIdsUsed++;
            }
        }

        /// <summary>
        /// This method deSerializes the record from a byte array.
        /// </summary>
        /// <param name="data">The byte array containing the escher record information</param>
        /// <param name="offset">The starting offset into data</param>
        /// <param name="recordFactory">May be null since this is not a container record.</param>
        /// <returns>The number of bytes Read from the byte array.</returns>
        public override int FillFields(byte[] data, int offset, IEscherRecordFactory recordFactory)
        {
            int bytesRemaining = ReadHeader(data, offset);
            int pos = offset + 8;
            int size = 0;
            field_1_shapeIdMax = LittleEndian.GetInt(data, pos + size); size += 4;
            int field_2_numIdClusters = LittleEndian.GetInt(data, pos + size); size += 4;
            field_3_numShapesSaved = LittleEndian.GetInt(data, pos + size); size += 4;
            field_4_drawingsSaved = LittleEndian.GetInt(data, pos + size); size += 4;
            field_5_fileIdClusters = new FileIdCluster[(bytesRemaining - size) / 8];  // Can't rely on field_2_numIdClusters
            for (int i = 0; i < field_5_fileIdClusters.Length; i++)
            {
                field_5_fileIdClusters[i] = new FileIdCluster(LittleEndian.GetInt(data, pos + size), LittleEndian.GetInt(data, pos + size + 4));
                maxDgId = Math.Max(maxDgId, field_5_fileIdClusters[i].DrawingGroupId);
                size += 8;
            }
            bytesRemaining -= size;
            if (bytesRemaining != 0)
                throw new RecordFormatException("Expecting no remaining data but got " + bytesRemaining + " byte(s).");
            return 8 + size + bytesRemaining;
        }

        /// <summary>
        /// This method Serializes this escher record into a byte array.
        /// </summary>
        /// <param name="offset">The offset into data to start writing the record data to.</param>
        /// <param name="data">The byte array to Serialize to.</param>
        /// <param name="listener">a listener for begin and end serialization events.</param>
        /// <returns>The number of bytes written.</returns>
        public override int Serialize(int offset, byte[] data, EscherSerializationListener listener)
        {
            listener.BeforeRecordSerialize(offset, RecordId, this);

            int pos = offset;
            LittleEndian.PutShort(data, pos, Options); pos += 2;
            LittleEndian.PutShort(data, pos, RecordId); pos += 2;
            int remainingBytes = RecordSize - 8;
            LittleEndian.PutInt(data, pos, remainingBytes); pos += 4;

            LittleEndian.PutInt(data, pos, field_1_shapeIdMax); pos += 4;
            LittleEndian.PutInt(data, pos, NumIdClusters); pos += 4;
            LittleEndian.PutInt(data, pos, field_3_numShapesSaved); pos += 4;
            LittleEndian.PutInt(data, pos, field_4_drawingsSaved); pos += 4;
            for (int i = 0; i < field_5_fileIdClusters.Length; i++)
            {
                LittleEndian.PutInt(data, pos, field_5_fileIdClusters[i].DrawingGroupId); pos += 4;
                LittleEndian.PutInt(data, pos, field_5_fileIdClusters[i].NumShapeIdsUsed); pos += 4;
            }

            listener.AfterRecordSerialize(pos, RecordId, RecordSize, this);
            return RecordSize;
        }

        /// <summary>
        /// Returns the number of bytes that are required to Serialize this record.
        /// </summary>
        /// <value>Number of bytes</value>
        public override int RecordSize
        {
            get { return 8 + 16 + (8 * field_5_fileIdClusters.Length); }
        }

        /// <summary>
        /// Return the current record id.
        /// </summary>
        /// <value>The 16 bit record id.</value>
        public override short RecordId
        {
            get { return RECORD_ID; }
        }

        /// <summary>
        /// The short name for this record
        /// </summary>
        /// <value></value>
        public override String RecordName
        {
            get { return "Dgg"; }
        }

        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override String ToString()
        {
            String nl = Environment.NewLine;

            //        String extraData;
            //        MemoryStream b = new MemoryStream();
            //        try
            //        {
            //            HexDump.dump(this.remainingData, 0, b, 0);
            //            extraData = b.ToString();
            //        }
            //        catch ( Exception e )
            //        {
            //            extraData = "error";
            //        }
            StringBuilder field_5_string = new StringBuilder();
            for (int i = 0; i < field_5_fileIdClusters.Length; i++)
            {
                field_5_string.Append("  DrawingGroupId").Append(i + 1).Append(": ");
                field_5_string.Append(field_5_fileIdClusters[i].DrawingGroupId);
                field_5_string.Append(nl);
                field_5_string.Append("  NumShapeIdsUsed").Append(i + 1).Append(": ");
                field_5_string.Append(field_5_fileIdClusters[i].NumShapeIdsUsed);
                field_5_string.Append(nl);
            }
            return GetType().Name + ":" + nl +
                    "  RecordId: 0x" + HexDump.ToHex(RECORD_ID) + nl +
                    "  Version: 0x" + HexDump.ToHex(Version) + nl +
                    "  Instance: 0x" + HexDump.ToHex(Instance) + nl +
                    "  ShapeIdMax: " + field_1_shapeIdMax + nl +
                    "  NumIdClusters: " + NumIdClusters + nl +
                    "  NumShapesSaved: " + field_3_numShapesSaved + nl +
                    "  DrawingsSaved: " + field_4_drawingsSaved + nl +
                    "" + field_5_string.ToString();

        }

        public override String ToXml(String tab)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(tab).Append(FormatXmlRecordHeader(GetType().Name, HexDump.ToHex(RecordId), HexDump.ToHex(Version), HexDump.ToHex(Instance)))
                    .Append(tab).Append("\t").Append("<ShapeIdMax>").Append(field_1_shapeIdMax).Append("</ShapeIdMax>\n")
                    .Append(tab).Append("\t").Append("<NumIdClusters>").Append(NumIdClusters).Append("</NumIdClusters>\n")
                    .Append(tab).Append("\t").Append("<NumShapesSaved>").Append(field_3_numShapesSaved).Append("</NumShapesSaved>\n")
                    .Append(tab).Append("\t").Append("<DrawingsSaved>").Append(field_4_drawingsSaved).Append("</DrawingsSaved>\n");
            builder.Append(tab).Append("</").Append(GetType().Name).Append(">\n");
            return builder.ToString();
        }

        /// <summary>
        /// Gets or sets the shape id max.
        /// </summary>
        /// <value>The shape id max.</value>
        public int ShapeIdMax
        {
            get { return field_1_shapeIdMax; }
            set { field_1_shapeIdMax = value; }
        }

        /// <summary>
        /// Gets the Number of id clusters + 1
        /// </summary>
        /// <value>The num id clusters.</value>
        public int NumIdClusters
        {
            get { return field_5_fileIdClusters.Length + 1; }
        }

        /// <summary>
        /// Gets or sets the num shapes saved.
        /// </summary>
        /// <value>The num shapes saved.</value>
        public int NumShapesSaved
        {
            get { return field_3_numShapesSaved; }
            set { field_3_numShapesSaved = value; }
        }


        /// <summary>
        /// Gets or sets the drawings saved.
        /// </summary>
        /// <value>The drawings saved.</value>
        public int DrawingsSaved
        {
            get { return field_4_drawingsSaved; }
            set { field_4_drawingsSaved = value; }
        }


        /// <summary>
        /// Gets or sets the max drawing group id.
        /// </summary>
        /// <value>The max drawing group id.</value>
        public int MaxDrawingGroupId
        {
            get { return maxDgId; }
            set { maxDgId = value; }
        }

        /// <summary>
        /// Gets or sets the file id clusters.
        /// </summary>
        /// <value>The file id clusters.</value>
        public FileIdCluster[] FileIdClusters
        {
            get { return field_5_fileIdClusters; }
            set { field_5_fileIdClusters = value; }
        }

        /// <summary>
        /// Adds the cluster.
        /// </summary>
        /// <param name="dgId">The dg id.</param>
        /// <param name="numShapedUsed">The num shaped used.</param>
        public void AddCluster(int dgId, int numShapedUsed)
        {
            AddCluster(dgId, numShapedUsed, true);
        }


        private class EscherDggRecordComparer : IComparer<FileIdCluster>
        {

            #region IComparer Members

            public int Compare(FileIdCluster f1, FileIdCluster f2)
            {
                if (f1.DrawingGroupId == f2.DrawingGroupId)
                    return 0;
                if (f1.DrawingGroupId < f2.DrawingGroupId)
                    return -1;
                else
                    return +1;
            }

            #endregion
        }
        /// <summary>
        /// Adds the cluster.
        /// </summary>
        /// <param name="dgId">id of the drawing group (stored in the record options)</param>
        /// <param name="numShapedUsed">initial value of the numShapedUsed field</param>
        /// <param name="sort">if set to <c>true</c> if true then sort clusters by drawing group id.(
        /// In Excel the clusters are sorted but in PPT they are not).</param>
        public void AddCluster(int dgId, int numShapedUsed, bool sort)
        {
            List<FileIdCluster> clusters = new List<FileIdCluster>(field_5_fileIdClusters);
            clusters.Add(new FileIdCluster(dgId, numShapedUsed));
            if (sort)
            {
                //ArrayList.Sort is not stable, we need a stable sort ,
                //see test case TestHSSFComment.TestBug56380InsertTooManyComments
                //clusters.Sort(new EscherDggRecordComparer());
                InsertionSort<FileIdCluster>(clusters, new EscherDggRecordComparer());
            }
            maxDgId = Math.Min(maxDgId, dgId);
            field_5_fileIdClusters = clusters.ToArray();
        }

        public static void InsertionSort<T>(List<T> list, IComparer<T> comparison)
        {
            if (list == null)
                throw new ArgumentNullException("list");
            if (comparison == null)
                throw new ArgumentNullException("comparison");

            int count = list.Count;
            for (int j = 1; j < count; j++)
            {
                T key = list[j];

                int i = j - 1;
                for (; i >= 0 && comparison.Compare(list[i], key) > 0; i--)
                {
                    list[i + 1] = list[i];
                }
                list[i + 1] = key;
            }
        }
    }
}