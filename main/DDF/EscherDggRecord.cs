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

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.DDF
{
    using NPOI.Util;
    using System.Linq;
    /// <summary>
    /// This record defines the Drawing groups used for a particular sheet.
    /// </summary>
    public sealed class EscherDggRecord : EscherRecord
    {
        public static short RECORD_ID = unchecked((short) 0xF006);
        public static string RECORD_DESCRIPTION = "MsofbtDgg";

        private int field_1_shapeIdMax;
        // for some reason the number of clusters is actually the real number + 1
        // private int field_2_numIdClusters;
        private int field_3_numShapesSaved;
        private int field_4_drawingsSaved;
        private List<FileIdCluster> field_5_fileIdClusters = new List<FileIdCluster>();
        private int maxDgId;

        public class FileIdCluster
        {
            internal int field_1_drawingGroupId;
            internal int field_2_numShapeIdsUsed;

            public FileIdCluster(int DrawingGroupId, int numShapeIdsUsed)
            {
                this.field_1_drawingGroupId = DrawingGroupId;
                this.field_2_numShapeIdsUsed = numShapeIdsUsed;
            }

            public int DrawingGroupId
            {
                get
                {
                    return field_1_drawingGroupId;
                }
            }

            public int NumShapeIdsUsed
            {
                get
                {
                    return field_2_numShapeIdsUsed;
                }
            }

            public void IncrementUsedShapeId()
            {
                field_2_numShapeIdsUsed++;
            }
        }
        public override int FillFields(byte[] data, int offset, IEscherRecordFactory recordFactory)
        {
            int bytesRemaining = ReadHeader( data, offset );
            int pos            = offset + 8;
            int size           = 0;
            field_1_shapeIdMax     =  LittleEndian.GetInt(data, pos + size);
            size+=4;
            // field_2_numIdClusters = LittleEndian.GetInt( data, pos + size );
            size+=4;
            field_3_numShapesSaved =  LittleEndian.GetInt(data, pos + size);
            size+=4;
            field_4_drawingsSaved  =  LittleEndian.GetInt(data, pos + size);
            size+=4;

            field_5_fileIdClusters.Clear();
            // Can't rely on field_2_numIdClusters
            int numIdClusters = (bytesRemaining-size) / 8;

            for(int i = 0; i < numIdClusters; i++)
            {
                int drawingGroupId = LittleEndian.GetInt( data, pos + size );
                int numShapeIdsUsed = LittleEndian.GetInt( data, pos + size + 4 );
                FileIdCluster fic = new FileIdCluster(drawingGroupId, numShapeIdsUsed);
                field_5_fileIdClusters.Add(fic);
                maxDgId = Math.Max(maxDgId, drawingGroupId);
                size += 8;
            }
            bytesRemaining -= size;
            if(bytesRemaining != 0)
            {
                throw new RecordFormatException("Expecting no remaining data but got " + bytesRemaining + " byte(s).");
            }
            return 8 + size;
        }
        public override int Serialize(int offset, byte[] data, EscherSerializationListener listener)
        {
            listener.BeforeRecordSerialize(offset, RecordId, this);

            int pos = offset;
            LittleEndian.PutShort(data, pos, Options);
            pos += 2;
            LittleEndian.PutShort(data, pos, RecordId);
            pos += 2;
            int remainingBytes = RecordSize - 8;
            LittleEndian.PutInt(data, pos, remainingBytes);
            pos += 4;

            LittleEndian.PutInt(data, pos, field_1_shapeIdMax);
            pos += 4;
            LittleEndian.PutInt(data, pos, NumIdClusters);
            pos += 4;
            LittleEndian.PutInt(data, pos, field_3_numShapesSaved);
            pos += 4;
            LittleEndian.PutInt(data, pos, field_4_drawingsSaved);
            pos += 4;

            foreach(FileIdCluster fic in field_5_fileIdClusters)
            {
                LittleEndian.PutInt(data, pos, fic.DrawingGroupId);
                pos += 4;
                LittleEndian.PutInt(data, pos, fic.NumShapeIdsUsed);
                pos += 4;
            }

            listener.AfterRecordSerialize(pos, RecordId, RecordSize, this);
            return RecordSize;
        }
        public override int RecordSize
        {
            get
            {
                return 8 + 16 + (8 * field_5_fileIdClusters.Count);
            }
        }
        public override short RecordId
        {
            get
            {
                return RECORD_ID;
            }
            set
            {
                
            }
        }
        public override string RecordName
        {
            get
            {
                return "Dgg";
            }
        }

        /// <summary>
        /// Gets or set the next available shape id
        /// </summary>
        /// <returns>the next available shape id</returns>
        public int ShapeIdMax
        {
            get
            {
                return field_1_shapeIdMax;
            }
            set
            {
                this.field_1_shapeIdMax = value;
            }
        }

        /// <summary>
        /// Number of id clusters + 1
        /// </summary>
        /// <returns>the number of id clusters + 1</returns>
        public int NumIdClusters
        {
            get
            {
                return (field_5_fileIdClusters.Count == 0 ? 0 : field_5_fileIdClusters.Count + 1);
            }
        }

        /// <summary>
        /// Gets or set the number of shapes saved
        /// </summary>
        /// <returns>the number of shapes saved</returns>
        public int NumShapesSaved
        {
            get
            {
                return field_3_numShapesSaved;
            }
            set
            {
                this.field_3_numShapesSaved = value;
            }
        }

        /// <summary>
        /// Get or set the number of Drawings saved
        /// </summary>
        /// <returns>the number of Drawings saved</returns>
        public int DrawingsSaved
        {
            get
            {
                return field_4_drawingsSaved;
            }
            set
            {
                this.field_4_drawingsSaved = value;
            }
        }

        /// <summary>
        /// Gets the maximum Drawing group ID
        /// </summary>
        /// <returns>The maximum Drawing group ID</returns>
        public int MaxDrawingGroupId
        {
            get
            {
                return maxDgId;
            }
        }

        /// <summary>
        /// Get or sets the file id clusters
        /// </summary>
        /// <returns>the file id clusters</returns>
        public FileIdCluster[] FileIdClusters
        {
            get
            {
                return field_5_fileIdClusters.ToArray();
            }
            set
            {
                field_5_fileIdClusters.Clear();
                if(value != null)
                {
                    field_5_fileIdClusters.AddRange(value);
                }
            }
        }

        /// <summary>
        /// Add a new cluster
        /// </summary>
        /// <param name="dgId"> id of the Drawing group (stored in the record options)</param>
        /// <param name="numShapedUsed">initial value of the numShapedUsed field</param>
        /// 
        /// <returns>the new <see cref="FileIdCluster"/></returns>
        public FileIdCluster AddCluster(int dgId, int numShapedUsed)
        {
            return AddCluster(dgId, numShapedUsed, true);
        }

        /// <summary>
        /// Add a new cluster
        /// </summary>
        /// <param name="dgId"> id of the Drawing group (stored in the record options)</param>
        /// <param name="numShapedUsed">initial value of the numShapedUsed field</param>
        /// <param name="sort">if true then sort clusters by Drawing group id.(
        /// In Excel the clusters are sorted but in PPT they are not)
        /// </param>
        /// <returns>the new <see cref="FileIdCluster"/></returns>
        public FileIdCluster AddCluster(int dgId, int numShapedUsed, bool sort)
        {
            FileIdCluster ficNew = new FileIdCluster(dgId, numShapedUsed);
            field_5_fileIdClusters.Add(ficNew);
            maxDgId = Math.Min(maxDgId, dgId);
            if(sort)
            {
                SortCluster();
            }

            return ficNew;
        }
        private sealed class EscherDggRecordComparer : IComparer<FileIdCluster>
        {
            #region IComparer Members

            public int Compare(FileIdCluster f1, FileIdCluster f2)
            {
                if(f1.DrawingGroupId == f2.DrawingGroupId)
                    return 0;
                if(f1.DrawingGroupId < f2.DrawingGroupId)
                    return -1;
                else
                    return +1;
            }

            #endregion
        }

        private void SortCluster()
        {
            field_5_fileIdClusters.Sort(new EscherDggRecordComparer());
            //        Collections.sort(field_5_fileIdClusters, new Comparator<FileIdCluster>() {
            //        public int compare(FileIdCluster f1, FileIdCluster f2)
            //    {
            //        int dgDif = f1.DrawingGroupId - f2.DrawingGroupId;
            //        int cntDif = f2.NumShapeIdsUsed - f1.NumShapeIdsUsed;
            //        return (dgDif != 0) ? dgDif : cntDif;
            //    }
            //});
        }

        /// <summary>
        /// Finds the next available (1 based) Drawing group id
        /// </summary>
        /// <returns>the next available Drawing group id</returns>
        public short FindNewDrawingGroupId()
        {
            BitArray bs = new BitArray(field_5_fileIdClusters.Count + 32);
            bs.Set(0, true);
            foreach(FileIdCluster fic in field_5_fileIdClusters)
            {
                bs.Set(fic.DrawingGroupId, true);
            }
            for(var i = 0; i<field_5_fileIdClusters.Count + 32; i++)
            {
                if(!bs.Get(i))
                    return (short)i;
            }
            //return (short) bs.nextClearBit(0);
            throw new InvalidOperationException();
        }

        /// <summary>
        /// Allocates new shape id for the Drawing group
        /// </summary>
        /// <param name="dg">the EscherDgRecord which receives the new shape</param>
        /// <param name="sort">if true then sort clusters by Drawing group id.(
        /// In Excel the clusters are sorted but in PPT they are not)
        /// </param>
        /// 
        /// <returns>a new shape id.</returns>
        public int AllocateShapeId(EscherDgRecord dg, bool sort)
        {
            short DrawingGroupId = dg.DrawingGroupId;
            field_3_numShapesSaved++;

            // check for an existing cluster, which has space available
            // see 2.2.46 OfficeArtIDCL (cspidCur) for the 1024 limitation
            // multiple clusters can belong to the same Drawing group
            FileIdCluster ficAdd = null;
            int index = 1;
            foreach(FileIdCluster fic in field_5_fileIdClusters)
            {
                if(fic.DrawingGroupId == DrawingGroupId
                    && fic.NumShapeIdsUsed < 1024)
                {
                    ficAdd = fic;
                    break;
                }
                index++;
            }

            if(ficAdd == null)
            {
                ficAdd = AddCluster(DrawingGroupId, 0, sort);
                maxDgId = Math.Max(maxDgId, DrawingGroupId);
            }

            int shapeId = index*1024 + ficAdd.NumShapeIdsUsed;
            ficAdd.IncrementUsedShapeId();

            dg.NumShapes = dg.NumShapes + 1;
            dg.LastMSOSPID = shapeId;
            field_1_shapeIdMax = Math.Max(field_1_shapeIdMax, shapeId + 1);

            return shapeId;
        }
        protected object[][] GetAttributeMap()
        {
            List<object> fldIds = new List<object>();
            fldIds.Add("FileId Clusters");
            fldIds.Add(field_5_fileIdClusters.Count);
            foreach(FileIdCluster fic in field_5_fileIdClusters)
            {
                fldIds.Add("Group"+fic.field_1_drawingGroupId);
                fldIds.Add(fic.field_2_numShapeIdsUsed);
            }

            return new object[][] {
                new object[] { "ShapeIdMax", field_1_shapeIdMax },
                new object[] { "NumIdClusters", NumIdClusters },
                new object[] { "NumShapesSaved", field_3_numShapesSaved },
                new object[] { "DrawingsSaved", field_4_drawingsSaved },
                fldIds.ToArray()
            };
        }

        /// <summary>
        /// Returns a <see cref="System.String"/> that represents the current <see cref="System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String"/> that represents the current <see cref="System.Object"/>.
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
            for (int i = 0; i < field_5_fileIdClusters.Count; i++)
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
    }
}
