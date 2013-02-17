/*
* Licensed to the Apache Software Foundation (ASF) under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed under the License is distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations under the License.
*/
namespace NPOI.DDF
{
    using System;
    using System.IO;
    using System.Text;
    using System.Drawing;
    using NPOI.Util;
    using ICSharpCode.SharpZipLib.Zip.Compression;
    using ICSharpCode.SharpZipLib.Zip.Compression.Streams;

    /// <summary>
    /// @author Daniel Noll
    /// </summary>
    public class EscherPictBlip : EscherBlipRecord
    {
        private static POILogger log = POILogFactory.GetLogger(typeof(EscherPictBlip));

        public const short RECORD_ID_EMF = unchecked((short)0xF018) + 2;
        public const short RECORD_ID_WMF = unchecked((short)0xF018) + 3;
        public const short RECORD_ID_PICT = unchecked((short)0xF018) + 4;

        private const int HEADER_SIZE = 8;

        private byte[] field_1_UID;
        private int field_2_cb;
        private int field_3_rcBounds_x1;
        private int field_3_rcBounds_y1;
        private int field_3_rcBounds_x2;
        private int field_3_rcBounds_y2;
        private int field_4_ptSize_w;
        private int field_4_ptSize_h;
        private int field_5_cbSave;
        private byte field_6_fCompression;
        private byte field_7_fFilter;

        private byte[] raw_pictureData;

        /// <summary>
        /// This method deSerializes the record from a byte array.
        /// </summary>
        /// <param name="data">The byte array containing the escher record information</param>
        /// <param name="offset">The starting offset into</param>
        /// <param name="recordFactory">May be null since this is not a container record.</param>
        /// <returns>
        /// The number of bytes Read from the byte array.
        /// </returns>
        public override int FillFields(byte[] data, int offset, IEscherRecordFactory recordFactory)
        {
            int bytesAfterHeader = ReadHeader(data, offset);
            int pos = offset + HEADER_SIZE;

            field_1_UID = new byte[16];
            Array.Copy(data, pos, field_1_UID, 0, 16); pos += 16;
            field_2_cb = LittleEndian.GetInt(data, pos); pos += 4;
            field_3_rcBounds_x1 = LittleEndian.GetInt(data, pos); pos += 4;
            field_3_rcBounds_y1 = LittleEndian.GetInt(data, pos); pos += 4;
            field_3_rcBounds_x2 = LittleEndian.GetInt(data, pos); pos += 4;
            field_3_rcBounds_y2 = LittleEndian.GetInt(data, pos); pos += 4;
            field_4_ptSize_w = LittleEndian.GetInt(data, pos); pos += 4;
            field_4_ptSize_h = LittleEndian.GetInt(data, pos); pos += 4;
            field_5_cbSave = LittleEndian.GetInt(data, pos); pos += 4;
            field_6_fCompression = data[pos]; pos++;
            field_7_fFilter = data[pos]; pos++;

            raw_pictureData = new byte[field_5_cbSave];
            Array.Copy(data, pos, raw_pictureData, 0, field_5_cbSave);

            // 0 means DEFLATE compression
            // 0xFE means no compression
            if (field_6_fCompression == 0)
            {
                field_pictureData = InflatePictureData(raw_pictureData);
            }
            else
            {
                field_pictureData = raw_pictureData;
            }

            return bytesAfterHeader + HEADER_SIZE;
        }

        /// <summary>
        /// Serializes the record to an existing byte array.
        /// </summary>
        /// <param name="offset">the offset within the byte array</param>
        /// <param name="data">the data array to Serialize to</param>
        /// <param name="listener">a listener for begin and end serialization events.</param>
        /// <returns>the number of bytes written.</returns>
        public override int Serialize(int offset, byte[] data, EscherSerializationListener listener)
        {
            listener.BeforeRecordSerialize(offset, RecordId, this);

            int pos = offset;
            LittleEndian.PutShort(data, pos, Options); pos += 2;
            LittleEndian.PutShort(data, pos, RecordId); pos += 2;
            LittleEndian.PutInt(data, 0, RecordSize - HEADER_SIZE); pos += 4;

            Array.Copy(field_1_UID, 0, data, pos, 16); pos += 16;
            LittleEndian.PutInt(data, pos, field_2_cb); pos += 4;
            LittleEndian.PutInt(data, pos, field_3_rcBounds_x1); pos += 4;
            LittleEndian.PutInt(data, pos, field_3_rcBounds_y1); pos += 4;
            LittleEndian.PutInt(data, pos, field_3_rcBounds_x2); pos += 4;
            LittleEndian.PutInt(data, pos, field_3_rcBounds_y2); pos += 4;
            LittleEndian.PutInt(data, pos, field_4_ptSize_w); pos += 4;
            LittleEndian.PutInt(data, pos, field_4_ptSize_h); pos += 4;
            LittleEndian.PutInt(data, pos, field_5_cbSave); pos += 4;
            data[pos] = field_6_fCompression; pos++;
            data[pos] = field_7_fFilter; pos++;

            Array.Copy(raw_pictureData, 0, data, pos, raw_pictureData.Length);

            listener.AfterRecordSerialize(offset + RecordSize, RecordId, RecordSize, this);
            return HEADER_SIZE + 16 + 1 + raw_pictureData.Length;
        }

        /// <summary>
        /// Decompresses the provided data, returning the inflated result.
        /// </summary>
        /// <param name="data">the deflated picture data.</param>
        /// <returns>the inflated picture data.</returns>
        private static byte[] InflatePictureData(byte[] data)
        {
            using (MemoryStream out1 = new MemoryStream())
            {
                try
                {
                    using (MemoryStream ms = new MemoryStream(data))
                    {
                        Inflater inflater = new Inflater(false);
                        using (InflaterInputStream in1 = new InflaterInputStream(ms, inflater))
                        {
                            byte[] buf = new byte[4096];
                            int ReadBytes;
                            while ((ReadBytes = in1.Read(buf, 0, buf.Length)) > 0)
                            {
                                out1.Write(buf, 0, ReadBytes);
                            }
                            return out1.ToArray();
                        }
                    }
                }
                catch (IOException e)
                {
                    log.Log(POILogger.INFO, "Possibly corrupt compression or non-compressed data", e);
                    return data;
                }
            }
        }

        /// <summary>
        /// Returns the number of bytes that are required to Serialize this record.
        /// </summary>
        /// <value>Number of bytes</value>
        public override int RecordSize
        {
            get { return 8 + 50 + raw_pictureData.Length; }
        }

        /// <summary>
        /// Gets or sets the UID.
        /// </summary>
        /// <value>The UID.</value>
        public byte[] UID
        {
            get { return field_1_UID; }
            set { this.field_1_UID = value; }
        }

        /// <summary>
        /// Gets or sets the size of the uncompressed.
        /// </summary>
        /// <value>The size of the uncompressed.</value>
        public int UncompressedSize
        {
            get { return field_2_cb; }
            set { field_2_cb = value; }
        }

        /// <summary>
        /// Gets or sets the bounds.
        /// </summary>
        /// <value>The bounds.</value>
        public Rectangle Bounds
        {
            get
            {
                return new Rectangle(field_3_rcBounds_x1,
                                     field_3_rcBounds_y1,
                                     field_3_rcBounds_x2 - field_3_rcBounds_x1,
                                     field_3_rcBounds_y2 - field_3_rcBounds_y1);
            }
            set
            {
                field_3_rcBounds_x1 = value.X;
                field_3_rcBounds_y1 = value.Y;
                field_3_rcBounds_x2 = value.X + value.Width;
                field_3_rcBounds_y2 = value.Y + value.Height;
            }
        }

        /// <summary>
        /// Gets or sets the size EMU.
        /// </summary>
        /// <value>The size EMU.</value>
        public Size SizeEMU
        {
            get { return new Size(field_4_ptSize_w, field_4_ptSize_h); }
            set
            {
                field_4_ptSize_w = value.Width;
                field_4_ptSize_h = value.Height;
            }
        }

        /// <summary>
        /// Gets or sets the size of the compressed.
        /// </summary>
        /// <value>The size of the compressed.</value>
        public int CompressedSize
        {
            get { return field_5_cbSave; }
            set { field_5_cbSave = value; }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is compressed.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is compressed; otherwise, <c>false</c>.
        /// </value>
        public bool IsCompressed
        {
            get { return (field_6_fCompression == 0); }
            set{field_6_fCompression = value ? (byte)0 : (byte)0xFE;}
        }


        // filtering is always 254 according to available docs, so no point giving it a Setter method.

        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override String ToString()
        {
            String nl = Environment.NewLine;

            String extraData;
            using (MemoryStream b = new MemoryStream())
            {
                try
                {
                    HexDump.Dump(this.field_pictureData, 0, b, 0);
                    extraData = b.ToString();
                }
                catch (Exception e)
                {
                    extraData = e.ToString();
                }
                return GetType().Name + ":" + nl +
                        "  RecordId: 0x" + HexDump.ToHex(RecordId) + nl +
                        "  Version: 0x" + HexDump.ToHex(Version) + '\n' +
                        "  Instance: 0x" + HexDump.ToHex(Instance) + '\n' +
                        "  UID: 0x" + HexDump.ToHex(field_1_UID) + nl +
                        "  Uncompressed Size: " + HexDump.ToHex(field_2_cb) + nl +
                        "  Bounds: " + Bounds + nl +
                        "  Size in EMU: " + SizeEMU + nl +
                        "  Compressed Size: " + HexDump.ToHex(field_5_cbSave) + nl +
                        "  Compression: " + HexDump.ToHex(field_6_fCompression) + nl +
                        "  Filter: " + HexDump.ToHex(field_7_fFilter) + nl +
                        "  Extra Data:" + nl + extraData;
            }
        }

        public override String ToXml(String tab)
        {
            String extraData = "";
            StringBuilder builder = new StringBuilder();
            builder.Append(tab).Append(FormatXmlRecordHeader(GetType().Name, HexDump.ToHex(RecordId), HexDump.ToHex(Version), HexDump.ToHex(Instance)))
                    .Append(tab).Append("\t").Append("<UID>0x").Append(HexDump.ToHex(field_1_UID)).Append("</UID>\n")
                    .Append(tab).Append("\t").Append("<UncompressedSize>0x").Append(HexDump.ToHex(field_2_cb)).Append("</UncompressedSize>\n")
                    .Append(tab).Append("\t").Append("<Bounds>").Append(Bounds).Append("</Bounds>\n")
                    .Append(tab).Append("\t").Append("<SizeInEMU>").Append(SizeEMU).Append("</SizeInEMU>\n")
                    .Append(tab).Append("\t").Append("<CompressedSize>0x").Append(HexDump.ToHex(field_5_cbSave)).Append("</CompressedSize>\n")
                    .Append(tab).Append("\t").Append("<Compression>0x").Append(HexDump.ToHex(field_6_fCompression)).Append("</Compression>\n")
                    .Append(tab).Append("\t").Append("<Filter>0x").Append(HexDump.ToHex(field_7_fFilter)).Append("</Filter>\n")
                    .Append(tab).Append("\t").Append("<ExtraData>").Append(extraData).Append("</ExtraData>\n");
            builder.Append(tab).Append("</").Append(GetType().Name).Append(">\n");
            return builder.ToString();
        }

    }

}