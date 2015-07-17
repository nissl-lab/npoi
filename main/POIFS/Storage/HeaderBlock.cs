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

/* ================================================================
 * About NPOI
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * Author's Blog: tonyqus.wordpress.com.cn (wp.tonyqus.cn)
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/


using System;
using System.IO;
using NPOI.POIFS.Common;
using NPOI.POIFS.FileSystem;
using NPOI.Util;
using NPOI.HSSF;



namespace NPOI.POIFS.Storage
{

    /// <summary>
    /// The block containing the archive header
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public class HeaderBlock : HeaderBlockConstants
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(HeaderBlock));
         /**
         * What big block Size the file uses. Most files
         *  use 512 bytes, but a few use 4096
         */
        private POIFSBigBlockSize bigBlockSize;

        // number of big block allocation table blocks (int)
        private int _bat_count;

        // start of the property Set block (int index of the property Set
        // chain's first big block)
        private int _property_start;

        // start of the small block allocation table (int index of small
        // block allocation table's first big block)
        private int _sbat_start;
            /**
         * Number of small block allocation table blocks (int)
         * (Number of MiniFAT Sectors in Microsoft parlance)
         */
        private int _sbat_count;
        // big block index for extension to the big block allocation table
        private int _xbat_start;
        private int _xbat_count;
        private byte[]       _data;

        private static byte _default_value = (byte)0xFF;
        /// <summary>
        /// create a new HeaderBlockReader from an Stream
        /// </summary>
        /// <param name="stream">the source Stream</param>
        public HeaderBlock(Stream stream)
        {
            try
            {
                stream.Position = 0;
                PrivateHeaderBlock(ReadFirst512(stream));
                if (bigBlockSize.GetBigBlockSize() != 512)
                {
                    int rest = bigBlockSize.GetBigBlockSize() - 512;
                    byte[] temp = new byte[rest];
                    IOUtils.ReadFully(stream, temp);
                }

            }
            catch(IOException ex)
            {
                throw ex;
            }
        }

        public HeaderBlock(ByteBuffer buffer)
            : this(IOUtils.ToByteArray(buffer, POIFSConstants.SMALLER_BIG_BLOCK_SIZE))
        {
        }

        public HeaderBlock(byte[] buffer)
        {
            try
            {
                PrivateHeaderBlock(buffer);
            }
            catch (IOException ex)
            {
                throw ex;
            }
        }
        public void PrivateHeaderBlock(byte[] data)
        {
            _data = data;

            long signature = LittleEndian.GetLong(_data, _signature_offset);

            if (signature != _signature)
            {
                byte[] OOXML_FILE_HEADER = POIFSConstants.OOXML_FILE_HEADER;
                if (_data[0] == OOXML_FILE_HEADER[0]
                    && _data[1] == OOXML_FILE_HEADER[1]
                    && _data[2] == OOXML_FILE_HEADER[2]
                    && _data[3] == OOXML_FILE_HEADER[3])
                {
                    throw new OfficeXmlFileException("The supplied data appears to be in the Office 2007+ XML. You are calling the part of POI that deals with OLE2 Office Documents. You need to call a different part of POI to process this data (eg XSSF instead of HSSF)");
                }
                if (_data[0] == 0x09 && _data[1] == 0x00 && // sid=0x0009
                _data[2] == 0x04 && _data[3] == 0x00 && // size=0x0004
                _data[4] == 0x00 && _data[5] == 0x00 && // unused
               (_data[6] == 0x10 || _data[6] == 0x20 || _data[6] == 0x40) &&
                _data[7] == 0x00)
                {
                    // BIFF2 raw stream
                    throw new OldExcelFormatException("The supplied data appears to be in BIFF2 format. " +
                            "HSSF only supports the BIFF8 format, try OldExcelExtractor");
                }
                if (_data[0] == 0x09 && _data[1] == 0x02 && // sid=0x0209
                    _data[2] == 0x06 && _data[3] == 0x00 && // size=0x0006
                    _data[4] == 0x00 && _data[5] == 0x00 && // unused
                   (_data[6] == 0x10 || _data[6] == 0x20 || _data[6] == 0x40) &&
                    _data[7] == 0x00)
                {
                    // BIFF3 raw stream
                    throw new OldExcelFormatException("The supplied data appears to be in BIFF3 format. " +
                            "HSSF only supports the BIFF8 format, try OldExcelExtractor");
                }
                if (_data[0] == 0x09 && _data[1] == 0x04 && // sid=0x0409
                    _data[2] == 0x06 && _data[3] == 0x00 && // size=0x0006
                    _data[4] == 0x00 && _data[5] == 0x00)
                { // unused
                    if (((_data[6] == 0x10 || _data[6] == 0x20 || _data[6] == 0x40) &&
                          _data[7] == 0x00) ||
                        (_data[6] == 0x00 && _data[7] == 0x01))
                    {
                        // BIFF4 raw stream
                        throw new OldExcelFormatException("The supplied data appears to be in BIFF4 format. " +
                                "HSSF only supports the BIFF8 format, try OldExcelExtractor");
                    }
                }

                // Give a generic error if the OLE2 signature isn't found
                throw new NotOLE2FileException("Invalid header signature; read "
                                    + LongToHex(signature) + ", expected "
                                    + LongToHex(_signature) + " - Your file appears "
                                    + "not to be a valid OLE2 document");
            }

            if (_data[30] == 12)
            {
                bigBlockSize = POIFSConstants.LARGER_BIG_BLOCK_SIZE_DETAILS;
            }
            else if (_data[30] == 9)
            {
                bigBlockSize = POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS;
            }
            else
            {
                throw new IOException("Unsupported blocksize  (2^" + _data[30] + "). Expected 2^9 or 2^12.");
            }

            // Setup the fields to read and write the counts and starts
            _bat_count = new IntegerField(HeaderBlockConstants._bat_count_offset, _data).Value;
            _property_start = new IntegerField(HeaderBlockConstants._property_start_offset, _data).Value;
            _sbat_start = new IntegerField(HeaderBlockConstants._sbat_start_offset, _data).Value;
            _sbat_count = new IntegerField(HeaderBlockConstants._sbat_block_count_offset, _data).Value;
            _xbat_start = new IntegerField(HeaderBlockConstants._xbat_start_offset, _data).Value;
            _xbat_count = new IntegerField(HeaderBlockConstants._xbat_count_offset, _data).Value;

        }

        public HeaderBlock(POIFSBigBlockSize bigBlockSize)
        {
            this.bigBlockSize = bigBlockSize;

            _data = new byte[POIFSConstants.SMALLER_BIG_BLOCK_SIZE];

            //fill the array.
            for (int i = 0; i < _data.Length; i++)
                _data[i] = _default_value;

            new LongField(_signature_offset, _signature, _data);
            new IntegerField(0x08, 0, _data);
            new IntegerField(0x0c, 0, _data);
            new IntegerField(0x10, 0, _data);
            new IntegerField(0x14, 0, _data);

            new ShortField((int)0x18, (short)0x3b, ref _data);
            new ShortField((int)0x1a, (short)0x3, ref _data);
            new ShortField((int)0x1c, (short)-2, ref _data);

            new ShortField(0x1e, bigBlockSize.GetHeaderValue(), ref _data);
            new IntegerField(0x20, 0x6, _data);
            new IntegerField(0x24, 0, _data);
            new IntegerField(0x28, 0, _data);
            new IntegerField(0x34, 0, _data);
            new IntegerField(0x38, 0x1000, _data);

            _bat_count = 0;
            _sbat_count = 0;
            _xbat_count = 0;
            _property_start = POIFSConstants.END_OF_CHAIN;
            _sbat_start = POIFSConstants.END_OF_CHAIN;
            _xbat_start = POIFSConstants.END_OF_CHAIN;

        }


        private static byte[] ReadFirst512(Stream stream)
        {
            // Grab the first 512 bytes
            // (For 4096 sized blocks, the remaining 3584 bytes are zero)
            byte[] data = new byte[512];
            int bsCount = IOUtils.ReadFully(stream, data);
            if (bsCount != 512)
            {
                AlertShortRead(bsCount, 512);
            }
            return data;
        }
        private static String LongToHex(long value)
        {
            return new String(HexDump.LongToHex(value));
        }
        /// <summary>
        /// Alerts the short read.
        /// </summary>
        /// <param name="read">The read.</param>
        /// <param name="expectedReadSize">The expected size.</param>
        private static IOException AlertShortRead(int read, int expectedReadSize)
        {
            if (read < 0)
                //Cant have -1 bytes Read in the error message!
                read = 0;
            String type = " byte" + ((read == 1) ? (""): ("s"));

            return new IOException("Unable to Read entire header; "
                                  + read + type + " Read; expected "
                                  + expectedReadSize + " bytes");
        }

        /// <summary>
        /// Get start of Property Table
        /// </summary>
        /// <value>the index of the first block of the Property Table</value>
        public int PropertyStart
        {
            get{return _property_start;}
            set { _property_start = value; }
        }

        /// <summary>
        /// Gets start of small block allocation table
        /// </summary>
        /// <value>The SBAT start.</value>
        public int SBATStart
        {
            get{return _sbat_start;}
            set { _sbat_start = value; }
        }

        /// <summary>
        /// Gets number of BAT blocks
        /// </summary>
        /// <value>The BAT count.</value>
        public int SBATCount
        {
            get{return _sbat_count;}
            set { _sbat_count = value; }
        }


        public int SBATBlockCount
        {
            get { return _sbat_count; }
            set { _sbat_count = value; }
        }


        public int BATCount
        {
            get { return _bat_count; }
            set { _bat_count = value; }
        }
        /// <summary>
        /// Gets the BAT array.
        /// </summary>
        /// <value>The BAT array.</value>
        public int [] BATArray
        {
            get{
               // int[] result = new int[HeaderBlockConstants._max_bats_in_header];
               // int offset = HeaderBlockConstants._bat_array_offset;
                int[] result = new int[Math.Min(_bat_count, _max_bats_in_header)];
                int offset = _bat_array_offset;

                for (int j = 0; j < result.Length; j++)
                {
                    result[ j ] = LittleEndian.GetInt(_data, offset);
                    offset += LittleEndianConsts.INT_SIZE;
                }
                return result;
            }
            set
            {
                int count = Math.Min(value.Length, _max_bats_in_header);
                int blank = _max_bats_in_header - count;

                int offset = _bat_array_offset;
                
                for (int i = 0; i < count; i++)
                {
                    LittleEndian.PutInt(_data, offset ,value[i]);
                    offset += LittleEndianConsts.INT_SIZE;
                }

                for (int i = 0; i < blank; i++)
                {
                    LittleEndian.PutInt(_data, offset, POIFSConstants.UNUSED_BLOCK);
                    offset += LittleEndianConsts.INT_SIZE;
                }
            }
        }

        /// <summary>
        /// Gets the XBAT count.
        /// </summary>
        /// <value>The XBAT count.</value>
        /// @return XBAT count
        public int XBATCount
        {
            get{return _xbat_count;}
            set { _xbat_count = value; }
        }

        /// <summary>
        /// Gets the index of the XBAT.
        /// </summary>
        /// <value>The index of the XBAT.</value>
        public int XBATIndex
        {
            get{return _xbat_start;}
            set { _xbat_count = value; }
        }

        public int XBATStart
        {
            set { _xbat_start = value; }
        }
        
        /// <summary>
        /// Gets The Big Block Size, normally 512 bytes, sometimes 4096 bytes
        /// </summary>
        /// <value>The size of the big block.</value>
        /// @return 
        public POIFSBigBlockSize BigBlockSize
        {
            get{return bigBlockSize;}
        }

        //public void Write(Stream stream)
        //{
        //    try
        //    {
        //        new IntegerField(_bat_count_offset, _bat_count, _data);
        //        new IntegerField(_property_start_offset, _property_start, _data);
        //        new IntegerField(_sbat_start_offset, _sbat_start, _data);
        //        new IntegerField(_sbat_block_count_offset, _sbat_count, _data);
        //        new IntegerField(_xbat_start_offset, _xbat_start, _data);
        //        new IntegerField(_xbat_count_offset, _xbat_count, _data);

        //        stream.Write(_data, 0, 512);

        //        for (int i = POIFSConstants.SMALLER_BIG_BLOCK_SIZE; i < bigBlockSize.GetBigBlockSize(); i++)
        //        {
        //            //stream.Write(Write(0);
        //            stream.WriteByte(0);
        //        }
        //    }
        //    catch (IOException ex)
        //    {
        //        throw ex;
        //    }
        //}

        public void WriteData(Stream stream)
        {
            try
            {
                new IntegerField(_bat_count_offset, _bat_count, _data);
                new IntegerField(_property_start_offset, _property_start, _data);
                new IntegerField(_sbat_start_offset, _sbat_start, _data);
                new IntegerField(_sbat_block_count_offset, _sbat_count, _data);
                new IntegerField(_xbat_start_offset, _xbat_start, _data);
                new IntegerField(_xbat_count_offset, _xbat_count, _data);

                stream.Write(_data, 0, 512);

                for (int i = POIFSConstants.SMALLER_BIG_BLOCK_SIZE; i < bigBlockSize.GetBigBlockSize(); i++)
                {
                    //stream.Write(Write(0);
                    stream.WriteByte(0);
                }
            }
            catch (IOException ex)
            {
                throw ex;
            }
        }
    }
}
