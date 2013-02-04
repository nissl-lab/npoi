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

using NPOI.Util;

using NPOI.POIFS.Common;
using NPOI.POIFS.FileSystem;


namespace NPOI.POIFS.Storage
{

    /// <summary>
    /// The block containing the archive header
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public class HeaderBlockReader
    {
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

        /// <summary>
        /// create a new HeaderBlockReader from an Stream
        /// </summary>
        /// <param name="stream">the source Stream</param>
        public HeaderBlockReader(Stream stream)
        {
            stream.Position = 0;
            this._data = ReadFirst512(stream);

            // verify signature
            long signature = LittleEndian.GetLong(_data, HeaderBlockConstants._signature_offset);

            if (signature != HeaderBlockConstants._signature)
            {
                // Is it one of the usual suspects?
                byte[] OOXML_FILE_HEADER = POIFSConstants.OOXML_FILE_HEADER;
                if (_data[0] == OOXML_FILE_HEADER[0] &&
                        _data[1] == OOXML_FILE_HEADER[1] &&
                        _data[2] == OOXML_FILE_HEADER[2] &&
                        _data[3] == OOXML_FILE_HEADER[3])
                {
                    throw new OfficeXmlFileException("The supplied data appears to be in the Office 2007+ XML. POI only supports OLE2 Office documents");
                }
                if ((signature & unchecked((long)0xFF8FFFFFFFFFFFFFL)) == 0x0010000200040009L)
                {
                    // BIFF2 raw stream starts with BOF (sid=0x0009, size=0x0004, data=0x00t0)
                    throw new ArgumentException("The supplied data appears to be in BIFF2 format.  "
                            + "POI only supports BIFF8 format");
                }

                // Give a generic error
                throw new IOException("Invalid header signature; Read "
                                      + LongToHex(signature) + ", expected "
                                      + LongToHex(HeaderBlockConstants._signature));
            }

            // Figure out our block size
            if (_data[30] == 12)
            {
                this.bigBlockSize = POIFSConstants.LARGER_BIG_BLOCK_SIZE_DETAILS;
            }
            else if (_data[30] == 9)
            {
                this.bigBlockSize = POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS;
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

            // Fetch the rest of the block if needed
            if (bigBlockSize.GetBigBlockSize() != 512)
            {
                int rest = bigBlockSize.GetBigBlockSize() - 512;
                byte[] tmp = new byte[rest];
                IOUtils.ReadFully(stream, tmp);
            }
        }
	

        private byte[] ReadFirst512(Stream stream)
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
        /// <param name="expectedReadSize">expected size to read</param>
        private void AlertShortRead(int read,int expectedReadSize)
        {
            if (read < 0)
    		    //Cant have -1 bytes Read in the error message!
    		    read = 0;
            String type = " byte" + ((read == 1) ? (""): ("s"));

            throw new IOException("Unable to Read entire header; "
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
        }

        /// <summary>
        /// Gets start of small block allocation table
        /// </summary>
        /// <value>The SBAT start.</value>
        public int SBATStart
        {
            get{return _sbat_start;}
        }

        /// <summary>
        /// Gets number of BAT blocks
        /// </summary>
        /// <value>The BAT count.</value>
        public int BATCount
        {
            get{return _bat_count;}
        }

        /// <summary>
        /// Gets the BAT array.
        /// </summary>
        /// <value>The BAT array.</value>
        public int [] BATArray
        {
            get{
                int[] result = new int[HeaderBlockConstants._max_bats_in_header];
                int offset = HeaderBlockConstants._bat_array_offset;

                for (int j = 0; j < HeaderBlockConstants._max_bats_in_header; j++)
                {
                    result[ j ] = LittleEndian.GetInt(_data, offset);
                    offset      += LittleEndianConsts.INT_SIZE;
                }
                return result;
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
        }

        /// <summary>
        /// Gets the index of the XBAT.
        /// </summary>
        /// <value>The index of the XBAT.</value>
        public int XBATIndex
        {
            get{return _xbat_start;}
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
    }
}
