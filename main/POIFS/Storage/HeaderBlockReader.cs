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
using System.Collections;
using System.IO;

using NPOI.Util;
using NPOI.Util.IO;
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
        private int bigBlockSize = POIFSConstants.BIG_BLOCK_SIZE;

        // number of big block allocation table blocks (int)
        private int _bat_count;

        // start of the property Set block (int index of the property Set
        // chain's first big block)
        private int _property_start;

        // start of the small block allocation table (int index of small
        // block allocation table's first big block)
        private int _sbat_start;

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
    	    // At this point, we don't know how big our
    	    //  block Sizes are
    	    // So, Read the first 32 bytes to check, then
    	    //  Read the rest of the block
            stream.Position = 0;

    	    byte[] blockStart = new byte[32];
    	    int bsCount = IOUtils.ReadFully(stream, blockStart);
    	    if(bsCount != 32) {
    		    AlertShortRead(bsCount);
    	    }
        	
    	    // Figure out our block Size
            //uSectorShift in StructuredStorageHeader struct
    	    if(blockStart[30] == 12) {
    		    bigBlockSize = POIFSConstants.LARGER_BIG_BLOCK_SIZE;
    	    }
            _data = new byte[ bigBlockSize ];
            Array.Copy(blockStart, 0, _data, 0, blockStart.Length);
        	
    	    // Now we can Read the rest of our header
            int byte_count = IOUtils.ReadFully(stream, _data, blockStart.Length, _data.Length - blockStart.Length);
            if (byte_count+bsCount != bigBlockSize) {
    		    AlertShortRead(byte_count);
            }

            // verify signature
            long signature = LittleEndian.GetLong(_data, HeaderBlockConstants._signature_offset);

            if (signature != HeaderBlockConstants._signature)
            {
			    // Is it one of the usual suspects?
        	    byte[] OOXML_FILE_HEADER = POIFSConstants.OOXML_FILE_HEADER;
			    if(_data[0] == OOXML_FILE_HEADER[0] && 
					    _data[1] == OOXML_FILE_HEADER[1] && 
					    _data[2] == OOXML_FILE_HEADER[2] &&
					    _data[3] == OOXML_FILE_HEADER[3]) {
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
                                      + signature + ", expected "
                                      + HeaderBlockConstants._signature);
            }
            //number of SECTs in the FAT chain
            _bat_count = LittleEndian.GetInt(_data, HeaderBlockConstants._bat_count_offset);
            //first SECT in the Directory chain
            _property_start = LittleEndian.GetInt(_data, HeaderBlockConstants._property_start_offset);
            //first SECT in the mini-FAT chain
            _sbat_start = LittleEndian.GetInt(_data, HeaderBlockConstants._sbat_start_offset);
            //first SECT in the DIF chain
            _xbat_start = LittleEndian.GetInt(_data, HeaderBlockConstants._xbat_start_offset);
            //number of SECTs in the DIF chain
            _xbat_count = LittleEndian.GetInt(_data, HeaderBlockConstants._xbat_count_offset);
        }

        /// <summary>
        /// Alerts the short read.
        /// </summary>
        /// <param name="Read">The read.</param>
        private void AlertShortRead(int Read)
        {
    	    if (Read == -1)
    		    //Cant have -1 bytes Read in the error message!
    		    Read = 0;
            String type = " byte" + ((Read == 1) ? ("")
                                                       : ("s"));

            throw new IOException("Unable to Read entire header; "
                                  + Read + type + " Read; expected "
                                  + bigBlockSize + " bytes");
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
                    offset      += LittleEndianConstants.INT_SIZE;
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
        public int BigBlockSize
        {
    	    get{return bigBlockSize;}
        }
    }
}
