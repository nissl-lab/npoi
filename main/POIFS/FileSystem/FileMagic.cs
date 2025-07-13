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

namespace NPOI.POIFS.FileSystem
{
    using NPOI.POIFS.Common;
    using NPOI.POIFS.Storage;

    using NPOI.Util;

    public enum FileMagic
    {
        OLE2,
        OOXML,
        XML,
        BIFF2,
        BIFF3,
        BIFF4,
        MSWRITE,
        RTF,
        PDF,
        UNKNOWN,
    }
    /// <summary>
    /// The file magic number, i.e. the file identification based on the first bytes
    /// of the file
    /// </summary>
    public class FileMagicContainer
    {
        /** OLE2 / BIFF8+ stream used for Office 97 and higher documents */
        public static FileMagicContainer OLE2 = new FileMagicContainer(HeaderBlockConstants._signature);
        /** OOXML / ZIP stream */
        public static FileMagicContainer OOXML = new FileMagicContainer(POIFSConstants.OOXML_FILE_HEADER);
        /** XML file */
        public static FileMagicContainer XML = new FileMagicContainer(POIFSConstants.RAW_XML_FILE_HEADER);
        /** BIFF2 raw stream - for Excel 2 */
        public static FileMagicContainer BIFF2 = new FileMagicContainer(new byte[]
        {
            0x09, 0x00, // sid=0x0009
            0x04, 0x00, // size=0x0004
            0x00, 0x00, // unused
            0x70, 0x00  // 0x70 = multiple values
        });
        /** BIFF3 raw stream - for Excel 3 */
        public static FileMagicContainer BIFF3 = new FileMagicContainer(new byte[]
        {
            0x09, 0x02, // sid=0x0209
            0x06, 0x00, // size=0x0006
            0x00, 0x00, // unused
            0x70, 0x00  // 0x70 = multiple values
        });
        /** BIFF4 raw stream - for Excel 4 */
        public static FileMagicContainer BIFF4 = new FileMagicContainer(new byte[]
        {
            0x09, 0x04, // sid=0x0409
            0x06, 0x00, // size=0x0006
            0x00, 0x00, // unused
            0x70, 0x00  // 0x70 = multiple values
        },new byte[]
        {
            0x09, 0x04, // sid=0x0409
            0x06, 0x00, // size=0x0006
            0x00, 0x00, // unused
            0x00, 0x01
        });
        /** Old MS Write raw stream */
        public static FileMagicContainer MSWRITE = new FileMagicContainer(
            new byte[] {0x31, (byte)0xbe, 0x00, 0x00 },
            new byte[] {0x32, (byte)0xbe, 0x00, 0x00 }
        );
        /** RTF document */
        public static FileMagicContainer RTF = new FileMagicContainer("{\\rtf");
        /** PDF document */
        public static FileMagicContainer PDF = new FileMagicContainer("%PDF");
        // keep UNKNOWN always as last enum!
        /** UNKNOWN magic */
        public static FileMagicContainer UNKNOWN = new FileMagicContainer(Array.Empty<byte>());

        byte[][] magic;

        public FileMagicContainer(long magic)
        {
            this.magic = new byte[1][];
            this.magic[0] = new byte[8];
            LittleEndian.PutLong(this.magic[0], 0, magic);
        }

        FileMagicContainer(byte[] magic)
        {
            this.magic = new byte[1][];
            this.magic[0] = magic;
        }
        FileMagicContainer(byte[] m1, byte[] m2)
        {
            this.magic = new byte[2][];
            this.magic[0] = m1;
            this.magic[1] = m2;
        }
        FileMagicContainer(string magic)
            : this(Encoding.GetEncoding(LocaleUtil.CHARSET_1252).GetBytes(magic))
        {

        }
        private static readonly Dictionary<FileMagic, FileMagicContainer> Values = 
            new Dictionary<FileMagic, FileMagicContainer>(){
            { FileMagic.OLE2, OLE2 },
            { FileMagic.OOXML , OOXML },
            { FileMagic.XML , XML },
            { FileMagic.BIFF2 , BIFF2 },
            { FileMagic.BIFF3 , BIFF3 },
            { FileMagic.BIFF4 , BIFF4 },
            { FileMagic.MSWRITE , MSWRITE },
            { FileMagic.RTF , RTF },
            { FileMagic.PDF , PDF },
            { FileMagic.UNKNOWN , UNKNOWN },
        };
        public static FileMagic ValueOf(byte[] magic)
        {
            foreach(var fm in Values)
            {
                int i=0;
                bool found = true;
                foreach(byte[] ma in fm.Value.magic)
                {
                    foreach(byte m in ma)
                    {
                        byte d = magic[i++];
                        if(!(d == m || (m == 0x70 && (d == 0x10 || d == 0x20 || d == 0x40))))
                        {
                            found = false;
                            break;
                        }
                    }
                    if(found)
                    {
                        return fm.Key;
                    }
                }
            }
            return FileMagic.UNKNOWN;
        }

        /**
         * Get the file magic of the supplied InputStream (which MUST
         *  support mark and reset).<p>
         *
         * If unsure if your InputStream does support mark / reset,
         *  use {@link #prepareToCheckMagic(InputStream)} to wrap it and make
         *  sure to always use that, and not the original!<p>
         *
         * Even if this method returns {@link FileMagic#UNKNOWN} it could potentially mean,
         *  that the ZIP stream has leading junk bytes
         *
         * @param inp An InputStream which supports either mark/reset
         */
        public static FileMagic ValueOf(InputStream inp)
        {
            if(!inp.MarkSupported())
            {
                throw new IOException("getFileMagic() only operates on streams which support mark(int)");
            }

            // Grab the first 8 bytes
            byte[] data = IOUtils.PeekFirst8Bytes(inp);

            return FileMagicContainer.ValueOf(data);
        }

        public static FileMagic ValueOf(Stream inp)
        {

            // Grab the first 8 bytes
            byte[] data = IOUtils.PeekFirstNBytes(inp, 8);

            return FileMagicContainer.ValueOf(data);
        }
        /**
         * Checks if an {@link InputStream} can be reseted (i.e. used for checking the header magic) and wraps it if not
         *
         * @param stream stream to be checked for wrapping
         * @return a mark enabled stream
         */
        public static InputStream PrepareToCheckMagic(InputStream stream)
        {
            if(stream.MarkSupported())
            {
                return stream;
            }
            // we used to process the data via a PushbackInputStream, but user code could provide a too small one
            // so we use a BufferedInputStream instead now
            return new BufferedInputStream(stream);
        }
    }
}
