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

using NPOI.Util;
using System;
using System.IO;
using Ionic.Zip;
using Ionic.Zlib;
namespace NPOI.HWPF.UserModel
{

    /**
     * Represents embedded picture extracted from Word Document
     * @author Dmitry Romanov
     */
    public class Picture
    {
        //private static POILogger log = POILogFactory.GetLogger(Picture.class);

        //  public static int FILENAME_OFFSET = 0x7C;
        //  public static int FILENAME_SIZE_OFFSET = 0x6C;
        static int PICF_OFFSET = 0x0;
        static int PICT_HEADER_OFFSET = 0x4;
        static int MFPMM_OFFSET = 0x6;
        static int PICF_SHAPE_OFFSET = 0xE;
        static int PICMD_OFFSET = 0x1C;
        static int UNKNOWN_HEADER_SIZE = 0x49;

        public static byte[] GIF = new byte[] { (byte)'G', (byte)'I', (byte)'F' };
        public static byte[] PNG = new byte[] { (byte)0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        public static byte[] JPG = new byte[] { (byte)0xFF, (byte)0xD8 };
        public static byte[] BMP = new byte[] { (byte)'B', (byte)'M' };
        public static byte[] TIFF = new byte[] { 0x49, 0x49, 0x2A, 0x00 };
        public static byte[] TIFF1 = new byte[] { 0x4D, 0x4D, 0x00, 0x2A };

        public static byte[] EMF = { 0x01, 0x00, 0x00, 0x00 };
        public static byte[] WMF1 = { (byte)0xD7, (byte)0xCD, (byte)0xC6, (byte)0x9A, 0x00, 0x00 };
        public static byte[] WMF2 = { 0x01, 0x00, 0x09, 0x00, 0x00, 0x03 }; // Windows 3.x
        // TODO: DIB, PICT

        public static byte[] IHDR = new byte[] { (byte)'I', (byte)'H', (byte)'D', (byte)'R' };

        public static byte[] COMPRESSED1 = { (byte)0xFE, 0x78, (byte)0xDA };
        public static byte[] COMPRESSED2 = { (byte)0xFE, 0x78, (byte)0x9C };

        private int dataBlockStartOfsset;
        private int pictureBytesStartOffset;
        private int dataBlockSize;
        private int size;
        //  private String fileName;
        private byte[] rawContent;
        private byte[] content;
        private byte[] _dataStream;
        private int aspectRatioX;
        private int aspectRatioY;
        private int height = -1;
        private int width = -1;


        public Picture(int dataBlockStartOfsset, byte[] _dataStream, bool FillBytes)
        {
            this._dataStream = _dataStream;
            this.dataBlockStartOfsset = dataBlockStartOfsset;
            this.dataBlockSize = LittleEndian.GetInt(_dataStream, dataBlockStartOfsset);
            this.pictureBytesStartOffset = GetPictureBytesStartOffset(dataBlockStartOfsset, _dataStream, dataBlockSize);
            this.size = dataBlockSize - (pictureBytesStartOffset - dataBlockStartOfsset);

            if (size < 0)
            {

            }

            this.aspectRatioX = ExtractAspectRatioX(_dataStream, dataBlockStartOfsset);
            this.aspectRatioY = ExtractAspectRatioY(_dataStream, dataBlockStartOfsset);

            if (FillBytes)
            {
                FillImageContent();
            }
        }

        public Picture(byte[] _dataStream)
        {
            this._dataStream = _dataStream;
            this.dataBlockStartOfsset = 0;
            this.dataBlockSize = _dataStream.Length;
            this.pictureBytesStartOffset = 0;
            this.size = _dataStream.Length;
        }

        private void FillWidthHeight()
        {
            String ext = SuggestFileExtension();
            // trying to extract width and height from pictures content:
            if ("jpg".Equals(ext,StringComparison.InvariantCultureIgnoreCase))
            {
                FillJPGWidthHeight();
            }
            else if ("png".Equals(ext,StringComparison.InvariantCultureIgnoreCase))
            {
                FillPNGWidthHeight();
            }
        }

        private static int ExtractAspectRatioX(byte[] _dataStream, int dataBlockStartOffSet)
        {
            return LittleEndian.GetShort(_dataStream, dataBlockStartOffSet + 0x20) / 10;
        }

        private static int ExtractAspectRatioY(byte[] _dataStream, int dataBlockStartOffSet)
        {
            return LittleEndian.GetShort(_dataStream, dataBlockStartOffSet + 0x22) / 10;
        }

        /**
         * Tries to suggest a filename: hex representation of picture structure offset in "Data" stream plus extension that
         * is tried to determine from first byte of picture's content.
         *
         * @return suggested file name
         */
        public String SuggestFullFileName()
        {
            String fileExt = SuggestFileExtension();
            return StringUtil.ToHexString(dataBlockStartOfsset) + (fileExt.Length > 0 ? "." + fileExt : "");
        }

        /**
         * Writes Picture's content bytes to specified OutputStream.
         * Is useful when there is need to write picture bytes directly to stream, omitting its representation in
         * memory as distinct byte array.
         *
         * @param out a stream to write to
         * @throws IOException if some exception is occured while writing to specified out
         */
        public void WriteImageContent(Stream out1)
        {
            if (rawContent != null && rawContent.Length > 0)
            {
                out1.Write(rawContent, 0, size);
            }
            else
            {
                out1.Write(_dataStream, dataBlockStartOfsset, size);
            }
        }

        /**
         * @return picture's content as byte array
         */
        public byte[] GetContent()
        {
            if (content == null || content.Length <= 0)
            {
                FillImageContent();
            }
            return content;
        }

        public byte[] GetRawContent()
        {
            if (rawContent == null || rawContent.Length <= 0)
            {
                FillRawImageContent();
            }
            return rawContent;
        }

        /**
         *
         * @return size in bytes of the picture
         */
        public int Size
        {
            get
            {
                return size;
            }
        }

        /**
         * returns horizontal aspect ratio for picture provided by user
         */
        public int AspectRatioX
        {
            get
            {
                return aspectRatioX;
            }
        }
        /**
         * returns vertical aspect ratio for picture provided by user
         */
        public int AspectRatioY
        {
            get
            {
                return aspectRatioY;
            }
        }

        /**
         * tries to suggest extension for picture's file by matching signatures of popular image formats to first bytes
         * of picture's contents
         * @return suggested file extension
         */
        public String SuggestFileExtension()
        {
            String extension = SuggestFileExtension(_dataStream, pictureBytesStartOffset);
            if ("".Equals(extension))
            {
                // May be compressed.  Get the uncompressed content and inspect that.
                extension = SuggestFileExtension(GetContent(), 0);
            }
            return extension;
        }

        /**
         * Returns the mime type for the image
         */
        public String MimeType
        {
            get
            {
                String extension = SuggestFileExtension();
                if ("jpg".Equals(extension))
                {
                    return "image/jpeg";
                }
                if ("png".Equals(extension))
                {
                    return "image/png";
                }
                if ("gif".Equals(extension))
                {
                    return "image/gif";
                }
                if ("bmp".Equals(extension))
                {
                    return "image/bmp";
                }
                if ("tiff".Equals(extension))
                {
                    return "image/tiff";
                }
                if ("wmf".Equals(extension))
                {
                    return "image/x-wmf";
                }
                if ("emf".Equals(extension))
                {
                    return "image/x-emf";
                }
                return "image/unknown";
            }
        }


        private String SuggestFileExtension(byte[] _dataStream, int pictureBytesStartOffSet)
        {
            if (MatchSignature(_dataStream, JPG, pictureBytesStartOffSet))
            {
                return "jpg";
            }
            else if (MatchSignature(_dataStream, PNG, pictureBytesStartOffSet))
            {
                return "png";
            }
            else if (MatchSignature(_dataStream, GIF, pictureBytesStartOffSet))
            {
                return "gif";
            }
            else if (MatchSignature(_dataStream, BMP, pictureBytesStartOffSet))
            {
                return "bmp";
            }
            else if (MatchSignature(_dataStream, TIFF, pictureBytesStartOffSet) ||
                     MatchSignature(_dataStream, TIFF1, pictureBytesStartOffSet))
            {
                return "tiff";
            }
            else
            {
                // Need to load the image content before we can try the following tests
                FillImageContent();

                if (MatchSignature(content, WMF1, 0) || MatchSignature(content, WMF2, 0))
                {
                    return "wmf";
                }
                else if (MatchSignature(content, EMF, 0))
                {
                    return "emf";
                }
            }
            // TODO: DIB, PICT
            return "";
        }

        private static bool MatchSignature(byte[] dataStream, byte[] signature, int pictureBytesOffset)
        {
            bool matched = pictureBytesOffset < dataStream.Length;
            for (int i = 0; (i + pictureBytesOffset) < dataStream.Length && i < signature.Length; i++)
            {
                if (dataStream[i + pictureBytesOffset] != signature[i])
                {
                    matched = false;
                    break;
                }
            }
            return matched;
        }

        //  public String GetFileName()
        //  {
        //    return fileName;
        //  }

        //  private static String extractFileName(int blockStartIndex, byte[] dataStream) {
        //        int fileNameStartOffset = blockStartIndex + 0x7C;
        //        int fileNameSizeOffset = blockStartIndex + FILENAME_SIZE_OFFSET;
        //        int fileNameSize = LittleEndian.GetShort(dataStream, fileNameSizeOffSet);
        //
        //        int fileNameIndex = fileNameStartOffSet;
        //        char[] fileNameChars = new char[(fileNameSize-1)/2];
        //        int charIndex = 0;
        //        while(charIndex<fileNameChars.Length) {
        //            short aChar = LittleEndian.GetShort(dataStream, fileNameIndex);
        //            fileNameChars[charIndex] = (char)aChar;
        //            charIndex++;
        //            fileNameIndex += 2;
        //        }
        //        String fileName = new String(fileNameChars);
        //        return fileName.Trim();
        //    }

        private void FillRawImageContent()
        {
            this.rawContent = new byte[size];
            Array.Copy(_dataStream, pictureBytesStartOffset, rawContent, 0, size);
        }

        private void FillImageContent()
          {
            byte[] rawContent = GetRawContent();

            // HACK: Detect compressed images.  In reality there should be some way to determine
            //       this from the first 32 bytes, but I can't see any similarity between all the
            //       samples I have obtained, nor any similarity in the data block contents.
            if (MatchSignature(rawContent, COMPRESSED1, 32) || MatchSignature(rawContent, COMPRESSED2, 32))
            {
              try
              {

                  ZlibStream gzip = new ZlibStream(new MemoryStream(rawContent, 33, rawContent.Length - 33), CompressionMode.Decompress);
                MemoryStream out1 = new MemoryStream();
                byte[] buf = new byte[4096];
                int readBytes;
                while ((readBytes = gzip.Read(buf,0,4096)) > 0)
                {
                    out1.Write(buf, 0, readBytes);
                }
                content = out1.ToArray();
              }
              catch (IOException)
              {
                // Problems Reading from the actual MemoryStream should never happen
                // so this will only ever be a ZipException.
                //log.log(POILogger.INFO, "Possibly corrupt compression or non-compressed data", e);
              }
            } else {
              // Raw data is not compressed.
              content = rawContent;
            }
          }

        private static int GetPictureBytesStartOffset(int dataBlockStartOffset, byte[] _dataStream, int dataBlockSize)
        {
            int realPicoffset = dataBlockStartOffset;
            int dataBlockEndOffset = dataBlockSize + dataBlockStartOffset;

            // Skip over the PICT block
            int PICTFBlockSize = LittleEndian.GetShort(_dataStream, dataBlockStartOffset + PICT_HEADER_OFFSET); // Should be 68 bytes

            // Now the PICTF1
            int PICTF1BlockOffset = PICTFBlockSize + PICT_HEADER_OFFSET;
            short MM_TYPE = LittleEndian.GetShort(_dataStream, dataBlockStartOffset + PICT_HEADER_OFFSET + 2);
            if (MM_TYPE == 0x66)
            {
                // Skip the stPicName
                int cchPicName = LittleEndian.GetUByte(_dataStream, PICTF1BlockOffset);
                PICTF1BlockOffset += 1 + cchPicName;
            }
            int PICTF1BlockSize = LittleEndian.GetShort(_dataStream, dataBlockStartOffset + PICTF1BlockOffset);

            int unknownHeaderOffset = (PICTF1BlockSize + PICTF1BlockOffset) < dataBlockEndOffset ? (PICTF1BlockSize + PICTF1BlockOffset) : PICTF1BlockOffset;
            realPicoffset += (unknownHeaderOffset + UNKNOWN_HEADER_SIZE);
            if (realPicoffset >= dataBlockEndOffset)
            {
                realPicoffset -= UNKNOWN_HEADER_SIZE;
            }
            return realPicoffset;
        }

        private void FillJPGWidthHeight()
        {
            /*
            http://www.codecomments.com/archive281-2004-3-158083.html

            Algorhitm proposed by Patrick TJ McPhee:

            read 2 bytes
            make sure they are 'ffd8'x
            repeatedly:
            read 2 bytes
            make sure the first one is 'ff'x
            if the second one is 'd9'x stop
            else if the second one is c0 or c2 (or possibly other values ...)
            skip 2 bytes
            read one byte into depth
            read two bytes into height
            read two bytes into width
            else
            read two bytes into length
            skip forward length-2 bytes

            Also used Ruby code snippet from: http://www.bigbold.com/snippets/posts/Show/805 for reference
            */
            int pointer = pictureBytesStartOffset + 2;
            int firstByte = _dataStream[pointer];
            int secondByte = _dataStream[pointer + 1];

            int endOfPicture = pictureBytesStartOffset + size;
            while (pointer < endOfPicture - 1)
            {
                do
                {
                    firstByte = _dataStream[pointer];
                    secondByte = _dataStream[pointer + 1];
                    pointer += 2;
                } while (!(firstByte == (byte)0xFF) && pointer < endOfPicture - 1);

                if (firstByte == ((byte)0xFF) && pointer < endOfPicture - 1)
                {
                    if (secondByte == (byte)0xD9 || secondByte == (byte)0xDA)
                    {
                        break;
                    }
                    else if ((secondByte & 0xF0) == 0xC0 && secondByte != (byte)0xC4 && secondByte != (byte)0xC8 && secondByte != (byte)0xCC)
                    {
                        pointer += 5;
                        this.height = GetBigEndianShort(_dataStream, pointer);
                        this.width = GetBigEndianShort(_dataStream, pointer + 2);
                        break;
                    }
                    else
                    {
                        pointer++;
                        pointer++;
                        int length = GetBigEndianShort(_dataStream, pointer);
                        pointer += length;
                    }
                }
                else
                {
                    pointer++;
                }
            }
        }

        private void FillPNGWidthHeight()
        {
            /*
             Used PNG file format description from http://www.wotsit.org/download.asp?f=png
            */
            int HEADER_START = pictureBytesStartOffset + PNG.Length + 4;
            if (MatchSignature(_dataStream, IHDR, HEADER_START))
            {
                int IHDR_CHUNK_WIDTH = HEADER_START + 4;
                this.width = GetBigEndianInt(_dataStream, IHDR_CHUNK_WIDTH);
                this.height = GetBigEndianInt(_dataStream, IHDR_CHUNK_WIDTH + 4);
            }
        }

        /**
         * returns pixel width of the picture or -1 if dimensions determining was failed
         */
        public int Width
        {
            get
            {
                if (width == -1)
                {
                    FillWidthHeight();
                }
                return width;
            }
        }

        /**
         * returns pixel height of the picture or -1 if dimensions determining was failed
         */
        public int Height
        {
            get
            {
                if (height == -1)
                {
                    FillWidthHeight();
                }
                return height;
            }
        }

        private static int GetBigEndianInt(byte[] data, int offset)
        {
            return (((data[offset] & 0xFF) << 24) + ((data[offset + 1] & 0xFF) << 16) + ((data[offset + 2] & 0xFF) << 8) + (data[offset + 3] & 0xFF));
        }

        private static int GetBigEndianShort(byte[] data, int offset)
        {
            return (((data[offset] & 0xFF) << 8) + (data[offset + 1] & 0xFF));
        }

    }

}