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

using System.IO;
using System;
using NPOI.Util;
namespace NPOI.POIFS.NIO
{
    /**
 * A POIFS {@link DataSource} backed by a File
 */
    public class FileBackedDataSource : DataSource
    {
        private FileStream fileStream;

        public FileBackedDataSource(string file)
        {
            if (!File.Exists(file))
            {
                throw new FileNotFoundException(file);
            }
            this.fileStream = new FileStream(file, FileMode.Open);
        }
        public FileBackedDataSource(FileStream channel)
        {
            this.fileStream = channel;
        }
        /// <summary>
        /// Reads a sequence of bytes from this FileStream starting at the given file position.
        /// </summary>
        /// <param name="length"></param>
        /// <param name="position">The file position at which the transfer is to begin;</param>
        /// <returns></returns>
        public override MemoryStream Read(int length, long position)
        {
            if (position >= Size)
            {
                throw new ArgumentException("Position " + position + " past the end of the file");
            }

            // Read
            fileStream.Position = position;
            byte[] buffer = new byte[length];
            int worked = IOUtils.ReadFully(fileStream, buffer);

            // Check
            if (worked == -1)
            {
                throw new ArgumentException("Position " + position + " past the end of the file");
            }
            MemoryStream dst = new MemoryStream(buffer);
            // Ready it for reading
            dst.Position = 0;

            // All done
            return dst;
        }
        /// <summary>
        /// Writes a sequence of bytes to this FileStream from the given Stream,
        /// starting at the given file position.
        /// </summary>
        /// <param name="src">The Stream from which bytes are to be transferred</param>
        /// <param name="position">The file position at which the transfer is to begin;
        /// must be non-negative</param>
        public override void Write(MemoryStream src, long position)
        {
            src.Seek(0, SeekOrigin.Begin);
            byte[] buffer = src.GetBuffer();
            fileStream.Position = position;
            fileStream.Write(buffer, 0, buffer.Length);
        }

        public override void CopyTo(Stream stream)
        {
            // Wrap the OutputSteam as a channel
            //WritableByteChannel out = Channels.newChannel(stream);
            // Now do the transfer
            //channel.transferTo(0, channel.size(), out);
            fileStream.Position = 0;
            byte[] buffer = new byte[fileStream.Length];
            fileStream.Read(buffer, 0, buffer.Length);
            stream.Write(buffer, 0, buffer.Length);
        }

        public override long Size
        {
            get
            {
                return fileStream.Length;
            }
        }

        public override void Close()
        {
            fileStream.Close();
        }
    }

}