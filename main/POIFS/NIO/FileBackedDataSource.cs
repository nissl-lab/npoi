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
//using System.IO.MemoryMappedFiles;
namespace NPOI.POIFS.NIO
{
    /// <summary>
    /// A POIFS DataSource backed by a File
    /// TODO - Return the ByteBuffers in such a way that in RW mode,
    /// changes to the buffer end up on the disk (will fix the HPSF TestWrite
    /// currently failing unit test when done)
    /// </summary>
    public class FileBackedDataSource : DataSource
    {
        private MemoryStream fileStream;
        //private MemoryMappedFile mmFile;
        //private MemoryMappedViewStream mmViewStream;
        private FileInfo fileinfo;
        private bool writable;
        public FileBackedDataSource(FileInfo file)
            : this(file, false)
        {
            
        }
        public FileBackedDataSource(FileInfo file, bool readOnly)
        {
            if (!file.Exists)
                throw new FileNotFoundException(file.FullName);
            this.fileinfo = file;
            FileStream stream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read);
            byte[] temp = new byte[stream.Length];
            stream.Read(temp, 0, (int)stream.Length);
            MemoryStream ms = new MemoryStream(temp, 0, temp.Length);
            fileStream = ms;
            this.writable = !readOnly;
            stream.Position = 0;
        }
        public FileBackedDataSource(FileStream stream, bool readOnly)
        {
            stream.Position = 0;
            byte[] temp = new byte[stream.Length];
            stream.Read(temp, 0, (int)stream.Length);
            MemoryStream ms = new MemoryStream(temp, 0, temp.Length);
            fileStream = ms;
            this.writable = !readOnly;
            stream.Position = 0;
        }

        #region IDisposable Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (null != fileStream)
                {
                    fileStream.Dispose();
                    fileStream = null;
                }
            }
        }

        ~FileBackedDataSource()
        {
            Dispose(false);
        }

        #endregion

        public bool IsWriteable
        {
            get
            {
                return this.writable;
            }
        }

        public Stream Stream
        {
            get { return this.fileStream; }
        }

        /// <summary>
        /// Reads a sequence of bytes from this FileStream starting at the given file position.
        /// </summary>
        /// <param name="length"></param>
        /// <param name="position">The file position at which the transfer is to begin;</param>
        /// <returns></returns>
        public override ByteBuffer Read(int length, long position)
        {
            if (position >= Size)
                throw new ArgumentException("Position " + position + " past the end of the file");

            // Do we read or map (for read/write?
            ByteBuffer dst;
            int worked = -1;
            if (writable)
            {
                //dst = channel.map(FileChannel.MapMode.READ_WRITE, position, length);
                dst = ByteBuffer.CreateBuffer(length);
                worked = 0;
            }
            else
            {
                // Read
                fileStream.Position = position;
                dst = ByteBuffer.CreateBuffer(length);

                worked = IOUtils.ReadFully(fileStream, dst.Buffer);
            }
            // Check
            if(worked == -1)
                throw new ArgumentException("Position " + position + " past the end of the file");

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
        public override void  Write(ByteBuffer src, long position)
        {
            fileStream.Write(src.Buffer, (int)position, src.Length);
        }

        public override void CopyTo(Stream stream)
        {
            //byte[] tempBuffer = new byte[stream.Length];
            //fileStream.Read(tempBuffer, 0, tempBuffer.Length);
            byte[] tempBuffer = fileStream.ToArray();
            stream.Write(tempBuffer, 0, tempBuffer.Length);        }

        public override long Size
        {
            get 
            {
                if (fileStream != null)
                {
                    return fileStream.Length;
                }
                else
                {
                    return fileinfo.Length;
                }
            }
        }

        public override void Close()
        {
            if (fileStream != null)
            {
                fileStream.Close();
            }
            
        }
    }

}