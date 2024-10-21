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
using System.Collections.Generic;
using System.Security;
using System.Reflection;
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

        // Buffers which map to a file-portion are not closed automatically when the Channel is closed
        // therefore we need to keep the list of mapped buffers and do some ugly reflection to try to 
        // clean the buffer during close().
        // See https://bz.apache.org/bugzilla/show_bug.cgi?id=58480, 
        // http://stackoverflow.com/questions/3602783/file-access-synchronized-on-java-object and
        // http://bugs.java.com/view_bug.do?bug_id=4724038 for related discussions
        private List<ByteBuffer> buffersToClean = new List<ByteBuffer>();

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
                throw new IndexOutOfRangeException("Position " + position + " past the end of the file");

            // Do we read or map (for read/write)?
            ByteBuffer dst;
            if (writable)
            {
                //dst = channel.map(FileChannel.MapMode.READ_WRITE, position, length);
                dst = ByteBuffer.CreateBuffer(length);
                // remember this buffer for cleanup
                buffersToClean.Add(dst);
            }
            else
            {
                // allocate the buffer on the heap if we cannot map the data in directly
                fileStream.Position = position;
                dst = ByteBuffer.CreateBuffer(length);

                // Read the contents and check that we could read some data
                int worked = IOUtils.ReadFully(fileStream, dst.Buffer);
                // Check
                if (worked == -1)
                    throw new IndexOutOfRangeException("Position " + position + " past the end of the file");
            }
            // make it ready for reading
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
            // also ensure that all buffers are unmapped so we do not keep files locked on Windows
            // We consider it a bug if a Buffer is still in use now! 
            foreach (ByteBuffer buffer in buffersToClean)
            {
                unmap(buffer);
            }
            buffersToClean.Clear();

            if (fileStream != null)
            {
                fileStream.Close();
            }
            
        }

        // need to use reflection to avoid depending on the sun.nio internal API
        // unfortunately this might break silently with newer/other Java implementations, 
        // but we at least have unit-tests which will indicate this when run on Windows
        private static void unmap(ByteBuffer bb)
        {
            //TODO: try add clean method for ByteBuffer class.
            //Type fcClass = bb.GetType();
            //try
            //{
            //    // invoke bb.cleaner().clean(), but do not depend on sun.nio
            //    // interfaces
            //    MethodInfo cleanerMethod = fcClass.GetMethod("cleaner");
            //    //cleanerMethod.setAccessible(true);
            //    Object cleaner = cleanerMethod.Invoke(bb, null);
            //    MethodInfo cleanMethod = cleaner.GetType().GetMethod("clean");
            //    cleanMethod.Invoke(cleaner, null);
            //}
            //catch (NotSupportedException e)
            //{
            //    // e.printStackTrace();
            //}
            //catch (SecurityException e)
            //{
            //    // e.printStackTrace();
            //}
            //catch (MethodAccessException e)
            //{
            //    // e.printStackTrace();
            //}
            //catch (ArgumentException e)
            //{
            //    // e.printStackTrace();
            //}
            //catch (TargetException e)
            //{
            //    // e.printStackTrace();
            //}
        }
    }

}