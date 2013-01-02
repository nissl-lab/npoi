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

namespace NPOI.Util
{
    using System;
    using System.IO;

    /// <summary>
    /// Wraps an <see cref="T:System.IO.Stream"/> providing <see cref="T:NPOI.Util.ILittleEndianOutput"/>
    /// </summary>
    /// <remarks>@author Josh Micich</remarks>
    public class LittleEndianOutputStream : ILittleEndianOutput, IDisposable
    {
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (null != out1)
                {
                    out1.Dispose();
                    out1 = null;
                }
            }
        }

        Stream out1 = null;

        public LittleEndianOutputStream(Stream out1)
        {
            this.out1 = out1;
        }

        public void WriteByte(int v)
        {
            try
            {
                out1.WriteByte((byte)v);
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }
        }

        public void WriteDouble(double v)
        {
            WriteLong(BitConverter.DoubleToInt64Bits(v));
        }

        public void WriteInt(int v)
        {
            int b3 = (v >> 24) & 0xFF;
            int b2 = (v >> 16) & 0xFF;
            int b1 = (v >> 8) & 0xFF;
            int b0 = (v >> 0) & 0xFF;
            try
            {
                out1.WriteByte((byte)b0);
                out1.WriteByte((byte)b1);
                out1.WriteByte((byte)b2);
                out1.WriteByte((byte)b3);
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }
        }

        public void WriteLong(long v)
        {
            WriteInt((int)(v >> 0));
            WriteInt((int)(v >> 32));
        }

        public void WriteShort(int v)
        {
            int b1 = (v >> 8) & 0xFF;
            int b0 = (v >> 0) & 0xFF;
            try
            {
                out1.WriteByte((byte)b0);
                out1.WriteByte((byte)b1);
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }
        }
        public void Write(byte[] b)
        {
            // suppress IOException for interface method

            try
            {
                out1.Write(b, 0, b.Length);
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }
        }
        public void Write(byte[] b, int off, int len)
        {
            // suppress IOException for interface method
            try
            {
                out1.Write(b, off, len);
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }
        }

        public void Flush()
        {
            out1.Flush();
        }
    }
}