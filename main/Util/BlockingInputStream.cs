
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
        

namespace NPOI.Util
{
    using System;




/**
 * Implementation of a BlockingInputStream to provide data to 
 * RawDataBlock that expects data in 512 byte chunks.  Useful to read
 * data from slow (ie, non FileInputStream) sources, for example when 
 * Reading an OLE2 Document over a network. 
 *
 * Possible extentions: add a timeout. Curently a call to Read(byte[]) on this
 *    class is blocking, so use at your own peril if your underlying stream blocks. 
 *
 * @author Jens Gerhard
 * @author aviks - documentation cleanups. 
 */
public class BlockingInputStream
      : InputStream
{
      protected InputStream is;

      public BlockingInputStream(InputStream is)
      {
          this.Is = is;
      }

      public int available()
        throws IOException
      {
          return is.Available();
      }

      public void close()
        throws IOException
      {
          is.Close();
      }

      public void mark(int ReadLimit)
      {
          is.Mark(readLimit);
      }

      public bool markSupported()
      {
          return is.MarkSupported();
      }

      public int Read()
        throws IOException
      {
          return is.Read();
      }
      
      /**
       * We had to revert to byte per byte Reading to keep
       * with slow network connections on one hand, without
       * missing the end-of-file. 
       * This is the only method that does its own thing in this class
       *    everything else is delegated to aggregated stream. 
       * THIS IS A BLOCKING BLOCK READ!!!
       */
      public int Read(byte[] bf)
        throws IOException
      {
          
          int i = 0;
          int b = 4611;
          while ( i < bf.Length )
          {
              b = is.Read();
              if ( b == -1 )
                  break;
              bf[i++] = (byte) b;
          }
          if ( i == 0 && b == -1 )
              return -1;
          return i;
      }

      public int Read(byte[] bf, int s, int l)
        throws IOException
      {
          return is.Read(bf, s, l);
      }

      public void reset()
        throws IOException
      {
          is.Reset();
      }

      public long skip(long n)
        throws IOException
      {
          return is.Skip(n);
      }
}


