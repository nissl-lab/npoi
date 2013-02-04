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

using System.IO;
using NPOI.POIFS.Common;
namespace NPOI.POIFS.Storage
{
    /// <summary>
    /// Abstract base class of all POIFS block storage classes. All
    /// extensions of BigBlock should write 512 bytes of data when
    /// requested to write their data.
    /// This class has package scope, as there is no reason at this time to
    /// make the class public.
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public abstract class BigBlock : BlockWritable
    {

        protected POIFSBigBlockSize bigBlockSize;

        protected BigBlock()
        {
        }

        protected BigBlock(POIFSBigBlockSize bigBlockSize)
        {
            this.bigBlockSize = bigBlockSize;
        }
        /// <summary>
        /// Default implementation of write for extending classes that
        /// contain their data in a simple array of bytes.
        /// </summary>
        /// <param name="stream">the OutputStream to which the data should be written.</param>
        /// <param name="data">the byte array of to be written.</param>
        protected void WriteData(Stream stream, byte[] data)
        {
            stream.Write(data, 0, data.Length);
        }
        /// <summary>
        /// Write the block's data to an OutputStream
        /// </summary>
        /// <param name="stream">the OutputStream to which the stored data should be written</param>
        public void WriteBlocks(Stream stream)
        {
            this.WriteData(stream);
        }

        /// <summary>
        /// Write the storage to an OutputStream
        /// </summary>
        /// <param name="stream">the OutputStream to which the stored data should be written </param>
        public abstract void WriteData(Stream stream);

    }


}
