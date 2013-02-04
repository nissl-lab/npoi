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


namespace NPOI.Util
{

    /// <summary>
    /// behavior of a field at a fixed location within a byte array
    /// @author Marc Johnson (mjohnson at apache dot org
    /// </summary>
    public interface FixedField
    {
        /// <summary>
        /// set the value from its offset into an array of bytes
        /// </summary>
        /// <param name="data">the byte array from which the value is to be read</param>
        void ReadFromBytes(byte[] data);
        /// <summary>
        /// set the value from an Stream
        /// </summary>
        /// <param name="stream">the Stream from which the value is to be read</param>
        void ReadFromStream(Stream stream);
        /// <summary>
        /// return the value as a String
        /// </summary>
        /// <returns></returns>
        string ToString();
        /// <summary>
        /// write the value out to an array of bytes at the appropriate offset
        /// </summary>
        /// <param name="data">the array of bytes to which the value is to be written</param>
        void WriteToBytes(byte[] data);
    }
}
