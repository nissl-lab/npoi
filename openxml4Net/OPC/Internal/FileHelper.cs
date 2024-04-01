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
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace NPOI.OpenXml4Net.OPC.Internal
{
    /// <summary>
    /// Provide useful method to manage file.
    /// </summary>
    /// <remarks>
    /// @author Julien Chable
    /// @version 0.1
    /// </remarks>
    public class FileHelper
    {

        /// <summary>
        /// Get the directory part of the specified file path.
        /// </summary>
        /// <param name="f">
        /// File to process.
        /// </param>
        /// <returns>The directory path from the specified</returns>
        public static string GetDirectory(string filepath)
        {
            return Path.GetDirectoryName(filepath).Replace("\\","/");
        }

        /// <summary>
        /// Copy a file.
        /// </summary>
        /// <param name="in">
        /// The source file.
        /// </param>
        /// <param name="out">
        /// The target location.
        /// </param>
        /// <exception cref="IOException">IOException
        /// If an I/O error occur.
        /// </exception>
        public static void CopyFile(string inpath, string outpath){
            File.Copy(inpath, outpath,true);
        }
        public static void CopyFile(FileInfo inpath, FileInfo outpath)
        {
            File.Copy(inpath.FullName, outpath.FullName, true);
        }
        /// <summary>
        /// Get file name from the specified File object.
        /// </summary>
        public static String GetFilename(string filepath)
        {
            String path = filepath;
            int len = path.Length;
            int num2 = len;
            while (--num2 >= 0)
            {
                char ch1 = path[num2];
                if (ch1 == '\\')
                    return path.Substring(num2 + 1, len);
            }
            return "";
        }

    }

}