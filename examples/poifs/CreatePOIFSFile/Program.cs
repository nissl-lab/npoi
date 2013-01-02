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
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * NPOI HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/


using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

using NPOI.POIFS.FileSystem;


namespace CreatePOIFSFile
{
    class Program
    {
        static void Main(string[] args)
        {
            POIFSFileSystem fs = new POIFSFileSystem();

            //get the root directory
            DirectoryEntry dir = fs.Root;
            //create a document entry
            dir.CreateDocument("Foo", new MemoryStream(new byte[] {0x01,0x02,0x03 }));

            //create a folder
            dir.CreateDirectory("Hello");

            //create a POIFS file called Foo.poifs
            FileStream output = new FileStream("Foo.poifs", FileMode.OpenOrCreate);
            fs.WriteFileSystem(output);
            output.Close();
        }
    }
}
