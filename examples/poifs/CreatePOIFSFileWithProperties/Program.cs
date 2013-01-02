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
using System.IO;
using System.Text;

using NPOI.HPSF;
using NPOI.POIFS.FileSystem;

namespace CreatePOIFSFileWithProperties
{
    class Program
    {
        static void Main(string[] args)
        {
            POIFSFileSystem fs = new POIFSFileSystem();

            //get the root directory
            DirectoryEntry dir = fs.Root;
            
            //create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            //Write the stream data of the DocumentSummaryInformation entry to the root directory
            dsi.Write(dir, DocumentSummaryInformation.DEFAULT_STREAM_NAME);

            //create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            //Write the stream data of the SummaryInformation entry to the root directory
            si.Write(dir, SummaryInformation.DEFAULT_STREAM_NAME);

            //create a POIFS file called Foo.poifs
            FileStream output = new FileStream("Foo.poifs", FileMode.OpenOrCreate);
            fs.WriteFileSystem(output);
            output.Close();
        }
    }
}
