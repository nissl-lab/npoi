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

using System;
using System.Collections;
using System.IO;

using NPOI.POIFS.FileSystem;

namespace NPOI.POIFS.Dev
{
    public class POIFSLister
    {
        public static void ViewFile(String filename)
        {
            using (Stream stream = new FileStream(filename, FileMode.Open))
            {
                POIFSFileSystem fs = new POIFSFileSystem(stream);
                DisplayDirectory(fs.Root, "");
            }
        }

        public static void DisplayDirectory(DirectoryNode dir, String indent)
        {
            Console.WriteLine(indent + dir.Name + " -");
            String newIndent = indent + "  ";

            IEnumerator it = dir.Entries;
            while (it.MoveNext())
            {
                Object entry = it.Current;
                if (entry is DirectoryNode)
                {
                    DisplayDirectory((DirectoryNode)entry, newIndent);
                }
                else
                {
                    DocumentNode doc = (DocumentNode)entry;
                    String name = doc.Name;
                    if (name[0] < 10)
                    {
                        String altname = "(0x0" + (int)name[0] + ")" + name.Substring(1);
                        name = name.Substring(1) + " <" + altname + ">";
                    }
                    Console.WriteLine(newIndent + name);
                }
            }
        }
    }
}