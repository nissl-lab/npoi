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
using System.Text;
using System.IO;
using System.Collections;


namespace NPOI.POIFS.Dev
{
    using NPOI.POIFS.FileSystem;

    public class POIFSViewer
    {
        public static void ViewFile(String filename, bool printName)
        {
            if (printName)
            {
                StringBuilder flowerbox = new StringBuilder();

                flowerbox.Append(".");
                for (int j = 0; j < filename.Length; j++)
                {
                    flowerbox.Append("-");
                }
                flowerbox.Append(".");
                Console.WriteLine(flowerbox);
                Console.WriteLine("|" + filename + "|");
                Console.WriteLine(flowerbox);
            }
            try
            {
                using (Stream fileStream = File.OpenRead(filename))
                {
                    POIFSViewable fs = (POIFSViewable)new POIFSFileSystem(fileStream);
                
                    IList strings = POIFSViewEngine.InspectViewable(fs, true,
                                                0, "  ");
                    IEnumerator iter = strings.GetEnumerator();

                    while (iter.MoveNext())
                    {
                        Console.Write(iter.Current);
                    }
                }
            }
            catch (IOException e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
