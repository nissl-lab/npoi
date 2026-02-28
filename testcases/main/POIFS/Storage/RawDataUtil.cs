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

using System;
using System.IO;
using ICSharpCode.SharpZipLib.GZip;
using NPOI.Util;

namespace TestCases.POIFS.Storage
{
    public class RawDataUtil
    {
        public static byte[] Decode(string[] hexDataLines)
        {
            MemoryStream baos = new MemoryStream(hexDataLines.Length * 32 + 32);

            for (int i = 0; i < hexDataLines.Length; i++)
            {
                byte[] lineData = HexRead.ReadFromString(hexDataLines[i]);
                baos.Write(lineData, 0, lineData.Length);
            }

            return baos.ToArray();
        }

        public static void DumpData(byte[] data)
        {
            int i = 0; 
            Console.WriteLine("String[] hexDataLines = {");
            Console.Write("\t\"");

            while (true)
            {
                char[] cc = HexDump.ByteToHex(data[i]);
                Console.Write(cc[2]);
                Console.Write(cc[3]);

                i++;

                if (i >= data.Length)
                    break;

                if (i % 32 == 0)
                {
                    Console.WriteLine("\",");
                    Console.Write("\t\"");
                }
                else
                {
                    Console.Write(" ");
                }
            }

            Console.WriteLine("\", ");
            Console.WriteLine(");");
        }

        public static void ConfirmEqual(byte[] expected, string[] hexDataLines)
        {
            MemoryStream ms = new MemoryStream(hexDataLines.Length * 32 + 32);

            for (int i = 0; i < hexDataLines.Length; i++)
            {
                byte[] lineData = HexRead.ReadFromString(hexDataLines[i]);
                ms.Write(lineData, 0, lineData.Length);
            }

            if (!Array.Equals(expected, ms.ToArray()))
            {
                throw new System.Exception("different");
            }
        }

        /**
         * Decompress previously gziped/base64ed data
         *
         * @param data the gziped/base64ed data
         * @return the raw bytes
         * @throws IOException if you copy and pasted the data wrong
         */
        public static byte[] Decompress(String data)
        {
            byte[] base64Bytes = Convert.FromBase64String(data);
            return IOUtils.ToByteArray(new GZipInputStream(new ByteArrayInputStream(base64Bytes)));
        }

        /**
         * Compress raw data for test runs - usually called while debugging :)
         *
         * @param data the raw data
         * @return the gziped/base64ed data as String
         * @throws IOException usually not ...
         */
        public static String Compress(byte[] data)
        {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            GZipOutputStream gz = new GZipOutputStream(bos);
            gz.Write(data, 0, data.Length);
            gz.Finish();
            return Convert.ToBase64String(bos.ToByteArray());
        }
    }
}
