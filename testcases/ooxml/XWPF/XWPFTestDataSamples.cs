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
namespace NPOI.XWPF
{
    using System;
    using NPOI.XWPF.UserModel;
    using NUnit.Framework;
    using System.IO;
    using TestCases;
    using NPOI.Util;

    /**
     * @author Yegor Kozlov
     */
    public class XWPFTestDataSamples
    {

        public static XWPFDocument OpenSampleDocument(String sampleName)
        {
            Stream is1 = POIDataSamples.GetDocumentInstance().OpenResourceAsStream(sampleName);
            return new XWPFDocument(is1);
        }

        public static XWPFDocument WriteOutAndReadBack(XWPFDocument doc)
        {
            MemoryStream baos = new MemoryStream(4096);
            doc.Write(baos);
            MemoryStream bais = new MemoryStream(baos.ToArray());
            //WriteTo("D:\\testdoc.zip", baos.ToArray());
            return new XWPFDocument(bais);
        }

        public static byte[] GetImage(String filename)
        {
            Stream is1 = POIDataSamples.GetDocumentInstance().OpenResourceAsStream(filename);
            byte[] result = IOUtils.ToByteArray(is1);
            return result;
        }

        public static void WriteTo(string fileName, byte[] data)
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite))
            {
                fs.Write(data, 0, data.Length);
                fs.Flush();
            }
        }
    }

}