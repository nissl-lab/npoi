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

using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXml4Net.OPC.Internal;
using System.IO;
using System.Collections.Generic;
using System;
using TestCases.OpenXml4Net;
using NPOI.Util;
using System.Reflection;
using System.Text.RegularExpressions;
using NUnit.Framework;
using System.Xml;
using System.Text;
using ICSharpCode.SharpZipLib.Zip;
using System.Collections;
using NPOI.SS.UserModel;
using NPOI;
using NPOI.Openxml4Net.Exceptions;
using NPOI.OpenXml4Net.Util;

namespace TestCases.OpenXml4Net.OPC
{
    [TestFixture]
    public class TestXmlReaderHelper
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(TestPackage));

        /**
         * Test that just opening and closing the file doesn't alter the document.
         */
        [Test]
        public void TestOpenSave()
        {
            //var filePath = @"D:\Documents\C04Tencent\02项目建设\20220331管家2022H1性能优化\资料\sheet2.xml";
            var filePath = @"D:\Documents\C04Tencent\02项目建设\20220331管家2022H1性能优化\资料\test.txt";
            try
            {
                XmlReaderHelper.SplitXml(filePath, 10);
            }
            catch (Exception ex)
            {
            }
        }

        [Test]
        public void TestCountPosition()
        {
            var fileName = @"D:\Documents\C04Tencent\02项目建设\20220331管家2022H1性能优化\资料\sheet2-XmlDocument.xml";
            string filePath = fileName.Substring(0, fileName.LastIndexOf("."));
            try
            {
                var rst = XmlReaderHelper.CountPosition(fileName, 1, "sheetData");
            }
            catch (Exception ex)
            {
            }
        }

        [Test]
        public void TestRemoveXmlElement()
        {
            var fileName = @"D:\Documents\C04Tencent\02项目建设\20220331管家2022H1性能优化\资料\sheet2-XmlDocument.xml";
            string filePath = fileName.Substring(0, fileName.LastIndexOf("."));
            var outfile = filePath + "_structure.xml";

            if (File.Exists(outfile))
            {
                try
                {
                    File.Delete(outfile);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            try
            {
                XmlReaderHelper.RemoveXmlElement(fileName, outfile, 1, "sheetData");
            }
            catch (Exception ex)
            {
            }
        }


        [Test]
        public void TestGetTextReaderXmlElement()
        {
            var fileName = @"D:\Documents\C04Tencent\02项目建设\20220331管家2022H1性能优化\资料\sheet2-XmlDocument.xml";
            //var fileName = @"D:\Documents\C04Tencent\02项目建设\20220331管家2022H1性能优化\资料\test.xml";
            string filePath = fileName.Substring(0, fileName.LastIndexOf("."));

            var sourceStream = File.Open(fileName, FileMode.Open);

            var strbytes = StreamToBytes(sourceStream);
            var str = ConvertBytes(strbytes);
            var position1 = str.IndexOf("<sheetData>");
            var position2 = str.IndexOf("</sheetData>") + "</sheetData>".Length;

            var virtualStream = new VirtualStream(sourceStream, position1, position2);

            WriteStream(filePath, virtualStream);

            virtualStream.Close();
            sourceStream.Close();
        }


        [Test]
        public void TestZip()
        {
            var fileName = @"D:\Documents\C04Tencent\02项目建设\20220331管家2022H1性能优化\资料\导入关键词30w物料模板.xlsx";
            string filePath = fileName.Substring(0, fileName.LastIndexOf("."));

            var opcPackage = OPCPackage.Open(fileName);
            var list = opcPackage.GetParts();

            foreach (var part in list)
            {
                if (part.PartName.Name == "/xl/worksheets/sheet1.xml")
                {
                    var sourceStream = part.GetInputStream();

                    var strbytes = StreamToBytes(sourceStream);
                    var str = ConvertBytes(strbytes);
                    var position1 = str.IndexOf("<sheetData>");
                    var position2 = str.IndexOf("</sheetData>") + "</sheetData>".Length;

                    var virtualStream = new VirtualStream(sourceStream, position1, position2);

                    WriteStream(filePath, virtualStream);

                    virtualStream.Close();
                    sourceStream.Close();
                }
            }
        }

        private void WriteStream(string filePath, VirtualStream virtualStream)
        {
            var outfile = filePath + "_data.xml";
            if (File.Exists(outfile))
            {
                try
                {
                    File.Delete(outfile);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            using (var fs = new FileStream(outfile, FileMode.OpenOrCreate))
            {
                byte[] srcBuf = StreamToBytes(virtualStream);
                fs.Write(srcBuf, 0, srcBuf.Length);
                fs.Flush();
            }
        }


        //Stream转Byte
        private byte[] StreamToBytes(Stream stream)
        {
            byte[] bytes = new byte[stream.Length];
            stream.Read(bytes, 0, bytes.Length);

            // 设置当前流的位置为流的开始 
            stream.Seek(0, SeekOrigin.Begin);
            return bytes;
        }

        private string ConvertBytes(byte[] bytes)
        {
            string strTemp = System.BitConverter.ToString(bytes);
            string[] strSplit = strTemp.Split('-');
            byte[] bytTemp2 = new byte[strSplit.Length];
            for (int i = 0; i < strSplit.Length; i++)
            {
                bytTemp2[i] = byte.Parse(strSplit[i], System.Globalization.NumberStyles.AllowHexSpecifier);
            }
            string strResult = System.Text.Encoding.Default.GetString(bytTemp2);
            return strResult;
        }
    }
}






