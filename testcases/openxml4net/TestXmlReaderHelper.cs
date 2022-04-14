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

using ICSharpCode.SharpZipLib.Zip.Compression.Streams;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXml4Net.Util;
using NPOI.Util;
using NUnit.Framework;
using System;
using System.IO;
using System.Xml;

namespace TestCases.OpenXml4Net.OPC
{
    [TestFixture]
    public class TestXmlReaderHelper
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(TestPackage));

        private string fileName = @"D:\Code\OpenSource\npoi\npoi\testcases\test-data\spreadsheet\ExcelBigData.xlsx";
        private string xmlfileName = @"D:\Code\OpenSource\npoi\npoi\testcases\test-data\spreadsheet\ExcelBigData_sheet2.xml";

        [Test]
        public void TestGetTextReaderXmlElement()
        {
            string filePath = xmlfileName.Substring(0, xmlfileName.LastIndexOf("."));

            var sourceStream = File.Open(xmlfileName, FileMode.Open);

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
        public void TestXmlTextReaderOFZipStream()
        {
            string filePath = fileName.Substring(0, fileName.LastIndexOf("."));

            var opcPackage = OPCPackage.Open(fileName);
            var list = opcPackage.GetParts();

            foreach (var part in list)
            {
                if (part.PartName.Name == "/xl/worksheets/sheet1.xml")
                {
                    var sourceStream = part.GetInputStream();

                    int elementCount = 0;

                    XmlTextReader reader = new XmlTextReader(sourceStream);
                    while (reader.Read())
                    {
                        switch (reader.NodeType)
                        {
                            case XmlNodeType.XmlDeclaration:

                                break;
                            case XmlNodeType.Whitespace:

                                break;
                            case XmlNodeType.Element:
                                if (reader.Depth == 2 && reader.Name == "row")
                                {
                                    elementCount++;
                                }
                                reader.MoveToElement();
                                break;
                            case XmlNodeType.Text:
                                break;
                            case XmlNodeType.EndElement:
                                break;
                        }
                    }

                    reader.Close();
                }
            }
        }


        [Test]
        public void TestXmlTextReaderOFVirtualZipStream()
        {
            string filePath = fileName.Substring(0, fileName.LastIndexOf("."));

            var opcPackage = OPCPackage.Open(fileName);
            var list = opcPackage.GetParts();

            foreach (var part in list)
            {
                if (part.PartName.Name == "/xl/worksheets/sheet1.xml")
                {
                    InflaterInputStream sourceStream = (InflaterInputStream)part.GetInputStream();

                    var virtualStream = new VirtualStream(sourceStream, 1872, 50000);

                }
            }
        }

        [Test]
        public void TestRemoveSheetData()
        {
            var opcPackage = OPCPackage.Open(fileName);
            var list = opcPackage.GetParts();

            foreach (var part in list)
            {
                if (part.PartName.Name == "/xl/worksheets/sheet1.xml")
                {
                    InflaterInputStream sourceStream = (InflaterInputStream)part.GetInputStream();

                    int rowCount = 0;
                    var ms = XmlReaderHelper.RemoveSheetData(sourceStream, out rowCount);

                    XmlDocument xmlDoc = new XmlDocument();
                    XmlHelper.LoadXmlSafe(xmlDoc, ms);

                    string filePath = fileName.Substring(0, fileName.LastIndexOf("."));
                    WriteStream(filePath, ms);
                }
            }

        }

        private void WriteStream(string filePath, Stream stream)
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
                if (stream is MemoryStream)
                {
                    BinaryWriter w = new BinaryWriter(fs);
                    w.Write(((MemoryStream)stream).ToArray());
                    fs.Close();
                    return;
                }
                byte[] srcBuf = StreamToBytes(stream);
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






