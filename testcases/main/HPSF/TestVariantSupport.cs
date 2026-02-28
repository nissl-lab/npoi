/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

using System;
using System.IO;
using NPOI.HPSF;
using NPOI.HPSF.Wellknown;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.FileSystem;
using NPOI.Util;
using NUnit.Framework;
using NUnit.Framework.Legacy;

namespace TestCases.HPSF
{
    [TestFixture]
    public class TestVariantSupport
    {
        [Test]
        public void Test52337()
        {

            // document summary stream   from test1-excel.doc attached to Bugzilla 52337
            string documentSummaryEnc =
            "H4sIAAAAAAAAAG2RsUvDQBjFXxsraiuNKDoI8ZwclIJOjhYCGpQitINbzXChgTQtyQ3+Hw52cHB0E"+
            "kdHRxfBpeAf4H/g5KK+M59Firn8eNx3d+++x31+AZVSGdOfrZTHz+Prxrp7eTWH7Z2PO5+1ylTtrA"+
            "SskBrXKOiROhnavWREZskNWSK3ZI3ckyp5IC55JMvkiaySF7JIXlF4v0tPbzOAR1XE18MwM32dGjW"+
            "IVJAanaVhoppRFMZZDjjSgyO9WT10Cq1vVX/uh/Txn3pucc7m6fTiXPEPldG5Qc0t2vEkXic2iZ5c"+
            "JDkd8VFS3pcMBzIvS7buaeB3j06C1nF7krFJPRdz62M4rM7/8f3NtyE+LQyQoY8QCfbQwAU1l/UF0"+
            "ubraA6DXWzC5x7gG6xzLtsAAgAA";
            byte[] bytes = POIFS.Storage.RawDataUtil.Decompress(documentSummaryEnc);

            PropertySet ps = PropertySetFactory.Create(new ByteArrayInputStream(bytes));
            DocumentSummaryInformation dsi = (DocumentSummaryInformation) ps;
            Section s = dsi.Sections[0];

            object hdrs =  s.GetProperty(PropertyIDMap.PID_HEADINGPAIR);
            ClassicAssert.IsNotNull(hdrs);
            ClassicAssert.AreEqual(typeof(byte[]), hdrs.GetType());

            // parse the value
            Vector v = new Vector((short)Variant.VT_VARIANT);
            LittleEndianByteArrayInputStream lei = new LittleEndianByteArrayInputStream((byte[])hdrs, 0);
            v.Read(lei);

            TypedPropertyValue[] items = v.Values;
            ClassicAssert.AreEqual(2, items.Length);

            object cp = items[0].Value;
            ClassicAssert.IsNotNull(cp);
            ClassicAssert.AreEqual(typeof(CodePageString), cp.GetType());
            object i = items[1].Value;
            ClassicAssert.IsNotNull(i);
            ClassicAssert.AreEqual(typeof(Int32), i.GetType());
            ClassicAssert.AreEqual(1, i);

        }

        [Test]
        public void NewNumberTypes()
        {

            ClipboardData cd = new ClipboardData();
            cd.Value = new byte[10];

            object[][] exp = [
                [ Variant.VT_CF, cd.ToByteArray() ],
                [ Variant.VT_BOOL, true ],
                [ Variant.VT_LPSTR, "codepagestring" ],
                [ Variant.VT_LPWSTR, "widestring" ],
                [ Variant.VT_I2, -1 ], // int, not short ... :(
                [ Variant.VT_UI2, 0xFFFF ],
                [ Variant.VT_I4, -1 ],
                [ Variant.VT_UI4, 0xFFFFFFFFL ],
                [ Variant.VT_I8, -1L ],
                [ Variant.VT_UI8, BigInteger.ValueOf(long.MaxValue).Add(BigInteger.TEN) ],
                [ Variant.VT_R4, -999.99f ],
                [ Variant.VT_R8, -999.99d ],
            ];

            HSSFWorkbook wb = new HSSFWorkbook();
            POIFSFileSystem poifs = new POIFSFileSystem();
            DocumentSummaryInformation dsi = PropertySetFactory.NewDocumentSummaryInformation();
            CustomProperties cpList = new CustomProperties();
            foreach(object[] o in exp)
            {
                int type = (int)o[0];
                Property p = new Property(PropertyIDMap.PID_MAX+type, type, o[1]);
                cpList.Put("testprop"+type, new CustomProperty(p, "testprop"+type));

            }
            dsi.CustomProperties = cpList;
            dsi.Write(poifs.Root, DocumentSummaryInformation.DEFAULT_STREAM_NAME);
            MemoryStream bos = new MemoryStream();
            poifs.WriteFileSystem(bos);
            poifs.Close();
            poifs = new POIFSFileSystem(new MemoryStream(bos.ToArray()));
            dsi = (DocumentSummaryInformation) PropertySetFactory.Create(poifs.Root, DocumentSummaryInformation.DEFAULT_STREAM_NAME);
            cpList = dsi.CustomProperties;
            int i=0;
            foreach(object[] o in exp)
            {
                object obj = cpList.Get("testprop"+o[0]);
                if(o[1] is byte[])
                {
                    POITestCase.AssertEquals((byte[]) o[1], (byte[]) obj, "property "+i);
                }
                else
                {
                    ClassicAssert.AreEqual(o[1], obj, "property "+i);
                }
                i++;
            }
            poifs.Close();
        }
    }
}