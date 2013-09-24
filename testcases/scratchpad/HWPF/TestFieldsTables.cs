/*
 *  ====================================================================
 * Licensed to the Apache Software Foundation (ASF) under one or more
 * contributor license agreements.  See the NOTICE file distributed with
 * this work for Additional information regarding copyright ownership.
 * The ASF licenses this file to You under the Apache License, Version 2.0
 * (the "License"); you may not use this file except in compliance with
 * the License.  You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * ====================================================================
 */

using NPOI.HWPF.Model;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;
namespace TestCases.HWPF
{


    /**
     * Test case for the fields tables, this test is based on the test-fields.doc
     * file instead of the usual test.doc.
     * 
     * @author Cedric Bosdonnat <cbosdonnat@novell.com>
     * 
     */
    [TestFixture]
    public class TestFieldsTables : HWPFTestCase
    {

        private static String[] EXPECTED = {

            "[19, 43) - FLD - 0x13; 0x1f\n" + "[43, 54) - FLD - 0x14; 0xff\n"
                    + "[54, 59) - FLD - 0x15; 0x81\n",

            "[31, 59) - FLD - 0x13; 0x45\n" + "[59, 61) - FLD - 0x14; 0xff\n"
                    + "[61, 66) - FLD - 0x15; 0x80\n",

            "[23, 49) - FLD - 0x13; 0x11\n" + "[49, 64) - FLD - 0x14; 0xff\n"
                    + "[64, 69) - FLD - 0x15; 0x80\n",

            "[18, 42) - FLD - 0x13; 0x21\n" + "[42, 44) - FLD - 0x14; 0xff\n"
                    + "[44, 47) - FLD - 0x15; 0x81\n"
                    + "[47, 75) - FLD - 0x13; 0x1d\n"
                    + "[75, 85) - FLD - 0x14; 0xff\n"
                    + "[85, 91) - FLD - 0x15; 0x81\n",

            "[30, 54) - FLD - 0x13; 0x20\n" + "[54, 62) - FLD - 0x14; 0xff\n"
                    + "[62, 68) - FLD - 0x15; 0x81\n",

            "[1, 31) - FLD - 0x13; 0x15\n" + "[31, 51) - FLD - 0x14; 0xff\n"
                    + "[51, 541) - FLD - 0x15; 0x81\n",

            "[19, 47) - FLD - 0x13; 0x19\n" + "[47, 49) - FLD - 0x14; 0xff\n"
                    + "[49, 55) - FLD - 0x15; 0x81\n"

    };

        public TestFieldsTables()
        {
        }

        protected String GetTestFile()
        {
            return "test-fields.doc";
        }
        [Test]
        public void TestReadFields()
        {
            FileInformationBlock fib = _hWPFDocFixture._fib;
            byte[] tableStream = _hWPFDocFixture._tableStream;

            FieldsTables fieldsTables = new FieldsTables(tableStream, fib);

            Array values = Enum.GetValues(typeof(FieldsDocumentPart));
            for (int i = 0; i < values.Length; i++)
            {
                FieldsDocumentPart part = (FieldsDocumentPart)values.GetValue(i);

                List<PlexOfField> fieldsPlexes = fieldsTables
                        .GetFieldsPLCF(part);
                String result = DumpPlexes(fieldsPlexes);
                Assert.AreEqual(EXPECTED[i], result);
            }
        }

        private String DumpPlexes(List<PlexOfField> fieldsPlexes)
        {
            StringBuilder dump = new StringBuilder();
            for (int i = 0; i < fieldsPlexes.Count; i++)
            {
                PlexOfField flds = fieldsPlexes[i];
                dump.Append(flds.ToString() + "\n");
            }
            return dump.ToString();
        }
    }
}


