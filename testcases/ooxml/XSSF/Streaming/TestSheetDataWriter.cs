/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.XSSF.Streaming
{
    using NPOI.Util;
    using NPOI.XSSF.Streaming;
    using NUnit.Framework;using NUnit.Framework.Legacy;

    [TestFixture]
    public sealed class TestSheetDataWriter
    {

        String unicodeSurrogates = "\uD835\uDF4A\uD835\uDF4B\uD835\uDF4C\uD835\uDF4D\uD835\uDF4E"
               + "\uD835\uDF4F\uD835\uDF50\uD835\uDF51\uD835\uDF52\uD835\uDF53\uD835\uDF54\uD835"
               + "\uDF55\uD835\uDF56\uD835\uDF57\uD835\uDF58\uD835\uDF59\uD835\uDF5A\uD835\uDF5B"
               + "\uD835\uDF5C\uD835\uDF5D\uD835\uDF5E\uD835\uDF5F\uD835\uDF60\uD835\uDF61\uD835"
               + "\uDF62\uD835\uDF63\uD835\uDF64\uD835\uDF65\uD835\uDF66\uD835\uDF67\uD835\uDF68"
               + "\uD835\uDF69\uD835\uDF6A\uD835\uDF6B\uD835\uDF6C\uD835\uDF6D\uD835\uDF6E\uD835"
               + "\uDF6F\uD835\uDF70\uD835\uDF71\uD835\uDF72\uD835\uDF73\uD835\uDF74\uD835\uDF75"
               + "\uD835\uDF76\uD835\uDF77\uD835\uDF78\uD835\uDF79\uD835\uDF7A";

        [Test]
        public void TestReplaceWithQuestionMark()
        {
            for (int i = 0; i < unicodeSurrogates.Length; i++)
            {
                ClassicAssert.IsFalse(SheetDataWriter.ReplaceWithQuestionMark(unicodeSurrogates[i]));
            }
            ClassicAssert.IsTrue(SheetDataWriter.ReplaceWithQuestionMark('\uFFFE'));
            ClassicAssert.IsTrue(SheetDataWriter.ReplaceWithQuestionMark('\uFFFF'));
            ClassicAssert.IsTrue(SheetDataWriter.ReplaceWithQuestionMark('\u0000'));
            ClassicAssert.IsTrue(SheetDataWriter.ReplaceWithQuestionMark('\u000F'));
            ClassicAssert.IsTrue(SheetDataWriter.ReplaceWithQuestionMark('\u001F'));
        }

        [Test]
        public void TestWriteUnicodeSurrogates()
        {
            SheetDataWriter writer = new SheetDataWriter();
            try
            {
                writer.OutputQuotedString(unicodeSurrogates);
                writer.Close();
                FileInfo file = writer.TempFileInfo;
                FileInputStream is1 = new FileInputStream(file.Open(FileMode.OpenOrCreate));
                String text;
                try
                {
                    byte[] data = IOUtils.ToByteArray(is1);
                    int index = 0;
                    if (data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF)
                        index = 3;
                    text = Encoding.UTF8.GetString(data, index, data.Length - index);
                }
                finally
                {
                    is1.Close();
                }
                ClassicAssert.AreEqual(unicodeSurrogates, text);
            }
            finally
            {
                IOUtils.CloseQuietly(writer);
            }
        }

        [Test]
        public void TestWriteNewLines()
        {
            SheetDataWriter writer = new SheetDataWriter();
            try
            {
                writer.OutputQuotedString("\r\n");
                writer.Close();
                FileInfo file = writer.TempFileInfo;
                FileInputStream is1 = new FileInputStream(file.Open(FileMode.OpenOrCreate));
                String text;
                try
                {
                    byte[] data = IOUtils.ToByteArray(is1);
                    int index = 0;
                    if (data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF)
                        index = 3;
                    text = Encoding.UTF8.GetString(data, index, data.Length - index);
                }
                finally
                {
                    is1.Close();
                }
                ClassicAssert.AreEqual("&#xd;&#xa;", text);
            }
            finally
            {
                IOUtils.CloseQuietly(writer);
            }
        }
    }
}
