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
using System.Collections;
using System.IO;

using NUnit.Framework;

using NPOI.POIFS.Common;
using NPOI.POIFS.Storage;
using NPOI.POIFS.Properties;
using TestCases.POIFS.Properties;
using System.Collections.Generic;

namespace TestCases.POIFS.Properties
{
    /**
     * Class to Test DirectoryProperty functionality
     *
     * @author Marc Johnson
     */
    [TestFixture]
    public class TestDirectoryProperty
    {
        private DirectoryProperty _property;
        private byte[] _testblock;

        /**
         * Constructor TestDirectoryProperty
         *
         * @param name
         */

        public TestDirectoryProperty()
        {

        }

        /**
         * Test constructing DirectoryProperty
         *
         * @exception IOException
         */
        [Test]
        public void TestConstructor()
        {
            CreateBasicDirectoryProperty();
            VerifyProperty();
        }

        /**
         * Test pre-Write functionality
         *
         * @exception IOException
         */
        [Test]
        public void TestPreWrite()
        {
            CreateBasicDirectoryProperty();
            _property.PreWrite();

            // shouldn't Change anything at all
            VerifyProperty();
            VerifyChildren(0);

            // now try Adding 1 property
            CreateBasicDirectoryProperty();
            _property.AddChild(new LocalProperty(1));
            _property.PreWrite();

            // update children index
            _testblock[0x4C] = 1;
            _testblock[0x4D] = 0;
            _testblock[0x4E] = 0;
            _testblock[0x4F] = 0;
            VerifyProperty();
            VerifyChildren(1);

            // now try Adding 2 properties
            CreateBasicDirectoryProperty();
            _property.AddChild(new LocalProperty(1));
            _property.AddChild(new LocalProperty(2));
            _property.PreWrite();

            // update children index
            _testblock[0x4C] = 2;
            _testblock[0x4D] = 0;
            _testblock[0x4E] = 0;
            _testblock[0x4F] = 0;
            VerifyProperty();
            VerifyChildren(2);

            // beat on the children allocation code
            for (int count = 1; count < 100; count++)
            {
                CreateBasicDirectoryProperty();
                for (int j = 1; j < (count + 1); j++)
                {
                    _property.AddChild(new LocalProperty(j));
                }
                _property.PreWrite();
                VerifyChildren(count);
            }
        }

        private void VerifyChildren(int count)
        {
            IEnumerator<Property> iter = _property.Children;
            List<Property> children = new List<Property>();

            while (iter.MoveNext())
            {
                children.Add(iter.Current);
            }
            Assert.AreEqual(count, children.Count);
            if (count != 0)
            {
                bool[] found = new bool[count];

                found[_property.ChildIndex - 1] = true;
                int total_found = 1;
                for (var i = 0; i < found.Length; i++)
                {
                    found[i] = false;
                }
                iter = children.GetEnumerator();
                while (iter.MoveNext())
                {
                    Property child = iter.Current;
                    Child next = child.NextChild;

                    if (next != null)
                    {
                        int index = ((Property)next).Index;

                        if (index != -1)
                        {
                            Assert.IsTrue(!found[index - 1], "found index " + index + " twice");
                            found[index - 1] = true;
                            total_found++;
                        }
                    }
                    Child previous = child.PreviousChild;

                    if (previous != null)
                    {
                        int index = ((Property)previous).Index;

                        if (index != -1)
                        {
                            Assert.IsTrue(!found[index - 1], "found index " + index + " twice");
                            found[index - 1] = true;
                            total_found++;
                        }
                    }
                }
                Assert.AreEqual(count, total_found);
            }
        }

        private void CreateBasicDirectoryProperty()
        {
            String name = "MyDirectory";

            _property = new DirectoryProperty(name);
            _testblock = new byte[128];
            int index = 0;

            for (; index < 0x40; index++)
            {
                _testblock[index] = (byte)0;
            }
            int limit = Math.Min(31, name.Length);

            _testblock[index++] = (byte)(2 * (limit + 1));
            _testblock[index++] = (byte)0;
            _testblock[index++] = (byte)1;
            _testblock[index++] = (byte)1;
            for (; index < 0x50; index++)
            {
                _testblock[index] = (byte)0xff;
            }
            for (; index < 0x80; index++)
            {
                _testblock[index] = (byte)0;
            }
            byte[] name_bytes = Encoding.GetEncoding(1252).GetBytes(name);

            for (index = 0; index < limit; index++)
            {
                _testblock[index * 2] = name_bytes[index];
            }
        }

        private void VerifyProperty()
        {
            MemoryStream stream = new MemoryStream(512);

            _property.WriteData(stream);
            byte[] output = stream.ToArray();

            Assert.AreEqual(_testblock.Length, output.Length);
            for (int j = 0; j < _testblock.Length; j++)
            {
                Assert.AreEqual(_testblock[j],
                             output[j], "mismatch at offset " + j);
            }
        }

        /**
         * Test AddChild
         *
         * @exception IOException
         */
        [Test]
        public void TestAddChild()
        {
            CreateBasicDirectoryProperty();
            _property.AddChild(new LocalProperty(1));
            _property.AddChild(new LocalProperty(2));
            try
            {
                _property.AddChild(new LocalProperty(1));
                Assert.Fail("should have caught IOException");
            }
            catch (IOException )
            {

                // as expected
            }
            try
            {
                _property.AddChild(new LocalProperty(2));
                Assert.Fail("should have caught IOException");
            }
            catch (IOException )
            {

                // as expected
            }
            _property.AddChild(new LocalProperty(3));
        }

        /**
         * Test DeleteChild
         *
         * @exception IOException
         */
        [Test]
        public void TestDeleteChild()
        {
            CreateBasicDirectoryProperty();
            Property p1 = new LocalProperty(1);

            _property.AddChild(p1);
            try
            {
                _property.AddChild(new LocalProperty(1));
                Assert.Fail("should have caught IOException");
            }
            catch (IOException )
            {

                // as expected
            }
            Assert.IsTrue(_property.DeleteChild(p1));
            Assert.IsTrue(!_property.DeleteChild(p1));
            _property.AddChild(new LocalProperty(1));
        }

        /**
         * Test ChangeName
         *
         * @exception IOException
         */
        [Test]
        public void TestChangeName()
        {
            CreateBasicDirectoryProperty();
            Property p1 = new LocalProperty(1);
            String originalName = p1.Name;

            _property.AddChild(p1);
            Assert.IsTrue(_property.ChangeName(p1, "foobar"));
            Assert.AreEqual("foobar", p1.Name);
            Assert.IsTrue(!_property.ChangeName(p1, "foobar"));
            Assert.AreEqual("foobar", p1.Name);
            Property p2 = new LocalProperty(1);

            _property.AddChild(p2);
            Assert.IsTrue(!_property.ChangeName(p1, originalName));
            Assert.IsTrue(_property.ChangeName(p2, "foo"));
            Assert.IsTrue(_property.ChangeName(p1, originalName));
        }

        /**
         * Test Reading constructor
         *
         * @exception IOException
         */
        [Test]
        public void TestReadingConstructor()
        {
            byte[] input =
        {
            ( byte ) 0x42, ( byte ) 0x00, ( byte ) 0x6F, ( byte ) 0x00,
            ( byte ) 0x6F, ( byte ) 0x00, ( byte ) 0x74, ( byte ) 0x00,
            ( byte ) 0x20, ( byte ) 0x00, ( byte ) 0x45, ( byte ) 0x00,
            ( byte ) 0x6E, ( byte ) 0x00, ( byte ) 0x74, ( byte ) 0x00,
            ( byte ) 0x72, ( byte ) 0x00, ( byte ) 0x79, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x16, ( byte ) 0x00, ( byte ) 0x01, ( byte ) 0x01,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0x02, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x20, ( byte ) 0x08, ( byte ) 0x02, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0xC0, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x46,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0xC0, ( byte ) 0x5C, ( byte ) 0xE8, ( byte ) 0x23,
            ( byte ) 0x9E, ( byte ) 0x6B, ( byte ) 0xC1, ( byte ) 0x01,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00
        };

            VerifyReadingProperty(0, input, 0, "Boot Entry");
        }

        private void VerifyReadingProperty(int index, byte[] input, int offset,
                                           String name)
        {
            DirectoryProperty property = new DirectoryProperty(index, input,
                                                 offset);
            MemoryStream stream = new MemoryStream(128);
            byte[] expected = new byte[128];

            Array.Copy(input, offset, expected, 0, 128);
            property.WriteData(stream);
            byte[] output = stream.ToArray();

            Assert.AreEqual(128, output.Length);
            for (int j = 0; j < 128; j++)
            {
                Assert.AreEqual(expected[j],
                             output[j], "mismatch at offset " + j);
            }
            Assert.AreEqual(index, property.Index);
            Assert.AreEqual(name, property.Name);
            Assert.IsTrue(!property.Children.MoveNext());
        }

    }
}