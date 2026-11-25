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
using NPOI.OpenXmlFormats.Spreadsheet;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using System.IO;
using System.Xml;
using TestCases;

namespace TestCases.XSSF.UserModel
{
    /// <summary>
    /// Verifies reading and writing of the threaded comments part
    /// (/xl/threadedComments/threadedComment#.xml) and the persons part
    /// (/xl/persons/person.xml) of an xlsx file.
    /// </summary>
    [TestFixture]
    public class TestXSSFThreadedComments
    {
        private const string THREADED_COMMENTS_NS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";

        [Test]
        public void ReadThreadedComments()
        {
            using (OPCPackage pkg = OpenSamplePackage())
            {
                PackagePart tcPart = GetPart(pkg, "/xl/threadedComments/threadedComment1.xml");
                ClassicAssert.IsNotNull(tcPart);
                ClassicAssert.AreEqual("application/vnd.ms-excel.threadedcomments+xml", tcPart.ContentType);

                CT_ThreadedComments tcs = ParseThreadedComments(tcPart.GetInputStream());
                ClassicAssert.AreEqual(2, tcs.SizeOfThreadedCommentArray());

                CT_ThreadedComment first = tcs.GetThreadedCommentArray(0);
                ClassicAssert.AreEqual("B5", first.@ref);
                ClassicAssert.AreEqual(new DateTime(2026, 3, 19, 4, 24, 56, 770), first.dT);
                ClassicAssert.AreEqual("{EE9C261F-D288-4FD8-911C-44FB2C2DDA41}", first.personId);
                ClassicAssert.AreEqual("{F3905F37-1CD1-4482-94B3-2EB134F45AAA}", first.id);
                ClassicAssert.IsNull(first.parentId);
                ClassicAssert.AreEqual("Here’s a threaded comment", first.text);

                CT_ThreadedComment reply = tcs.GetThreadedCommentArray(1);
                ClassicAssert.AreEqual("{F3905F37-1CD1-4482-94B3-2EB134F45AAA}", reply.parentId);
                ClassicAssert.AreEqual("And a reply", reply.text);

                PackagePart personPart = GetPart(pkg, "/xl/persons/person.xml");
                ClassicAssert.IsNotNull(personPart);
                ClassicAssert.AreEqual("application/vnd.ms-excel.person+xml", personPart.ContentType);

                CT_PersonList persons = ParsePersons(personPart.GetInputStream());
                ClassicAssert.AreEqual(1, persons.SizeOfPersonArray());
                CT_Person person = persons.GetPersonArray(0);
                ClassicAssert.AreEqual("Dean MacGregor", person.displayName);
                ClassicAssert.AreEqual("{EE9C261F-D288-4FD8-911C-44FB2C2DDA41}", person.id);
                ClassicAssert.AreEqual("S::Dean.MacGregor@baywa-re.com::427304bd-5413-4334-abdf-45e809d9fcbc", person.userId);
                ClassicAssert.AreEqual("AD", person.providerId);
            }
        }

        [Test]
        public void RoundTripThreadedComments()
        {
            using (OPCPackage pkg = OpenSamplePackage())
            {
                PackagePart tcPart = GetPart(pkg, "/xl/threadedComments/threadedComment1.xml");
                CT_ThreadedComments tcs = ParseThreadedComments(tcPart.GetInputStream());

                // write out, then read back in
                MemoryStream ms = new MemoryStream();
                ThreadedCommentsDocument doc = new ThreadedCommentsDocument(tcs);
                doc.Save(ms);
                ms.Flush();

                CT_ThreadedComments reparsed = ParseThreadedComments(new MemoryStream(ms.ToArray()));
                ClassicAssert.AreEqual(2, reparsed.SizeOfThreadedCommentArray());
                for (int i = 0; i < tcs.SizeOfThreadedCommentArray(); i++)
                {
                    CT_ThreadedComment expected = tcs.GetThreadedCommentArray(i);
                    CT_ThreadedComment actual = reparsed.GetThreadedCommentArray(i);
                    ClassicAssert.AreEqual(expected.@ref, actual.@ref);
                    ClassicAssert.AreEqual(expected.dT, actual.dT);
                    ClassicAssert.AreEqual(expected.personId, actual.personId);
                    ClassicAssert.AreEqual(expected.id, actual.id);
                    ClassicAssert.AreEqual(expected.parentId, actual.parentId);
                    ClassicAssert.AreEqual(expected.text, actual.text);
                }

                MemoryStream ms2 = new MemoryStream();
                PersonListDocument pDoc = new PersonListDocument(ParsePersons(GetPart(pkg, "/xl/persons/person.xml").GetInputStream()));
                pDoc.Save(ms2);
                ms2.Flush();
                CT_PersonList persons = ParsePersons(new MemoryStream(ms2.ToArray()));
                ClassicAssert.AreEqual(1, persons.SizeOfPersonArray());
                ClassicAssert.AreEqual("Dean MacGregor", persons.GetPersonArray(0).displayName);
            }
        }

        private static OPCPackage OpenSamplePackage()
        {
            Stream input = POIDataSamples.GetSpreadSheetInstance().OpenResourceAsStream("threaded_example.xlsx");
            return OPCPackage.Open(input);
        }

        private static PackagePart GetPart(OPCPackage pkg, string name)
        {
            foreach (PackagePart part in pkg.GetParts())
            {
                if (part.PartName.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                    return part;
            }
            return null;
        }

        private static CT_ThreadedComments ParseThreadedComments(Stream input)
        {
            XmlDocument xmlDoc = new XmlDocument();
            NPOI.OpenXml4Net.Util.XmlHelper.LoadXmlSafe(xmlDoc, input);
            XmlNamespaceManager ns = new XmlNamespaceManager(xmlDoc.NameTable);
            ns.AddNamespace(string.Empty, THREADED_COMMENTS_NS);
            return ThreadedCommentsDocument.Parse(xmlDoc, ns).GetThreadedComments();
        }

        private static CT_PersonList ParsePersons(Stream input)
        {
            XmlDocument xmlDoc = new XmlDocument();
            NPOI.OpenXml4Net.Util.XmlHelper.LoadXmlSafe(xmlDoc, input);
            XmlNamespaceManager ns = new XmlNamespaceManager(xmlDoc.NameTable);
            ns.AddNamespace(string.Empty, THREADED_COMMENTS_NS);
            return PersonListDocument.Parse(xmlDoc, ns).GetPersonList();
        }
    }
}
