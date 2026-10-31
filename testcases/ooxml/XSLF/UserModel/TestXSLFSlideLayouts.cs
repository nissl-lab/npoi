using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using NPOI;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Presentation;
using NPOI.XSLF.UserModel;
using NUnit.Framework;

namespace TestCases.XSLF.UserModel
{
    [TestFixture]
    public class TestXSLFSlideLayouts
    {
        [Test]
        public void MasterAndLayoutExtensionsRetainOrderAfterReadingAndEditing()
        {
            using var original = new XMLSlideShow();
            var masterName = original.GetSlideMasters()[0].GetPackagePart().PartName;
            var layoutName = original.GetSlideMasters()[0].GetSlideLayouts()[0].GetPackagePart().PartName;
            using var package = OPCPackage.Open(new MemoryStream(Save(original)));
            var masterXml = AddInterleavedExtensions(package.GetPart(masterName));
            var layoutXml = AddInterleavedExtensions(package.GetPart(layoutName));

            using var presentation = new XMLSlideShow(package);
            var master = presentation.GetSlideMasters()[0];
            var layout = master.GetSlideLayouts()[0];
            Assert.That(master.GetXmlObject().cSld, Is.Not.Null);
            Assert.That(master.GetXmlObject().clrMap, Is.Not.Null);
            Assert.That(master.GetXmlObject().sldLayoutIdLst, Is.Not.Null);
            Assert.That(layout.GetXmlObject().cSld, Is.Not.Null);
            Assert.That(master.GetXmlObject().AnyElements.Count(element => element.NamespaceURI == "urn:npoi:test"),
                Is.EqualTo(5));
            Assert.That(layout.GetXmlObject().AnyElements.Count(element => element.NamespaceURI == "urn:npoi:test"),
                Is.EqualTo(3));
            using (var unchanged = Open(Save(presentation)))
            {
                Assert.That(ReadBytes(unchanged.Package.GetPart(masterName)), Is.EqualTo(masterXml));
                Assert.That(ReadBytes(unchanged.Package.GetPart(layoutName)), Is.EqualTo(layoutXml));
            }

            master.GetXmlObject().cSld.SetAttribute("name", "Changed Master");
            layout.GetXmlObject().cSld.SetAttribute("name", "Changed Layout");
            presentation.CreateSlide(layout);
            using var reopened = Open(Save(presentation));
            var savedMaster = reopened.GetSlideMasters()[0];
            var savedLayout = savedMaster.GetSlideLayouts()[0];
            Assert.That(savedMaster.Name, Is.EqualTo("Changed Master"));
            Assert.That(savedLayout.Name, Is.EqualTo("Changed Layout"));
            AssertEditedXmlEquivalent(masterXml, savedMaster.GetPackagePart(), "Changed Master");
            AssertEditedXmlEquivalent(layoutXml, savedLayout.GetPackagePart(), "Changed Layout");
            using var again = Open(Save(reopened));
            AssertEditedXmlEquivalent(masterXml, again.Package.GetPart(masterName), "Changed Master");
            AssertEditedXmlEquivalent(layoutXml, again.Package.GetPart(layoutName), "Changed Layout");
        }

        [TestCase(false)]
        [TestCase(true)]
        public void BothCreationOverloadsReserveCompatibilitySlideIdentifiers(bool selectedLayout)
        {
            using var presentation = new XMLSlideShow();
            presentation.CreateSlide();
            var extension = new XmlDocument { XmlResolver = null };
            extension.LoadXml("<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">" +
                "<mc:Fallback><p:sldId xmlns:p=\"" + CT_Presentation.Namespace +
                "\" id=\"258\"/></mc:Fallback></mc:AlternateContent>");
            presentation.GetCTPresentation().sldIdLst.AnyElements = new[] { extension.DocumentElement };
            var slide = selectedLayout
                ? presentation.CreateSlide(presentation.GetSlideMasters()[0].GetSlideLayouts()[0])
                : presentation.CreateSlide();
            Assert.That(presentation.GetCTPresentation().sldIdLst.sldId.Last().id, Is.EqualTo(259U));
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlides().Last().GetPackagePart().PartName, Is.EqualTo(slide.GetPackagePart().PartName));
            Assert.That(reopened.GetCTPresentation().sldIdLst.sldId.Last().id, Is.EqualTo(259U));
        }

        [TestCase(false)]
        [TestCase(true)]
        public void BothCreationOverloadsRejectWrappedSlideListsBeforeMutation(bool selectedLayout)
        {
            using var presentation = new XMLSlideShow();
            var extension = new XmlDocument { XmlResolver = null };
            extension.LoadXml("<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">" +
                "<mc:Fallback><p:sldIdLst xmlns:p=\"" + CT_Presentation.Namespace +
                "\"/></mc:Fallback></mc:AlternateContent>");
            presentation.GetCTPresentation().AnyElements = new[] { extension.DocumentElement };
            AssertNoMutation<NotSupportedException>(presentation, () =>
            {
                if (selectedLayout)
                {
                    presentation.CreateSlide(presentation.GetSlideMasters()[0].GetSlideLayouts()[0]);
                }
                else
                {
                    presentation.CreateSlide();
                }
            });
        }

        [Test]
        public void NewPresentationExposesTypedMasterLayoutAndTheme()
        {
            using var presentation = new XMLSlideShow();
            var masters = presentation.GetSlideMasters();
            Assert.That(masters.Count, Is.EqualTo(1));
            Assert.That(((ICollection<XSLFSlideMaster>)masters).IsReadOnly, Is.True);
            var master = masters[0];
            var layout = master.GetSlideLayouts()[0];
            Assert.That(master.Name, Is.EqualTo("Blank Master"));
            Assert.That(layout.Name, Is.EqualTo("Blank"));
            Assert.That(layout.Type, Is.EqualTo("blank"));
            Assert.That(master.GetLayout("bLaNk"), Is.SameAs(layout));
            Assert.That(master.GetLayout("missing"), Is.Null);
            Assert.Throws<ArgumentNullException>(() => master.GetLayout(null));
            Assert.That(master.GetSlideShow(), Is.SameAs(presentation));
            Assert.That(layout.GetSlideShow(), Is.SameAs(presentation));
            Assert.That(master.GetMasterSheet(), Is.Null);
            Assert.That(layout.GetMasterSheet(), Is.SameAs(master));
            var theme = master.GetTheme();
            Assert.That(theme, Is.SameAs(layout.GetTheme()));
            Assert.That(theme.Name, Is.EqualTo("NPOI"));
            Assert.That(theme.MajorFont, Is.EqualTo("Arial"));
            Assert.That(theme.MinorFont, Is.EqualTo("Arial"));
            Assert.That(XSLFRelation.SLIDE_LAYOUT.RelationClass, Is.EqualTo(typeof(XSLFSlideLayout)));
            Assert.That(XSLFRelation.SLIDE_MASTER.RelationClass, Is.EqualTo(typeof(XSLFSlideMaster)));
            Assert.That(XSLFRelation.THEME.RelationClass, Is.EqualTo(typeof(XSLFTheme)));
        }

        [Test]
        public void SelectedLayoutAndInheritedNavigationSurviveReopening()
        {
            using var input = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("layouts.pptx");
            using var presentation = new XMLSlideShow(input);
            var master = presentation.GetSlideMasters()[0];
            var layout = master.GetSlideLayouts().First(item => item.Type == "title");
            var slide = presentation.CreateSlide(layout);
            Assert.That(slide.GetSlideLayout(), Is.SameAs(layout));
            Assert.That(slide.GetMasterSheet(), Is.SameAs(layout));
            Assert.That(slide.GetSlideMaster(), Is.SameAs(master));
            Assert.That(slide.GetTheme(), Is.SameAs(master.GetTheme()));
            Assert.That(slide.GetSlideShow(), Is.SameAs(presentation));
            Assert.That(slide.GetXmlObject().cSld["spTree", CT_Presentation.Namespace].ChildNodes.Count,
                Is.EqualTo(2), "Attaching a layout must not imply that placeholders have been copied.");

            using var reopened = Open(Save(presentation));
            var savedSlide = reopened.GetSlides().Last();
            Assert.That(savedSlide.GetSlideLayout().GetPackagePart().PartName.Name,
                Is.EqualTo(layout.GetPackagePart().PartName.Name));
            Assert.That(savedSlide.GetSlideLayout().Type, Is.EqualTo("title"));
            Assert.That(savedSlide.GetSlideMaster(), Is.SameAs(reopened.GetSlideMasters()[0]));
            Assert.That(savedSlide.GetTheme().Name, Is.EqualTo(master.GetTheme().Name));
        }

        [Test]
        public void MasterAndLayoutEnumerationUsesIdentifierListOrder()
        {
            using var input = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("themes.pptx");
            using var presentation = new XMLSlideShow(input);
            var masters = presentation.GetSlideMasters().ToArray();
            Assert.That(masters.Length, Is.GreaterThan(1));
            var layouts = masters[0].GetSlideLayouts();
            Assert.That(layouts.Length, Is.GreaterThan(1));
            var masterList = presentation.GetCTPresentation().sldMasterIdLst;
            masterList.PrependChild(masterList.LastChild);
            var layoutList = masters[0].GetXmlObject().sldLayoutIdLst;
            layoutList.PrependChild(layoutList.LastChild);
            Assert.That(presentation.GetSlideMasters()[0], Is.SameAs(masters.Last()));
            Assert.That(masters[0].GetSlideLayouts()[0], Is.SameAs(layouts.Last()));

            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlideMasters().Select(master => master.GetPackagePart().PartName.Name),
                Is.EqualTo(new[] { masters.Last() }.Concat(masters.Take(masters.Length - 1))
                    .Select(master => master.GetPackagePart().PartName.Name)));
            var savedMaster = reopened.GetSlideMasters().Single(master =>
                master.GetPackagePart().PartName.Equals(masters[0].GetPackagePart().PartName));
            Assert.That(savedMaster.GetSlideLayouts()[0].GetPackagePart().PartName.Name,
                Is.EqualTo(layouts.Last().GetPackagePart().PartName.Name));
        }

        [Test]
        public void LayoutFromAnotherReferencedMasterCanBeSelected()
        {
            using var input = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("themes.pptx");
            using var presentation = new XMLSlideShow(input);
            var master = presentation.GetSlideMasters().Last();
            var layout = master.GetSlideLayouts()[0];
            var slide = presentation.CreateSlide(layout);
            Assert.That(slide.GetSlideMaster(), Is.SameAs(master));
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlides().Last().GetSlideMaster(), Is.SameAs(reopened.GetSlideMasters().Last()));
        }

        [Test]
        public void ModifiedLayoutAndMasterRetainTheirUnmodeledXml()
        {
            using var presentation = new XMLSlideShow();
            var master = presentation.GetSlideMasters()[0];
            var layout = master.GetSlideLayouts()[0];
            master.GetXmlObject().cSld.SetAttribute("name", "Changed Master");
            layout.GetXmlObject().cSld.SetAttribute("name", "Changed Layout");
            layout.GetXmlObject().type = "cust";
            var extension = new XmlDocument { XmlResolver = null };
            extension.LoadXml("<extra:content xmlns:extra=\"urn:npoi:xslf:test\" value=\"preserve\"/>");
            layout.GetXmlObject().AnyElements =
                layout.GetXmlObject().AnyElements.Concat(new[] { extension.DocumentElement }).ToArray();
            presentation.CreateSlide(layout);
            using var reopened = Open(Save(presentation));
            var savedMaster = reopened.GetSlideMasters()[0];
            var savedLayout = savedMaster.GetSlideLayouts()[0];
            Assert.That(savedMaster.Name, Is.EqualTo("Changed Master"));
            Assert.That(savedLayout.Name, Is.EqualTo("Changed Layout"));
            Assert.That(savedLayout.Type, Is.EqualTo("cust"));
            Assert.That(savedLayout.GetXmlObject().AnyElements.Last().GetAttribute("value"), Is.EqualTo("preserve"));
            using var again = Open(Save(reopened));
            Assert.That(again.GetSlideMasters()[0].GetSlideLayouts()[0].Name, Is.EqualTo("Changed Layout"));
        }

        [Test]
        public void UntouchedMasterLayoutAndThemeBytesSurviveRepeatedSaves()
        {
            using var input = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("themes.pptx");
            using var presentation = new XMLSlideShow(input);
            var originals = presentation.Package.GetParts().Where(part =>
                part.ContentType == XSLFRelation.SLIDE_MASTER.ContentType ||
                part.ContentType == XSLFRelation.SLIDE_LAYOUT.ContentType ||
                part.ContentType == XSLFRelation.THEME.ContentType)
                .ToDictionary(part => part.PartName.Name, ReadBytes);
            foreach (var master in presentation.GetSlideMasters())
            {
                Assert.That(master.GetTheme().MajorFont, Is.Not.Empty);
                foreach (var layout in master.GetSlideLayouts())
                {
                    Assert.That(layout.GetTheme(), Is.SameAs(master.GetTheme()));
                    Assert.That(layout.GetSlideMaster(), Is.SameAs(master));
                }
            }
            presentation.CreateSlide(presentation.GetSlideMasters().Last().GetSlideLayouts().Last());
            Save(presentation);
            using var reopened = Open(Save(presentation));
            using var again = Open(Save(reopened));
            foreach (var original in originals)
            {
                Assert.That(ReadBytes(again.Package.GetPart(PackagingUriHelper.CreatePartName(original.Key))),
                    Is.EqualTo(original.Value), original.Key);
            }
        }

        [Test]
        public void NullAndForeignLayoutsAreRejectedWithoutMutation()
        {
            using var presentation = new XMLSlideShow();
            using var other = new XMLSlideShow();
            var foreignLayout = other.GetSlideMasters()[0].GetSlideLayouts()[0];
            AssertNoMutation<ArgumentNullException>(presentation, () => presentation.CreateSlide(null));
            AssertNoMutation<ArgumentException>(presentation, () => presentation.CreateSlide(foreignLayout));
        }

        [Test]
        public void AnUnregisteredWrapperForTheSamePackageIsRejected()
        {
            using var presentation = new XMLSlideShow();
            var part = presentation.GetSlideMasters()[0].GetSlideLayouts()[0].GetPackagePart();
            var layout = (XSLFSlideLayout)XSLFFactory.GetInstance().CreateDocumentPart(presentation, part);
            AssertNoMutation<POIXMLException>(presentation, () => presentation.CreateSlide(layout));
        }

        [TestCase(true)]
        [TestCase(false)]
        public void UnlistedMasterOrLayoutCannotBeUsedToCreateASlide(bool removeMaster)
        {
            using var presentation = new XMLSlideShow();
            var master = presentation.GetSlideMasters()[0];
            var layout = master.GetSlideLayouts()[0];
            var list = removeMaster ? presentation.GetCTPresentation().sldMasterIdLst :
                master.GetXmlObject().sldLayoutIdLst;
            list.RemoveChild(list.FirstChild);
            AssertNoMutation<ArgumentException>(presentation, () => presentation.CreateSlide(layout));
            AssertNoMutation<NotSupportedException>(presentation, () => presentation.CreateSlide());
        }

        [Test]
        public void LayoutWithoutAMasterBackReferenceCannotBeUsedToCreateASlide()
        {
            using var presentation = new XMLSlideShow();
            var layout = presentation.GetSlideMasters()[0].GetSlideLayouts()[0];
            var relationships = layout.GetPackagePart().GetRelationshipsByType(XSLFRelation.SLIDE_MASTER.Relation);
            layout.GetPackagePart().RemoveRelationship(relationships.GetRelationship(0).Id);
            AssertNoMutation<POIXMLException>(presentation, () => presentation.CreateSlide(layout));
        }

        [Test]
        public void LayoutPointingToADifferentMasterIsRejected()
        {
            using var input = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("themes.pptx");
            using var presentation = new XMLSlideShow(input);
            var masters = presentation.GetSlideMasters();
            var layout = masters[0].GetSlideLayouts()[0];
            var part = layout.GetPackagePart();
            var id = part.GetRelationshipsByType(XSLFRelation.SLIDE_MASTER.Relation).GetRelationship(0).Id;
            part.RemoveRelationship(id);
            part.AddRelationship(masters[1].GetPackagePart().PartName, TargetMode.Internal,
                XSLFRelation.SLIDE_MASTER.Relation, id);
            AssertNoMutation<POIXMLException>(presentation, () => presentation.CreateSlide(layout));
            AssertInvalid(Save(presentation));
        }

        [TestCase(true, "missing")]
        [TestCase(false, "missing")]
        [TestCase(true, "0")]
        [TestCase(false, "2147483647")]
        public void InvalidMasterAndLayoutReferencesFailExplicitly(bool masterReference, string value)
        {
            using var presentation = new XMLSlideShow();
            var master = presentation.GetSlideMasters()[0];
            var list = masterReference ? presentation.GetCTPresentation().sldMasterIdLst :
                master.GetXmlObject().sldLayoutIdLst;
            var entry = (XmlElement)list.FirstChild;
            if (value == "missing")
            {
                entry.SetAttribute("id", CT_Presentation.RelationshipsNamespace, value);
            }
            else
            {
                entry.SetAttribute("id", value);
            }
            AssertInvalid(Save(presentation));
        }

        [TestCase(true, false)]
        [TestCase(false, false)]
        [TestCase(true, true)]
        [TestCase(false, true)]
        public void DuplicateMasterAndLayoutIdsOrPartsFailExplicitly(bool masterReference, bool uniqueId)
        {
            using var presentation = new XMLSlideShow();
            var list = masterReference ? presentation.GetCTPresentation().sldMasterIdLst :
                presentation.GetSlideMasters()[0].GetXmlObject().sldLayoutIdLst;
            var duplicate = (XmlElement)list.FirstChild.CloneNode(true);
            if (uniqueId)
            {
                duplicate.SetAttribute("id", "2147483650");
            }
            list.AppendChild(duplicate);
            AssertInvalid(Save(presentation));
        }

        [TestCase(true)]
        [TestCase(false)]
        public void ExternalMasterAndLayoutReferencesAreRejected(bool masterReference)
        {
            using var presentation = new XMLSlideShow();
            var master = presentation.GetSlideMasters()[0];
            var owner = masterReference ? presentation.GetPackagePart() : master.GetPackagePart();
            var descriptor = masterReference ? XSLFRelation.SLIDE_MASTER : XSLFRelation.SLIDE_LAYOUT;
            var id = owner.GetRelationshipsByType(descriptor.Relation).GetRelationship(0).Id;
            owner.RemoveRelationship(id);
            owner.AddExternalRelationship("https://example.com/part.xml", descriptor.Relation, id);
            AssertInvalid(Save(presentation));
        }

        [TestCase(true)]
        [TestCase(false)]
        public void IncorrectMasterOrLayoutRelationshipTypesAreRejected(bool masterReference)
        {
            using var presentation = new XMLSlideShow();
            var master = presentation.GetSlideMasters()[0];
            var owner = masterReference ? presentation.GetPackagePart() : master.GetPackagePart();
            var descriptor = masterReference ? XSLFRelation.SLIDE_MASTER : XSLFRelation.SLIDE_LAYOUT;
            var relationship = owner.GetRelationshipsByType(descriptor.Relation).GetRelationship(0);
            var target = owner.GetRelatedPart(relationship);
            owner.RemoveRelationship(relationship.Id);
            owner.AddRelationship(target.PartName, TargetMode.Internal, XSLFRelation.THEME.Relation, relationship.Id);
            AssertInvalid(Save(presentation));
        }

        [TestCase(true)]
        [TestCase(false)]
        public void MissingMasterOrLayoutPartIsRejected(bool masterPart)
        {
            using var presentation = new XMLSlideShow();
            var master = presentation.GetSlideMasters()[0];
            var part = masterPart ? master.GetPackagePart() : master.GetSlideLayouts()[0].GetPackagePart();
            using var package = OPCPackage.Open(new MemoryStream(Save(presentation)));
            package.RemovePart(part.PartName);
            Assert.Throws<POIXMLException>(() => new XMLSlideShow(package));
        }

        [Test]
        public void MissingThemeIsReportedAsAbsentRatherThanInvented()
        {
            using var presentation = new XMLSlideShow();
            var master = presentation.GetSlideMasters()[0];
            var part = master.GetPackagePart();
            part.RemoveRelationship(part.GetRelationshipsByType(XSLFRelation.THEME.Relation).GetRelationship(0).Id);
            Assert.That(master.GetTheme(), Is.Null);
            Assert.That(master.GetSlideLayouts()[0].GetTheme(), Is.Null);
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlideMasters()[0].GetTheme(), Is.Null);
        }

        [Test]
        public void DuplicateThemeRelationshipsFailExplicitly()
        {
            using var presentation = new XMLSlideShow();
            var master = presentation.GetSlideMasters()[0];
            master.GetPackagePart().AddRelationship(master.GetTheme().GetPackagePart().PartName,
                TargetMode.Internal, XSLFRelation.THEME.Relation, "duplicateTheme");
            Assert.Throws<POIXMLException>(() => master.GetTheme());
            AssertInvalid(Save(presentation));
        }

        [TestCase(true)]
        [TestCase(false)]
        public void MissingMasterOrLayoutCommonSlideDataFailsExplicitly(bool masterPart)
        {
            using var presentation = new XMLSlideShow();
            var master = presentation.GetSlideMasters()[0];
            if (masterPart)
            {
                master.GetXmlObject().cSld = null;
            }
            else
            {
                master.GetSlideLayouts()[0].GetXmlObject().cSld = null;
            }
            AssertInvalid(Save(presentation));
        }

        [TestCase(false)]
        [TestCase(true)]
        public void InvalidThemeRootOrDocumentTypeDeclarationIsRejected(bool dtd)
        {
            using var presentation = new XMLSlideShow();
            var themeName = presentation.GetSlideMasters()[0].GetTheme().GetPackagePart().PartName;
            using var package = OPCPackage.Open(new MemoryStream(Save(presentation)));
            var theme = package.GetPart(themeName);
            theme.Clear();
            using (var output = theme.GetOutputStream())
            {
                var xml = Encoding.UTF8.GetBytes(dtd
                    ? "<!DOCTYPE theme [<!ENTITY text 'test'>]><a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>"
                    : "<notATheme/>");
                output.Write(xml, 0, xml.Length);
            }
            Assert.Throws<POIXMLException>(() => new XMLSlideShow(package));
        }

        private static void AssertNoMutation<T>(XMLSlideShow presentation, Action action) where T : Exception
        {
            var parts = presentation.Package.GetParts().Count;
            var relations = presentation.GetPackagePart().Relationships.Size;
            var slides = presentation.GetSlides().Count;
            Assert.Throws<T>(action);
            Assert.That(presentation.Package.GetParts().Count, Is.EqualTo(parts));
            Assert.That(presentation.GetPackagePart().Relationships.Size, Is.EqualTo(relations));
            Assert.That(presentation.GetSlides().Count, Is.EqualTo(slides));
        }

        private static byte[] AddInterleavedExtensions(PackagePart part)
        {
            var document = new XmlDocument { XmlResolver = null };
            using (var stream = part.GetInputStream())
            {
                document.Load(stream);
            }
            var root = document.DocumentElement;
            var children = root.ChildNodes.OfType<XmlElement>().ToArray();
            root.SetAttribute("xmlns:future", "urn:npoi:test");
            root.SetAttribute("Ignorable", "http://schemas.openxmlformats.org/markup-compatibility/2006", "future");
            for (var i = 0; i < children.Length; i++)
            {
                var extension = document.CreateElement("future", "extension", "urn:npoi:test");
                extension.SetAttribute("position", i.ToString(System.Globalization.CultureInfo.InvariantCulture));
                extension.InnerText = "A\r\nB";
                root.InsertBefore(extension, children[i]);
            }
            root.AppendChild(document.CreateElement("future", "trailing", "urn:npoi:test"));
            using var output = new MemoryStream();
            using (var writer = XmlWriter.Create(output, new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(false),
                NewLineHandling = NewLineHandling.Entitize
            }))
            {
                document.Save(writer);
            }
            var xml = output.ToArray();
            part.Clear();
            using (var stream = part.GetOutputStream())
            {
                stream.Write(xml, 0, xml.Length);
            }
            return xml;
        }

        private static void AssertEditedXmlEquivalent(byte[] original, PackagePart saved, string name)
        {
            var expected = XElement.Parse(Encoding.UTF8.GetString(original));
            var actual = XElement.Parse(Encoding.UTF8.GetString(ReadBytes(saved)));
            Assert.That(actual.GetNamespaceOfPrefix("future")?.NamespaceName, Is.EqualTo("urn:npoi:test"));
            expected.Element(XName.Get("cSld", CT_Presentation.Namespace)).SetAttributeValue("name", name);
            expected.DescendantsAndSelf().Attributes().Where(attribute => attribute.IsNamespaceDeclaration).Remove();
            actual.DescendantsAndSelf().Attributes().Where(attribute => attribute.IsNamespaceDeclaration).Remove();
            Assert.That(XNode.DeepEquals(expected, actual), Is.True);
        }

        private static void AssertInvalid(byte[] bytes)
        {
            using var package = OPCPackage.Open(new MemoryStream(bytes));
            Assert.Throws<POIXMLException>(() => new XMLSlideShow(package));
        }

        private static XMLSlideShow Open(byte[] bytes)
        {
            return new XMLSlideShow(new MemoryStream(bytes));
        }

        private static byte[] Save(XMLSlideShow presentation)
        {
            using var output = new MemoryStream();
            presentation.Write(output);
            return output.ToArray();
        }

        private static byte[] ReadBytes(PackagePart part)
        {
            using var input = part.GetInputStream();
            using var output = new MemoryStream();
            input.CopyTo(output);
            return output.ToArray();
        }
    }
}
