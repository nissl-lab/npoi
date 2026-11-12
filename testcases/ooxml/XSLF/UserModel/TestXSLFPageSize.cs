using System;
using System.IO;
using NPOI;
using NPOI.OpenXml4Net.OPC;
using NPOI.Util;
using NPOI.XSLF.UserModel;
using NUnit.Framework;
using SkiaSharp;

namespace TestCases.XSLF.UserModel
{
    [TestFixture]
    public class TestXSLFPageSize
    {
        [Test]
        public void DefaultPageSizeUsesPointsRatherThanPixelsOrEmus()
        {
            using var presentation = new XMLSlideShow();
            Assert.That(presentation.GetPageSize(), Is.EqualTo(new SKSize(720, 540)));
        }

        [TestCase(72f, 4032f)]
        [TestCase(4032f, 72f)]
        [TestCase(960f, 540f)]
        [TestCase(960.125f, 540.25f)]
        public void ValidSizesUseExactEmuConversionAndSurviveReopening(float width, float height)
        {
            using var presentation = new XMLSlideShow();
            presentation.CreateSlide();
            var notes = presentation.GetCTPresentation().notesSz.OuterXml;
            presentation.GetCTPresentation().sldSz.SetAttribute("hint", "urn:npoi:test", "retained");
            presentation.SetPageSize(new SKSize(width, height));
            var size = presentation.GetCTPresentation().sldSz;
            Assert.That(size.GetAttribute("cx"), Is.EqualTo(Units.ToEMU(width).ToString()));
            Assert.That(size.GetAttribute("cy"), Is.EqualTo(Units.ToEMU(height).ToString()));
            Assert.That(size.GetAttribute("type"), Is.EqualTo("custom"));
            Assert.That(presentation.GetCTPresentation().notesSz.OuterXml, Is.EqualTo(notes));
            Assert.That(size.GetAttribute("hint", "urn:npoi:test"), Is.EqualTo("retained"));
            using var reopened = new XMLSlideShow(new MemoryStream(Save(presentation)));
            Assert.That(reopened.GetPageSize().Width, Is.EqualTo(width).Within(0.001));
            Assert.That(reopened.GetPageSize().Height, Is.EqualTo(height).Within(0.001));
            Assert.That(reopened.GetCTPresentation().sldSz.GetAttribute("cx"), Is.EqualTo(Units.ToEMU(width).ToString()));
        }

        [TestCase(0f)]
        [TestCase(-1f)]
        [TestCase(71.99f)]
        [TestCase(4032.01f)]
        [TestCase(float.NaN)]
        [TestCase(float.PositiveInfinity)]
        [TestCase(float.NegativeInfinity)]
        [TestCase(float.MaxValue)]
        public void InvalidDimensionsFailWithoutChangingEitherAxis(float invalid)
        {
            using var presentation = new XMLSlideShow();
            var original = presentation.GetCTPresentation().sldSz.OuterXml;
            Assert.Throws<ArgumentOutOfRangeException>(() => presentation.SetPageSize(new SKSize(invalid, 540)));
            Assert.That(presentation.GetCTPresentation().sldSz.OuterXml, Is.EqualTo(original));
            Assert.Throws<ArgumentOutOfRangeException>(() => presentation.SetPageSize(new SKSize(720, invalid)));
            Assert.That(presentation.GetCTPresentation().sldSz.OuterXml, Is.EqualTo(original));
        }

        [TestCase("0")]
        [TestCase("914399")]
        [TestCase("51206401")]
        [TestCase("not-a-number")]
        [TestCase("99999999999999999999")]
        public void InvalidStoredSizeFailsExplicitly(string value)
        {
            using var presentation = new XMLSlideShow();
            presentation.GetCTPresentation().sldSz.SetAttribute("cx", value);
            Assert.Throws<POIXMLException>(() => presentation.GetPageSize());
        }

        [Test]
        public void MissingSizeHasNoSilentDefaultAndCanBeSetExplicitly()
        {
            using var presentation = new XMLSlideShow();
            presentation.GetCTPresentation().sldSz = null;
            Assert.Throws<POIXMLException>(() => presentation.GetPageSize());
            presentation.SetPageSize(new SKSize(720, 540));
            using var reopened = new XMLSlideShow(new MemoryStream(Save(presentation)));
            Assert.That(reopened.GetPageSize(), Is.EqualTo(new SKSize(720, 540)));
        }

        [Test]
        public void ReadOnlyAndClosedPresentationsRejectPageSizeChanges()
        {
            using var original = new XMLSlideShow();
            using var package = OPCPackage.Open(new MemoryStream(Save(original)), true);
            using var readOnly = new XMLSlideShow(package);
            Assert.Throws<InvalidOperationException>(() => readOnly.SetPageSize(new SKSize(960, 540)));
            Assert.That(readOnly.GetPageSize(), Is.EqualTo(new SKSize(720, 540)));
            original.Dispose();
            Assert.Throws<ObjectDisposedException>(() => original.SetPageSize(new SKSize(960, 540)));
        }

        private static byte[] Save(XMLSlideShow presentation)
        {
            using var output = new MemoryStream();
            presentation.Write(output);
            return output.ToArray();
        }
    }
}
