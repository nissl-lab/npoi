using System;
using System.Collections.Generic;
using System.Globalization;
using NPOI.SL.Draw.Geom;
using NUnit.Framework;
using SkiaSharp;

namespace TestCases.SL.Draw.Geom
{
    [TestFixture]
    public class TestGeometry
    {
        private static Context CreateContext()
        {
            return new Context(new CustomGeometry(), SKRect.Create(10, 20, 320, 160), null);
        }

        [TestCase("h", 160)]
        [TestCase("ssd8", 20)]
        [TestCase("wd2", 160)]
        [TestCase("wd3", 320.0 / 3)]
        [TestCase("wd4", 80)]
        [TestCase("wd5", 64)]
        [TestCase("wd6", 320.0 / 6)]
        [TestCase("wd8", 40)]
        [TestCase("wd10", 32)]
        [TestCase("wd32", 10)]
        [TestCase("l", 10)]
        [TestCase("t", 20)]
        [TestCase("r", 330)]
        [TestCase("b", 180)]
        [TestCase("hc", 170)]
        [TestCase("vc", 100)]
        public void EvaluatesBuiltInGuides(string name, double expected)
        {
            Assert.That(CreateContext().GetValue(name), Is.EqualTo(expected).Within(1e-9));
            Assert.That(BuiltInGuide.ValueOf(name).Name, Is.EqualTo(name));
        }

        [TestCase("val 7", 7)]
        [TestCase("abs -7", 7)]
        [TestCase("sqrt 81", 9)]
        [TestCase("max 3 7", 7)]
        [TestCase("min 3 7", 3)]
        [TestCase("*/ 6 4 3", 8)]
        [TestCase("+/ 6 4 2", 5)]
        [TestCase("+- 6 4 3", 7)]
        [TestCase("?: 1 7 3", 7)]
        [TestCase("?: 0 7 3", 3)]
        [TestCase("at2 0 1", 5400000)]
        [TestCase("sin 2 5400000", 2)]
        [TestCase("cos 2 0", 2)]
        [TestCase("tan 2 2700000", 2)]
        [TestCase("cat2 5 3 4", 3)]
        [TestCase("sat2 5 3 4", 4)]
        [TestCase("pin 2 9 7", 7)]
        [TestCase("mod 3 4 12", 13)]
        [TestCase("*/ 6 4 0", 0)]
        [TestCase("+/ 6 4 0", 0)]
        [TestCase("val\t7", 7)]
        public void EvaluatesFormulas(string formula, double expected)
        {
            var guide = new GuideIf { Fmla = formula };
            Assert.That(guide.Evaluate(CreateContext()), Is.EqualTo(expected).Within(1e-9));
        }

        [Test]
        public void EvaluatesAdjustmentsBeforeDependentGuides()
        {
            var geometry = new CustomGeometry();
            geometry.Adjusts.Add(new GuideIf { Name = "adj1", Fmla = "val 40" });
            geometry.Guides.Add(new GuideIf { Name = "g1", Fmla = "*/ w adj1 100" });
            geometry.Guides.Add(new GuideIf { Name = "g2", Fmla = "+- g1 2 0" });

            var context = new Context(geometry, SKRect.Create(10, 20, 320, 160), null);
            Assert.That(context.GetValue("g2"), Is.EqualTo(130));
        }

        [Test]
        public void ParsesNumbersIndependentlyOfCurrentCulture()
        {
            var previousCulture = CultureInfo.CurrentCulture;
            try
            {
                CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
                var context = CreateContext();
                Assert.That(context.GetValue("1.5"), Is.EqualTo(1.5));
                Assert.That(context.GetValue("-2.5e2"), Is.EqualTo(-250));
            }
            finally
            {
                CultureInfo.CurrentCulture = previousCulture;
            }
        }

        [TestCase(null)]
        [TestCase("")]
        [TestCase(" ")]
        [TestCase("unknown")]
        [TestCase("adj1")]
        [TestCase("ssd9")]
        public void RejectsUnknownValues(string name)
        {
            Assert.Throws<ArgumentException>(() => CreateContext().GetValue(name));
        }

        [TestCase(null)]
        [TestCase("")]
        [TestCase("val")]
        [TestCase("val 1 2")]
        [TestCase("unknown 1")]
        [TestCase("0 1 2 3")]
        public void RejectsMalformedFormulas(string formula)
        {
            Assert.Throws<ArgumentException>(() => new GuideIf { Fmla = formula }.Evaluate(CreateContext()));
        }

        [Test]
        public void EqualAdjustmentPointsWorkAsDictionaryKeys()
        {
            var point = new AdjustPoint { X = "wd2", Y = "hd2" };
            var equivalent = new AdjustPoint { X = "wd2", Y = "hd2" };
            var values = new Dictionary<AdjustPoint, int> { [point] = 7 };
            Assert.That(values[equivalent], Is.EqualTo(7));
            Assert.That(new AdjustPoint(), Is.EqualTo(new AdjustPoint()));
        }
    }
}
