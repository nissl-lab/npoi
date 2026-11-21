using SkiaSharp;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace NPOI.SL.Draw.Geom
{
    public class Context
    {
        private Dictionary<string, double> _ctx= new Dictionary<string, double>();
        private IAdjustableShape _props;
        private SKRect _anchor;

        public Context(CustomGeometry geom, SKRect anchor, IAdjustableShape props)
        {
            if (geom == null)
                throw new ArgumentNullException(nameof(geom));
            _props = props;
            _anchor = anchor;
            foreach(GuideIf gd in geom.Adjusts)
            {
                Evaluate(gd);
            }
            foreach(GuideIf gd in geom.Guides)
            {
                Evaluate(gd);
            }
        }

        public double Evaluate(Formula fmla)
        {
            double result = fmla.Evaluate(this);
            if(fmla is GuideIf) {
                String key = ((GuideIf)fmla).Name;
                if(key != null)
                {
                    _ctx[key]=result;
                }
            }
            return result;
        }
        internal SKRect GetShapeAnchor()
        {
            return _anchor;
        }

        internal GuideIf GetAdjustValue(String name)
        {
            return _props?.GetAdjustValue(name);
        }
        public double GetValue(String key)
        {
            if (string.IsNullOrWhiteSpace(key))
                throw new ArgumentException("A geometry value or guide name is required.", nameof(key));
            if (double.TryParse(key, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
            {
                return value;
            }

            return _ctx.TryGetValue(key, out value) ? value : Evaluate(BuiltInGuide.ValueOf(key));
        }
    }
}
