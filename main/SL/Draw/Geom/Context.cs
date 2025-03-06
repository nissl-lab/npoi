using SixLabors.ImageSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace NPOI.SL.Draw.Geom
{
    public class Context
    {
        private static Regex DOUBLE_PATTERN = new Regex(
        "[\\x00-\\x20]*[+-]?(NaN|Infinity|((((\\p{Digit}+)(\\.)?((\\p{Digit}+)?)" +
        "([eE][+-]?(\\p{Digit}+))?)|(\\.(\\p{Digit}+)([eE][+-]?(\\p{Digit}+))?)|" +
        "(((0[xX](\\p{XDigit}+)(\\.)?)|(0[xX](\\p{XDigit}+)?(\\.)(\\p{XDigit}+)))" +
        "[pP][+-]?(\\p{Digit}+)))[fFdD]?))[\\x00-\\x20]*", RegexOptions.Compiled);
        private Dictionary<string, double> _ctx= new Dictionary<string, double>();
        private IAdjustableShape _props;
        private Rectangle _anchor;

        public Context(CustomGeometry geom, Rectangle anchor, IAdjustableShape props)
        {
            _props = props;
            _anchor = anchor;
            foreach(GuideIf gd in geom.adjusts)
            {
                Evaluate(gd);
            }
            foreach(GuideIf gd in geom.guides)
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
        internal Rectangle GetShapeAnchor()
        {
            return _anchor;
        }

        internal GuideIf GetAdjustValue(String name)
        {
            return _props.GetAdjustValue(name);
        }
        public double GetValue(String key)
        {
            if(DOUBLE_PATTERN.Matches(key).Count>0)
            {
                return Double.Parse(key);
            }

            // BuiltInGuide throws IllegalArgumentException if key is not defined
            return _ctx.ContainsKey(key) ? _ctx[key] : Evaluate(BuiltInGuide.ValueOf(key));
        }
    }
}
