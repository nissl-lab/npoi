using System;
using System.Collections.Generic;

namespace NPOI.SL.Draw.Geom
{
    public class BuiltInGuide:Formula
    {
        private static readonly Dictionary<string, BuiltInGuide> indexes = new Dictionary<string, BuiltInGuide>(StringComparer.Ordinal);
        public static readonly BuiltInGuide _3cd4 = new BuiltInGuide("3cd4");
        public static readonly BuiltInGuide _3cd8 = new BuiltInGuide("3cd8");
        public static readonly BuiltInGuide _5cd8 = new BuiltInGuide("5cd8");
        public static readonly BuiltInGuide _7cd8 = new BuiltInGuide("7cd8");
        public static readonly BuiltInGuide b = new BuiltInGuide("b");
        public static readonly BuiltInGuide cd2 = new BuiltInGuide("cd2");
        public static readonly BuiltInGuide cd4 = new BuiltInGuide("cd4");
        public static readonly BuiltInGuide cd8 = new BuiltInGuide("cd8");
        public static readonly BuiltInGuide hc = new BuiltInGuide("hc");
        public static readonly BuiltInGuide h = new BuiltInGuide("h");
        public static readonly BuiltInGuide hd2 = new BuiltInGuide("hd2");
        public static readonly BuiltInGuide hd3 = new BuiltInGuide("hd3");
        public static readonly BuiltInGuide hd4 = new BuiltInGuide("hd4");
        public static readonly BuiltInGuide hd5 = new BuiltInGuide("hd5");
        public static readonly BuiltInGuide hd6 = new BuiltInGuide("hd6");
        public static readonly BuiltInGuide hd8 = new BuiltInGuide("hd8");
        public static readonly BuiltInGuide l = new BuiltInGuide("l");
        public static readonly BuiltInGuide ls = new BuiltInGuide("ls");
        public static readonly BuiltInGuide r = new BuiltInGuide("r");
        public static readonly BuiltInGuide ss = new BuiltInGuide("ss");
        public static readonly BuiltInGuide ssd2 = new BuiltInGuide("ssd2");
        public static readonly BuiltInGuide ssd4 = new BuiltInGuide("ssd4");
        public static readonly BuiltInGuide ssd6 = new BuiltInGuide("ssd6");
        public static readonly BuiltInGuide ssd8 = new BuiltInGuide("ssd8");
        public static readonly BuiltInGuide ssd16 = new BuiltInGuide("ssd16");
        public static readonly BuiltInGuide ssd32= new BuiltInGuide("ssd32");
        public static readonly BuiltInGuide t= new BuiltInGuide("t");
        public static readonly BuiltInGuide vc= new BuiltInGuide("vc");
        public static readonly BuiltInGuide w= new BuiltInGuide("w");
        public static readonly BuiltInGuide w2= new BuiltInGuide("wd2");
        public static readonly BuiltInGuide w3= new BuiltInGuide("wd3");
        public static readonly BuiltInGuide w4= new BuiltInGuide("wd4");
        public static readonly BuiltInGuide w5= new BuiltInGuide("wd5");
        public static readonly BuiltInGuide w6= new BuiltInGuide("wd6");
        public static readonly BuiltInGuide w8= new BuiltInGuide("wd8");
        public static readonly BuiltInGuide w10= new BuiltInGuide("wd10");
        public static readonly BuiltInGuide w32= new BuiltInGuide("wd32");

        private string name=null;
        protected BuiltInGuide(string name)
        {
            this.name=name;
            indexes.Add(name, this);
        }
        public string Name {
            get {
                return this.name;
            }
        }

        public static BuiltInGuide ValueOf(string name)
        {
            if (name == null)
                throw new ArgumentNullException(nameof(name));
            if (indexes.TryGetValue(name, out var guide))
                return guide;
            throw new ArgumentException("Unknown built-in geometry guide: " + name, nameof(name));
        }
        public override double Evaluate(Context ctx)
        {
            var anchor = ctx.GetShapeAnchor();
            double height = anchor.Height, width = anchor.Width, ss = Math.Min(width, height);
            switch(this.name)
            {
                case "3cd4":
                    // 3 circles div 4: 3 x 360 / 4 = 270
                    return 270 * OOXML_DEGREE;
                case "3cd8":
                    // 3 circles div 8: 3 x 360 / 8 = 135
                    return 135 * OOXML_DEGREE;
                case "5cd8":
                    // 5 circles div 8: 5 x 360 / 8 = 225
                    return 225 * OOXML_DEGREE;
                case "7cd8":
                    // 7 circles div 8: 7 x 360 / 8 = 315
                    return 315 * OOXML_DEGREE;
                case "t":
                    // top
                    return anchor.Top;
                case "b":
                    // bottom
                    return anchor.Bottom;
                case "l":
                    // left
                    return anchor.Left;
                case "r":
                    // right
                    return anchor.Right;
                case "cd2":
                    // circle div 2": 360 / 2 = 180
                    return 180 * OOXML_DEGREE;
                case "cd4":
                    // circle div 4": 360 / 4 = 90
                    return 90 * OOXML_DEGREE;
                case "cd8":
                    // circle div 8": 360 / 8 = 45
                    return 45 * OOXML_DEGREE;
                case "hc":
                    // horizontal center
                    return anchor.Left + width / 2.0;
                case "h":
                    // height
                    return height;
                case "hd2":
                    // height div 2
                    return height / 2.0;
                case "hd3":
                    // height div 3
                    return height / 3.0;
                case "hd4":
                    // height div 4
                    return height / 4.0;
                case "hd5":
                    // height div 5
                    return height / 5.0;
                case "hd6":
                    // height div 6
                    return height / 6.0;
                case "hd8":
                    // height div 8
                    return height / 8.0;
                case "ls":
                    // long side
                    return Math.Max(width, height);
                case "ss":
                    // short side
                    return ss;
                case "ssd2":
                    // short side div 2
                    return ss / 2.0;
                case "ssd4":
                    // short side div 4
                    return ss / 4.0;
                case "ssd6":
                    // short side div 6
                    return ss / 6.0;
                case "ssd8":
                    // short side div 8
                    return ss / 8.0;
                case "ssd16":
                    // short side div 16
                    return ss / 16.0;
                case "ssd32":
                    // short side div 32
                    return ss / 32.0;
                case "vc":
                    // vertical center
                    return anchor.Top + height / 2.0;
                case "w":
                    // width
                    return width;
                case "wd2":
                    // width div 2
                    return width / 2.0;
                case "wd3":
                    // width div 3
                    return width / 3.0;
                case "wd4":
                    // width div 4
                    return width / 4.0;
                case "wd5":
                    // width div 5
                    return width / 5.0;
                case "wd6":
                    // width div 6
                    return width / 6.0;
                case "wd8":
                    // width div 8
                    return width / 8.0;
                case "wd10":
                    // width div 10
                    return width / 10.0;
                case "wd32":
                    // width div 32
                    return width / 32.0;
                default:
                    throw new InvalidOperationException("Unsupported built-in geometry guide: " + name);
            }
        }
    }
}
