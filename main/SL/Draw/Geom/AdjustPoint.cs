using System;

namespace NPOI.SL.Draw.Geom
{
    public class AdjustPoint : AdjustPointIf
    {
        public string X { get; set; }

        public bool IsSetX
        {
            get
            {
                return X!=null;
            }
        }

        public string Y { get; set; }
        public bool IsSetY
        {
            get {
                return Y!=null;
            }
        }

        public override bool Equals(object o)
        {
            if(this == o)
                return true;
            if(!(o is AdjustPoint)) return false;
            AdjustPoint that = (AdjustPoint) o;
            return string.Equals(X, that.X, StringComparison.Ordinal) &&
                    string.Equals(Y, that.Y, StringComparison.Ordinal);
        }
        public override int GetHashCode()
        {
            unchecked
            {
                return ((X?.GetHashCode() ?? 0) * 397) ^ (Y?.GetHashCode() ?? 0);
            }
        }
    }
}
