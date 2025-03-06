using Org.BouncyCastle.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
            return Objects.Equals(X, that.X) &&
                    Objects.Equals(Y, that.Y);
        }
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
}
