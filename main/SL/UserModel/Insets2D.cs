using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SL.UserModel
{
    public sealed class Insets2D:ICloneable
    {
        public double Top { get; set; }
        public double Left { get; set; }
        public double Bottom { get; set; }
        public double Right { get; set; }

        /// <summary>
        /// Creates and initializes a new <code>Insets</code> object with the specified top, left, bottom, and right insets.
        /// </summary>
        /// <param name="top">the inset from the top</param>
        /// <param name="left">the inset from the left</param>
        /// <param name="bottom">the inset from the bottom</param>
        /// <param name="right">the inset from the right</param>
        public Insets2D(double top, double left, double bottom, double right)
        {
            this.Top = top;
            this.Left = left;
            this.Bottom = bottom;
            this.Right = right;
        }
        public override bool Equals(object obj)
        {
            if(obj is Insets2D) {
                Insets2D insets = (Insets2D)obj;
                return ((Top == insets.Top) && (Left == insets.Left) &&
                    (Bottom == insets.Bottom) && (Right == insets.Right));
            }
            return false;
        }
        public override int GetHashCode()
        {
            double sum1 = Left + Bottom;
            double sum2 = Right + Top;
            double val1 = sum1 * (sum1 + 1)/2 + Left;
            double val2 = sum2 * (sum2 + 1)/2 + Top;
            double sum3 = val1 + val2;
            return (int) (sum3 * (sum3 + 1)/2 + val2);
        }
        public override string ToString()
        {
            return this.GetType().Name+ "[Top="  + Top + ",Left=" + Left + ",Bottom=" + Bottom + ",Right=" + Right + "]";
        }

        public object Clone()
        {
            return new Insets2D(Top, Left, Bottom, Right);
        }
    }
}
