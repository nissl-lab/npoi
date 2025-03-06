using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SL.Draw.Geom
{
    public abstract class Formula
    {
        /**
         * OOXML units are 60000ths of a degree 
         */
        public const double OOXML_DEGREE = 60000;
        public abstract double Evaluate(Context ctx);
        protected double RadianToDegree(double angle)
        {
            return angle * (180.0 / Math.PI);
        }
        protected double DegreeToRadian(double angle)
        {
            return Math.PI * angle / 180.0;
        }
    }
}
