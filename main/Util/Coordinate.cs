using System;

namespace NPOI.Util
{
    public class Coords //<T> where T : INumber<T> {
    {   

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// x coordinate
        /// </summary>
        public long x { get; internal set; }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// y coordinate
        /// </summary>
        public long y { get; internal set; }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// constructer
        /// </summary>
        public Coords()
        {
        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// constructor
        /// </summary>
        /// <param name="x">x coordinate</param>
        /// <param name="y">y coordinate</param>
        public Coords(long X, long Y)
        {
            x = X;
            y = Y;
        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// constructor
        /// </summary>
        /// <param name="val">coordinate</param>
        public Coords(Coords val)
        {
            x = val.x;
            y = val.y;
        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// Find the minimum
        /// </summary>
        /// <param name="val"></param>
        public void Min(Coords val)
        {
            if(x > val.x)
            {
                x = val.x;
            }
            if(y > val.y)
            {
                y = val.y;
            }

        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// Find the maximum
        /// </summary>
        /// <param name="val"></param>
        public void Max(Coords val)
        {
            if(x < val.x)
            {
                x = val.x;
            }
            if(y < val.y)
            {
                y = val.y;
            }
        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// addition value
        /// </summary>
        /// <param name="val"></param>
        public void Add(Coords val)
        {
            x += val.x;
            y += val.y;
        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// subtraction value
        /// </summary>
        /// <param name="val"></param>
        public void Sub(Coords val)
        {
            x -= val.x;
            y -= val.y;
        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// 
        /// </summary>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <returns></returns>
        public static Coords operator +(Coords left, Coords right)
        {
            return new Coords(left.x + right.x, left.y + right.y);
        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// 
        /// </summary>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <returns></returns>
        public static Coords operator -(Coords left, Coords right)
        {
            return new Coords(left.x - right.x, left.y - right.y);
        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// 
        /// </summary>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <returns></returns>
        public static Coords operator *(Coords left, Coords right)
        {
            return new Coords(left.x * right.x, left.y * right.y);
        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        public static Coords operator *(Coords V, double Constant)
        {
            return new Coords((long) (V.x*Constant), (long) (V.y*Constant));
        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Coords"></param>
        /// <returns></returns>
        public double InnerProduct(Coords Coords)
        {
            var v = this*Coords;
            return (double) (v.x+v.y)/(Norm()*Coords.Norm());
        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// 
        /// </summary>
        /// <returns>Norm</returns>
        public long Norm()
        {
            return (long) Math.Sqrt(x*x+y*y);
        }
    }
    public class DblVect2D
    {

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// x coordinate
        /// </summary>
        public double x { get; internal set; }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// y coordinate
        /// </summary>
        public double y { get; internal set; }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// constructer
        /// </summary>
        public DblVect2D()
        {
        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// constructor
        /// </summary>
        /// <param name="x">x coordinate</param>
        /// <param name="y">y coordinate</param>
        public DblVect2D(double X, double Y)
        {
            x = X;
            y = Y;
        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// constructor
        /// </summary>
        /// <param name="val">coordinate</param>
        public DblVect2D(DblVect2D val)
        {
            x = val.x;
            y = val.y;
        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// Find the minimum
        /// </summary>
        /// <param name="val"></param>
        public void Min(DblVect2D val)
        {
            if(x > val.x)
            {
                x = val.x;
            }
            if(y > val.y)
            {
                y = val.y;
            }

        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// Find the maximum
        /// </summary>
        /// <param name="val"></param>
        public void Max(DblVect2D val)
        {
            if(x < val.x)
            {
                x = val.x;
            }
            if(y < val.y)
            {
                y = val.y;
            }
        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// addition value
        /// </summary>
        /// <param name="val"></param>
        public void Add(DblVect2D val)
        {
            x += val.x;
            y += val.y;
        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// subtraction value
        /// </summary>
        /// <param name="val"></param>
        public void Sub(DblVect2D val)
        {
            x -= val.x;
            y -= val.y;
        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// 
        /// </summary>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <returns></returns>
        public static DblVect2D operator +(DblVect2D left, DblVect2D right)
        {
            return new DblVect2D(left.x + right.x, left.y + right.y);
        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// 
        /// </summary>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <returns></returns>
        public static DblVect2D operator -(DblVect2D left, DblVect2D right)
        {
            return new DblVect2D(left.x - right.x, left.y - right.y);
        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// 
        /// </summary>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <returns></returns>
        public static DblVect2D operator *(DblVect2D left, DblVect2D right)
        {
            return new DblVect2D(left.x * right.x, left.y * right.y);
        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        public static DblVect2D operator *(DblVect2D V, double Constant)
        {
            return new DblVect2D((long) (V.x*Constant), (long) (V.y*Constant));
        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        public static DblVect2D operator /(DblVect2D V, double Constant)
        {
            return new DblVect2D((long) (V.x/Constant), (long) (V.y/Constant));
        }
        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Coords"></param>
        /// <returns></returns>
        public double InnerProduct(DblVect2D vect)
        {
            var v = this*vect;
            return (v.x+v.y)/(Norm()*vect.Norm());
        }

        /*=======+=========+=========+=========+=========+=========+=========+=========+														+*/
        /// <summary>
        /// 
        /// </summary>
        /// <returns>Norm</returns>
        public double Norm()
        {
            return Math.Sqrt(x*x+y*y);
        }

        public static DblVect2D Conv(Coords val)
        {
            return new DblVect2D(val.x, val.y);
        }
    }
}
