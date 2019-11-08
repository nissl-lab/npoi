/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */using System;
namespace NPOI.Util
{

    /**
     * @author Yegor Kozlov
     */
    public class Units
    {
        /**
         * In Escher absolute distances are specified in
         * English Metric Units (EMUs), occasionally referred to as A units;
         * there are 360000 EMUs per centimeter, 914400 EMUs per inch, 12700 EMUs per point.
         */
        public static int EMU_PER_PIXEL = 9525;
        public static int EMU_PER_POINT = 12700;
        public static int EMU_PER_CENTIMETER = 360000;

        /**
         * Master DPI (576 pixels per inch).
         * Used by the reference coordinate system in PowerPoint (HSLF)
         */
        public static int MASTER_DPI = 576;

        /**
         * Pixels DPI (96 pixels per inch)
         */
        public static int PIXEL_DPI = 96;

        /**
         * Points DPI (72 pixels per inch)
         */
        public static int POINT_DPI = 72;
        /// <summary>
        /// Converts points to EMUs
        /// </summary>
        /// <param name="value"></param>
        /// <returns>EMUs</returns>
        public static int ToEMU(double value)
        {
            return (int)Math.Round(EMU_PER_POINT * value);
        }
        /**
         * Converts pixels to EMUs
         * @param pixels pixels
         * @return EMUs
         */
        public static int PixelToEMU(int pixels)
        {
            return pixels * EMU_PER_PIXEL;
        }
        public static double ToPoints(long emu)
        {
            return (double)emu / EMU_PER_POINT;
        }
        
        /**
         * Converts a value of type FixedPoint to a decimal number
         *
         * @param fixedPoint
         * @return decimal number
         * 
         * @see <a href="http://msdn.microsoft.com/en-us/library/dd910765(v=office.12).aspx">[MS-OSHARED] - 2.2.1.6 FixedPoint</a>
         */
        public static double FixedPointToDecimal(int fixedPoint)
        {
            int i = (fixedPoint >> 16);
            int f = (fixedPoint >> 0) & 0xFFFF;
            return i + f / 65536.0;
        }

        /**
         * Converts a value of type floating point to a FixedPoint
         *
         * @param floatPoint
         * @return fixedPoint
         * 
         * @see <a href="http://msdn.microsoft.com/en-us/library/dd910765(v=office.12).aspx">[MS-OSHARED] - 2.2.1.6 FixedPoint</a>
         */
        public static int DoubleToFixedPoint(double floatPoint)
        {
            double fractionalPart = floatPoint % 1d;
            double integralPart = floatPoint - fractionalPart;
            int i = (int)Math.Floor(integralPart);
            //int f = (int)Math.rint(fractionalPart * 65536d);
            int f = (int)Math.Round(fractionalPart * 65536d, MidpointRounding.ToEven);
            return (i << 16) | (f & 0xFFFF);
        }

        public static double MasterToPoints(int masterDPI)
        {
            double points = masterDPI;
            points *= POINT_DPI;
            points /= MASTER_DPI;
            return points;
        }

        public static int PointsToMaster(double points)
        {
            points *= MASTER_DPI;
            points /= POINT_DPI;
            //return (int)Math.rint(points);
            return (int)Math.Round(points, MidpointRounding.ToEven);
        }

        public static int PointsToPixel(double points)
        {
            points *= PIXEL_DPI;
            points /= POINT_DPI;
            //return (int)Math.rint(points);
            return (int)Math.Round(points, MidpointRounding.ToEven);
        }

        public static double PixelToPoints(int pixel)
        {
            double points = pixel;
            points *= POINT_DPI;
            points /= PIXEL_DPI;
            return points;
        }
    }
}