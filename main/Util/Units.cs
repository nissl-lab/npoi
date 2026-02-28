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
        /// <summary>
        /// In Escher absolute distances are specified in
        /// English Metric Units (EMUs), occasionally referred to as A units;
        /// there are 360000 EMUs per centimeter, 914400 EMUs per inch, 12700 EMUs per point.
        /// </summary>
        public static int EMU_PER_PIXEL = 9525;
        public static int EMU_PER_POINT = 12700;
        public static int EMU_PER_CENTIMETER = 360000;

        /// <summary>
        /// Master DPI (576 pixels per inch).
        /// Used by the reference coordinate system in PowerPoint (HSLF)
        /// </summary>
        public static int MASTER_DPI = 576;

        /// <summary>
        /// Pixels DPI (96 pixels per inch)
        /// </summary>
        public static int PIXEL_DPI = 96;

        /// <summary>
        /// Points DPI (72 pixels per inch)
        /// </summary>
        public static int POINT_DPI = 72;

        /// <summary>
        /// Width of one "standard character" of the default font in pixels. Same for Calibri and Arial.
        /// "Standard character" defined as the widest digit character in the given font.
        /// Copied from XSSFWorkbook, since that isn't available here.
        /// <p/>
        /// Note this is only valid for workbooks using the default Excel font.
        /// <p/>
        /// Would be nice to eventually support arbitrary document default fonts.
        /// </summary>
        public static float DEFAULT_CHARACTER_WIDTH = 7.0017f;

        /// <summary>
        /// Column widths are in fractional characters, this is the EMU equivalent.
        /// One character is defined as the widest value for the integers 0-9 in the
        /// default font.
        /// </summary>
        public static int EMU_PER_CHARACTER = (int)(EMU_PER_PIXEL * DEFAULT_CHARACTER_WIDTH);

        /// <summary>
        /// Converts points to EMUs
        /// </summary>
        /// <param name="value"></param>
        /// <returns>EMUs</returns>
        public static int ToEMU(double value)
        {
            return (int)Math.Round(EMU_PER_POINT * value);
        }
        /// <summary>
        /// Converts pixels to EMUs
        /// </summary>
        /// <param name="pixels">pixels</param>
        /// <return>EMUs</return>
        public static int PixelToEMU(int pixels)
        {
            return pixels * EMU_PER_PIXEL;
        }

        /// <summary>
        /// Converts EMUs to points
        /// </summary>
        /// <param name="emu">emu</param>
        /// <return>points</return>
        public static double ToPoints(long emu)
        {
            return (double)emu / EMU_PER_POINT;
        }

        /// <summary>
        /// Converts a value of type FixedPoint to a floating point
        /// </summary>
        /// <param name="fixedPoint">value in fixed point notation</param>
        /// <return>floating point (double)</return>
        /// <see href="http://msdn.microsoft.com/en-us/library/dd910765(v=office.12).aspx">[MS-OSHARED] - 2.2.1.6 FixedPoint</see>
        public static double FixedPointToDecimal(int fixedPoint)
        {
            int i = (fixedPoint >> 16);
            int f = (fixedPoint >> 0) & 0xFFFF;
            return i + f / 65536d;
        }

        /// <summary>
        /// Converts a value of type floating point to a FixedPoint
        /// </summary>
        /// <param name="floatPoint">value in floating point notation</param>
        /// <return>fixedPoint value in fixed points notation</return>
        /// <see href="http://msdn.microsoft.com/en-us/library/dd910765(v=office.12).aspx">[MS-OSHARED] - 2.2.1.6 FixedPoint</see>
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

        public static int CharactersToEMU(double characters)
        {
            return (int)characters * EMU_PER_CHARACTER;
        }

        /// <summary>
        /// </summary>
        /// <param name="columnWidth">specified in 256ths of a standard character</param>
        /// <return>equivalent EMUs</return>
        public static int ColumnWidthToEMU(int columnWidth)
        {
            return CharactersToEMU(columnWidth / 256d);
        }

        /// <summary>
        /// </summary>
        /// <param name="twips">(1/20th of a point) typically used for row heights</param>
        /// <return>equivalent EMUs</return>
        public static int TwipsToEMU(short twips)
        {
            return (int)(twips / 20d * EMU_PER_POINT);
        }
    }
}