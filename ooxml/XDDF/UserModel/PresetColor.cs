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
==================================================================== */


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.XDDF.UserModel
{

    using NPOI.OpenXmlFormats.Dml;

    public enum PresetColor
    {
        AliceBlue,
        AntiqueWhite,
        Aqua,
        Aquamarine,
        Azure,
        Beige,
        Bisque,
        Black,
        BlanchedAlmond,
        Blue,
        BlueViolet,
        CadetBlue,
        Chartreuse,
        Chocolate,
        Coral,
        CornflowerBlue,
        Cornsilk,
        Crimson,
        Cyan,
        DeepPink,
        DeepSkyBlue,
        DimGray,
        DarkBlue,
        DarkCyan,
        DarkGoldenrod,
        DarkGray,
        DarkGreen,
        DarkKhaki,
        DarkMagenta,
        DarkOliveGreen,
        DarkOrange,
        DarkOrchid,
        DarkRed,
        DarkSalmon,
        DarkSeaGreen,
        DarkSlateBlue,
        DarkSlateGray,
        DarkTurquoise,
        DarkViolet,
        DodgerBlue,
        Firebrick,
        FloralWhite,
        ForestGreen,
        Fuchsia,
        Gainsboro,
        GhostWhite,
        Gold,
        Goldenrod,
        Gray,
        Green,
        GreenYellow,
        Honeydew,
        HotPink,
        IndianRed,
        Indigo,
        Ivory,
        Khaki,
        Lavender,
        LavenderBlush,
        LawnGreen,
        LemonChiffon,
        Lime,
        LimeGreen,
        Linen,
        LightBlue,
        LightCoral,
        LightCyan,
        LightGoldenrodYellow,
        LightGray,
        LightGreen,
        LightPink,
        LightSalmon,
        LightSeaGreen,
        LightSkyBlue,
        LightSlateGray,
        LightSteelBlue,
        LightYellow,
        Magenta,
        Maroon,
        MediumAquamarine,
        MediumBlue,
        MediumOrchid,
        MediumPurple,
        MediumSeaGreen,
        MediumSlateBlue,
        MediumSpringGreen,
        MediumTurquoise,
        MediumVioletRed,
        MidnightBlue,
        MintCream,
        MistyRose,
        Moccasin,
        NavajoWhite,
        Navy,
        OldLace,
        Olive,
        OliveDrab,
        Orange,
        OrangeRed,
        Orchid,
        PaleGoldenrod,
        PaleGreen,
        PaleTurquoise,
        PaleVioletRed,
        PapayaWhip,
        PeachPuff,
        Peru,
        Pink,
        Plum,
        PowderBlue,
        Purple,
        Red,
        RosyBrown,
        RoyalBlue,
        SaddleBrown,
        Salmon,
        SandyBrown,
        SeaGreen,
        SeaShell,
        Sienna,
        Silver,
        SkyBlue,
        SlateBlue,
        SlateGray,
        Snow,
        SpringGreen,
        SteelBlue,
        Tan,
        Teal,
        Thistle,
        Tomato,
        Turquoise,
        Violet,
        Wheat,
        White,
        WhiteSmoke,
        Yellow,
        YellowGreen
    }
    public static class PresetColorExtensions
    {
        private static Dictionary<ST_PresetColorVal, PresetColor> reverse = new Dictionary<ST_PresetColorVal, PresetColor>()
        {
            { ST_PresetColorVal.aliceBlue, PresetColor.AliceBlue },
            { ST_PresetColorVal.antiqueWhite, PresetColor.AntiqueWhite },
            { ST_PresetColorVal.aqua, PresetColor.Aqua },
            { ST_PresetColorVal.aquamarine, PresetColor.Aquamarine },
            { ST_PresetColorVal.azure, PresetColor.Azure },
            { ST_PresetColorVal.beige, PresetColor.Beige },
            { ST_PresetColorVal.bisque, PresetColor.Bisque },
            { ST_PresetColorVal.black, PresetColor.Black },
            { ST_PresetColorVal.blanchedAlmond, PresetColor.BlanchedAlmond },
            { ST_PresetColorVal.blue, PresetColor.Blue },
            { ST_PresetColorVal.blueViolet, PresetColor.BlueViolet },
            { ST_PresetColorVal.cadetBlue, PresetColor.CadetBlue },
            { ST_PresetColorVal.chartreuse, PresetColor.Chartreuse },
            { ST_PresetColorVal.chocolate, PresetColor.Chocolate },
            { ST_PresetColorVal.coral, PresetColor.Coral },
            { ST_PresetColorVal.cornflowerBlue, PresetColor.CornflowerBlue },
            { ST_PresetColorVal.cornsilk, PresetColor.Cornsilk },
            { ST_PresetColorVal.crimson, PresetColor.Crimson },
            { ST_PresetColorVal.cyan, PresetColor.Cyan },
            { ST_PresetColorVal.deepPink, PresetColor.DeepPink },
            { ST_PresetColorVal.deepSkyBlue, PresetColor.DeepSkyBlue },
            { ST_PresetColorVal.dimGray, PresetColor.DimGray },
            { ST_PresetColorVal.dkBlue, PresetColor.DarkBlue },
            { ST_PresetColorVal.dkCyan, PresetColor.DarkCyan },
            { ST_PresetColorVal.dkGoldenrod, PresetColor.DarkGoldenrod },
            { ST_PresetColorVal.dkGray, PresetColor.DarkGray },
            { ST_PresetColorVal.dkGreen, PresetColor.DarkGreen },
            { ST_PresetColorVal.dkKhaki, PresetColor.DarkKhaki },
            { ST_PresetColorVal.dkMagenta, PresetColor.DarkMagenta },
            { ST_PresetColorVal.dkOliveGreen, PresetColor.DarkOliveGreen },
            { ST_PresetColorVal.dkOrange, PresetColor.DarkOrange },
            { ST_PresetColorVal.dkOrchid, PresetColor.DarkOrchid },
            { ST_PresetColorVal.dkRed, PresetColor.DarkRed },
            { ST_PresetColorVal.dkSalmon, PresetColor.DarkSalmon },
            { ST_PresetColorVal.dkSeaGreen, PresetColor.DarkSeaGreen },
            { ST_PresetColorVal.dkSlateBlue, PresetColor.DarkSlateBlue },
            { ST_PresetColorVal.dkSlateGray, PresetColor.DarkSlateGray },
            { ST_PresetColorVal.dkTurquoise, PresetColor.DarkTurquoise },
            { ST_PresetColorVal.dkViolet, PresetColor.DarkViolet },
            { ST_PresetColorVal.dodgerBlue, PresetColor.DodgerBlue },
            { ST_PresetColorVal.firebrick, PresetColor.Firebrick },
            { ST_PresetColorVal.floralWhite, PresetColor.FloralWhite },
            { ST_PresetColorVal.forestGreen, PresetColor.ForestGreen },
            { ST_PresetColorVal.fuchsia, PresetColor.Fuchsia },
            { ST_PresetColorVal.gainsboro, PresetColor.Gainsboro },
            { ST_PresetColorVal.ghostWhite, PresetColor.GhostWhite },
            { ST_PresetColorVal.gold, PresetColor.Gold },
            { ST_PresetColorVal.goldenrod, PresetColor.Goldenrod },
            { ST_PresetColorVal.gray, PresetColor.Gray },
            { ST_PresetColorVal.green, PresetColor.Green },
            { ST_PresetColorVal.greenYellow, PresetColor.GreenYellow },
            { ST_PresetColorVal.honeydew, PresetColor.Honeydew },
            { ST_PresetColorVal.hotPink, PresetColor.HotPink },
            { ST_PresetColorVal.indianRed, PresetColor.IndianRed },
            { ST_PresetColorVal.indigo, PresetColor.Indigo },
            { ST_PresetColorVal.ivory, PresetColor.Ivory },
            { ST_PresetColorVal.khaki, PresetColor.Khaki },
            { ST_PresetColorVal.lavender, PresetColor.Lavender },
            { ST_PresetColorVal.lavenderBlush, PresetColor.LavenderBlush },
            { ST_PresetColorVal.lawnGreen, PresetColor.LawnGreen },
            { ST_PresetColorVal.lemonChiffon, PresetColor.LemonChiffon },
            { ST_PresetColorVal.lime, PresetColor.Lime },
            { ST_PresetColorVal.limeGreen, PresetColor.LimeGreen },
            { ST_PresetColorVal.linen, PresetColor.Linen },
            { ST_PresetColorVal.ltBlue, PresetColor.LightBlue },
            { ST_PresetColorVal.ltCoral, PresetColor.LightCoral },
            { ST_PresetColorVal.ltCyan, PresetColor.LightCyan },
            { ST_PresetColorVal.ltGoldenrodYellow, PresetColor.LightGoldenrodYellow },
            { ST_PresetColorVal.ltGray, PresetColor.LightGray },
            { ST_PresetColorVal.ltGreen, PresetColor.LightGreen },
            { ST_PresetColorVal.ltPink, PresetColor.LightPink },
            { ST_PresetColorVal.ltSalmon, PresetColor.LightSalmon },
            { ST_PresetColorVal.ltSeaGreen, PresetColor.LightSeaGreen },
            { ST_PresetColorVal.ltSkyBlue, PresetColor.LightSkyBlue },
            { ST_PresetColorVal.ltSlateGray, PresetColor.LightSlateGray },
            { ST_PresetColorVal.ltSteelBlue, PresetColor.LightSteelBlue },
            { ST_PresetColorVal.ltYellow, PresetColor.LightYellow },
            { ST_PresetColorVal.magenta, PresetColor.Magenta },
            { ST_PresetColorVal.maroon, PresetColor.Maroon },
            { ST_PresetColorVal.medAquamarine, PresetColor.MediumAquamarine },
            { ST_PresetColorVal.medBlue, PresetColor.MediumBlue },
            { ST_PresetColorVal.medOrchid, PresetColor.MediumOrchid },
            { ST_PresetColorVal.medPurple, PresetColor.MediumPurple },
            { ST_PresetColorVal.medSeaGreen, PresetColor.MediumSeaGreen },
            { ST_PresetColorVal.medSlateBlue, PresetColor.MediumSlateBlue },
            { ST_PresetColorVal.medSpringGreen, PresetColor.MediumSpringGreen },
            { ST_PresetColorVal.medTurquoise, PresetColor.MediumTurquoise },
            { ST_PresetColorVal.medVioletRed, PresetColor.MediumVioletRed },
            { ST_PresetColorVal.midnightBlue, PresetColor.MidnightBlue },
            { ST_PresetColorVal.mintCream, PresetColor.MintCream },
            { ST_PresetColorVal.mistyRose, PresetColor.MistyRose },
            { ST_PresetColorVal.moccasin, PresetColor.Moccasin },
            { ST_PresetColorVal.navajoWhite, PresetColor.NavajoWhite },
            { ST_PresetColorVal.navy, PresetColor.Navy },
            { ST_PresetColorVal.oldLace, PresetColor.OldLace },
            { ST_PresetColorVal.olive, PresetColor.Olive },
            { ST_PresetColorVal.oliveDrab, PresetColor.OliveDrab },
            { ST_PresetColorVal.orange, PresetColor.Orange },
            { ST_PresetColorVal.orangeRed, PresetColor.OrangeRed },
            { ST_PresetColorVal.orchid, PresetColor.Orchid },
            { ST_PresetColorVal.paleGoldenrod, PresetColor.PaleGoldenrod },
            { ST_PresetColorVal.paleGreen, PresetColor.PaleGreen },
            { ST_PresetColorVal.paleTurquoise, PresetColor.PaleTurquoise },
            { ST_PresetColorVal.paleVioletRed, PresetColor.PaleVioletRed },
            { ST_PresetColorVal.papayaWhip, PresetColor.PapayaWhip },
            { ST_PresetColorVal.peachPuff, PresetColor.PeachPuff },
            { ST_PresetColorVal.peru, PresetColor.Peru },
            { ST_PresetColorVal.pink, PresetColor.Pink },
            { ST_PresetColorVal.plum, PresetColor.Plum },
            { ST_PresetColorVal.powderBlue, PresetColor.PowderBlue },
            { ST_PresetColorVal.purple, PresetColor.Purple },
            { ST_PresetColorVal.red, PresetColor.Red },
            { ST_PresetColorVal.rosyBrown, PresetColor.RosyBrown },
            { ST_PresetColorVal.royalBlue, PresetColor.RoyalBlue },
            { ST_PresetColorVal.saddleBrown, PresetColor.SaddleBrown },
            { ST_PresetColorVal.salmon, PresetColor.Salmon },
            { ST_PresetColorVal.sandyBrown, PresetColor.SandyBrown },
            { ST_PresetColorVal.seaGreen, PresetColor.SeaGreen },
            { ST_PresetColorVal.seaShell, PresetColor.SeaShell },
            { ST_PresetColorVal.sienna, PresetColor.Sienna },
            { ST_PresetColorVal.silver, PresetColor.Silver },
            { ST_PresetColorVal.skyBlue, PresetColor.SkyBlue },
            { ST_PresetColorVal.slateBlue, PresetColor.SlateBlue },
            { ST_PresetColorVal.slateGray, PresetColor.SlateGray },
            { ST_PresetColorVal.snow, PresetColor.Snow },
            { ST_PresetColorVal.springGreen, PresetColor.SpringGreen },
            { ST_PresetColorVal.steelBlue, PresetColor.SteelBlue },
            { ST_PresetColorVal.tan, PresetColor.Tan },
            { ST_PresetColorVal.teal, PresetColor.Teal },
            { ST_PresetColorVal.thistle, PresetColor.Thistle },
            { ST_PresetColorVal.tomato, PresetColor.Tomato },
            { ST_PresetColorVal.turquoise, PresetColor.Turquoise },
            { ST_PresetColorVal.violet, PresetColor.Violet },
            { ST_PresetColorVal.wheat, PresetColor.Wheat },
            { ST_PresetColorVal.white, PresetColor.White },
            { ST_PresetColorVal.whiteSmoke, PresetColor.WhiteSmoke },
            { ST_PresetColorVal.yellow, PresetColor.Yellow },
            { ST_PresetColorVal.yellowGreen, PresetColor.YellowGreen },
        };
        public static PresetColor ValueOf(ST_PresetColorVal value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_PresetColorVal", nameof(value));
        }
        public static ST_PresetColorVal ToST_PresetColorVal(this PresetColor value)
        {
            switch(value)
            {
                case PresetColor.AliceBlue:
                    return ST_PresetColorVal.aliceBlue;
                case PresetColor.AntiqueWhite:
                    return ST_PresetColorVal.antiqueWhite;
                case PresetColor.Aqua:
                    return ST_PresetColorVal.aqua;
                case PresetColor.Aquamarine:
                    return ST_PresetColorVal.aquamarine;
                case PresetColor.Azure:
                    return ST_PresetColorVal.azure;
                case PresetColor.Beige:
                    return ST_PresetColorVal.beige;
                case PresetColor.Bisque:
                    return ST_PresetColorVal.bisque;
                case PresetColor.Black:
                    return ST_PresetColorVal.black;
                case PresetColor.BlanchedAlmond:
                    return ST_PresetColorVal.blanchedAlmond;
                case PresetColor.Blue:
                    return ST_PresetColorVal.blue;
                case PresetColor.BlueViolet:
                    return ST_PresetColorVal.blueViolet;
                case PresetColor.CadetBlue:
                    return ST_PresetColorVal.cadetBlue;
                case PresetColor.Chartreuse:
                    return ST_PresetColorVal.chartreuse;
                case PresetColor.Chocolate:
                    return ST_PresetColorVal.chocolate;
                case PresetColor.Coral:
                    return ST_PresetColorVal.coral;
                case PresetColor.CornflowerBlue:
                    return ST_PresetColorVal.cornflowerBlue;
                case PresetColor.Cornsilk:
                    return ST_PresetColorVal.cornsilk;
                case PresetColor.Crimson:
                    return ST_PresetColorVal.crimson;
                case PresetColor.Cyan:
                    return ST_PresetColorVal.cyan;
                case PresetColor.DeepPink:
                    return ST_PresetColorVal.deepPink;
                case PresetColor.DeepSkyBlue:
                    return ST_PresetColorVal.deepSkyBlue;
                case PresetColor.DimGray:
                    return ST_PresetColorVal.dimGray;
                case PresetColor.DarkBlue:
                    return ST_PresetColorVal.dkBlue;
                case PresetColor.DarkCyan:
                    return ST_PresetColorVal.dkCyan;
                case PresetColor.DarkGoldenrod:
                    return ST_PresetColorVal.dkGoldenrod;
                case PresetColor.DarkGray:
                    return ST_PresetColorVal.dkGray;
                case PresetColor.DarkGreen:
                    return ST_PresetColorVal.dkGreen;
                case PresetColor.DarkKhaki:
                    return ST_PresetColorVal.dkKhaki;
                case PresetColor.DarkMagenta:
                    return ST_PresetColorVal.dkMagenta;
                case PresetColor.DarkOliveGreen:
                    return ST_PresetColorVal.dkOliveGreen;
                case PresetColor.DarkOrange:
                    return ST_PresetColorVal.dkOrange;
                case PresetColor.DarkOrchid:
                    return ST_PresetColorVal.dkOrchid;
                case PresetColor.DarkRed:
                    return ST_PresetColorVal.dkRed;
                case PresetColor.DarkSalmon:
                    return ST_PresetColorVal.dkSalmon;
                case PresetColor.DarkSeaGreen:
                    return ST_PresetColorVal.dkSeaGreen;
                case PresetColor.DarkSlateBlue:
                    return ST_PresetColorVal.dkSlateBlue;
                case PresetColor.DarkSlateGray:
                    return ST_PresetColorVal.dkSlateGray;
                case PresetColor.DarkTurquoise:
                    return ST_PresetColorVal.dkTurquoise;
                case PresetColor.DarkViolet:
                    return ST_PresetColorVal.dkViolet;
                case PresetColor.DodgerBlue:
                    return ST_PresetColorVal.dodgerBlue;
                case PresetColor.Firebrick:
                    return ST_PresetColorVal.firebrick;
                case PresetColor.FloralWhite:
                    return ST_PresetColorVal.floralWhite;
                case PresetColor.ForestGreen:
                    return ST_PresetColorVal.forestGreen;
                case PresetColor.Fuchsia:
                    return ST_PresetColorVal.fuchsia;
                case PresetColor.Gainsboro:
                    return ST_PresetColorVal.gainsboro;
                case PresetColor.GhostWhite:
                    return ST_PresetColorVal.ghostWhite;
                case PresetColor.Gold:
                    return ST_PresetColorVal.gold;
                case PresetColor.Goldenrod:
                    return ST_PresetColorVal.goldenrod;
                case PresetColor.Gray:
                    return ST_PresetColorVal.gray;
                case PresetColor.Green:
                    return ST_PresetColorVal.green;
                case PresetColor.GreenYellow:
                    return ST_PresetColorVal.greenYellow;
                case PresetColor.Honeydew:
                    return ST_PresetColorVal.honeydew;
                case PresetColor.HotPink:
                    return ST_PresetColorVal.hotPink;
                case PresetColor.IndianRed:
                    return ST_PresetColorVal.indianRed;
                case PresetColor.Indigo:
                    return ST_PresetColorVal.indigo;
                case PresetColor.Ivory:
                    return ST_PresetColorVal.ivory;
                case PresetColor.Khaki:
                    return ST_PresetColorVal.khaki;
                case PresetColor.Lavender:
                    return ST_PresetColorVal.lavender;
                case PresetColor.LavenderBlush:
                    return ST_PresetColorVal.lavenderBlush;
                case PresetColor.LawnGreen:
                    return ST_PresetColorVal.lawnGreen;
                case PresetColor.LemonChiffon:
                    return ST_PresetColorVal.lemonChiffon;
                case PresetColor.Lime:
                    return ST_PresetColorVal.lime;
                case PresetColor.LimeGreen:
                    return ST_PresetColorVal.limeGreen;
                case PresetColor.Linen:
                    return ST_PresetColorVal.linen;
                case PresetColor.LightBlue:
                    return ST_PresetColorVal.ltBlue;
                case PresetColor.LightCoral:
                    return ST_PresetColorVal.ltCoral;
                case PresetColor.LightCyan:
                    return ST_PresetColorVal.ltCyan;
                case PresetColor.LightGoldenrodYellow:
                    return ST_PresetColorVal.ltGoldenrodYellow;
                case PresetColor.LightGray:
                    return ST_PresetColorVal.ltGray;
                case PresetColor.LightGreen:
                    return ST_PresetColorVal.ltGreen;
                case PresetColor.LightPink:
                    return ST_PresetColorVal.ltPink;
                case PresetColor.LightSalmon:
                    return ST_PresetColorVal.ltSalmon;
                case PresetColor.LightSeaGreen:
                    return ST_PresetColorVal.ltSeaGreen;
                case PresetColor.LightSkyBlue:
                    return ST_PresetColorVal.ltSkyBlue;
                case PresetColor.LightSlateGray:
                    return ST_PresetColorVal.ltSlateGray;
                case PresetColor.LightSteelBlue:
                    return ST_PresetColorVal.ltSteelBlue;
                case PresetColor.LightYellow:
                    return ST_PresetColorVal.ltYellow;
                case PresetColor.Magenta:
                    return ST_PresetColorVal.magenta;
                case PresetColor.Maroon:
                    return ST_PresetColorVal.maroon;
                case PresetColor.MediumAquamarine:
                    return ST_PresetColorVal.medAquamarine;
                case PresetColor.MediumBlue:
                    return ST_PresetColorVal.medBlue;
                case PresetColor.MediumOrchid:
                    return ST_PresetColorVal.medOrchid;
                case PresetColor.MediumPurple:
                    return ST_PresetColorVal.medPurple;
                case PresetColor.MediumSeaGreen:
                    return ST_PresetColorVal.medSeaGreen;
                case PresetColor.MediumSlateBlue:
                    return ST_PresetColorVal.medSlateBlue;
                case PresetColor.MediumSpringGreen:
                    return ST_PresetColorVal.medSpringGreen;
                case PresetColor.MediumTurquoise:
                    return ST_PresetColorVal.medTurquoise;
                case PresetColor.MediumVioletRed:
                    return ST_PresetColorVal.medVioletRed;
                case PresetColor.MidnightBlue:
                    return ST_PresetColorVal.midnightBlue;
                case PresetColor.MintCream:
                    return ST_PresetColorVal.mintCream;
                case PresetColor.MistyRose:
                    return ST_PresetColorVal.mistyRose;
                case PresetColor.Moccasin:
                    return ST_PresetColorVal.moccasin;
                case PresetColor.NavajoWhite:
                    return ST_PresetColorVal.navajoWhite;
                case PresetColor.Navy:
                    return ST_PresetColorVal.navy;
                case PresetColor.OldLace:
                    return ST_PresetColorVal.oldLace;
                case PresetColor.Olive:
                    return ST_PresetColorVal.olive;
                case PresetColor.OliveDrab:
                    return ST_PresetColorVal.oliveDrab;
                case PresetColor.Orange:
                    return ST_PresetColorVal.orange;
                case PresetColor.OrangeRed:
                    return ST_PresetColorVal.orangeRed;
                case PresetColor.Orchid:
                    return ST_PresetColorVal.orchid;
                case PresetColor.PaleGoldenrod:
                    return ST_PresetColorVal.paleGoldenrod;
                case PresetColor.PaleGreen:
                    return ST_PresetColorVal.paleGreen;
                case PresetColor.PaleTurquoise:
                    return ST_PresetColorVal.paleTurquoise;
                case PresetColor.PaleVioletRed:
                    return ST_PresetColorVal.paleVioletRed;
                case PresetColor.PapayaWhip:
                    return ST_PresetColorVal.papayaWhip;
                case PresetColor.PeachPuff:
                    return ST_PresetColorVal.peachPuff;
                case PresetColor.Peru:
                    return ST_PresetColorVal.peru;
                case PresetColor.Pink:
                    return ST_PresetColorVal.pink;
                case PresetColor.Plum:
                    return ST_PresetColorVal.plum;
                case PresetColor.PowderBlue:
                    return ST_PresetColorVal.powderBlue;
                case PresetColor.Purple:
                    return ST_PresetColorVal.purple;
                case PresetColor.Red:
                    return ST_PresetColorVal.red;
                case PresetColor.RosyBrown:
                    return ST_PresetColorVal.rosyBrown;
                case PresetColor.RoyalBlue:
                    return ST_PresetColorVal.royalBlue;
                case PresetColor.SaddleBrown:
                    return ST_PresetColorVal.saddleBrown;
                case PresetColor.Salmon:
                    return ST_PresetColorVal.salmon;
                case PresetColor.SandyBrown:
                    return ST_PresetColorVal.sandyBrown;
                case PresetColor.SeaGreen:
                    return ST_PresetColorVal.seaGreen;
                case PresetColor.SeaShell:
                    return ST_PresetColorVal.seaShell;
                case PresetColor.Sienna:
                    return ST_PresetColorVal.sienna;
                case PresetColor.Silver:
                    return ST_PresetColorVal.silver;
                case PresetColor.SkyBlue:
                    return ST_PresetColorVal.skyBlue;
                case PresetColor.SlateBlue:
                    return ST_PresetColorVal.slateBlue;
                case PresetColor.SlateGray:
                    return ST_PresetColorVal.slateGray;
                case PresetColor.Snow:
                    return ST_PresetColorVal.snow;
                case PresetColor.SpringGreen:
                    return ST_PresetColorVal.springGreen;
                case PresetColor.SteelBlue:
                    return ST_PresetColorVal.steelBlue;
                case PresetColor.Tan:
                    return ST_PresetColorVal.tan;
                case PresetColor.Teal:
                    return ST_PresetColorVal.teal;
                case PresetColor.Thistle:
                    return ST_PresetColorVal.thistle;
                case PresetColor.Tomato:
                    return ST_PresetColorVal.tomato;
                case PresetColor.Turquoise:
                    return ST_PresetColorVal.turquoise;
                case PresetColor.Violet:
                    return ST_PresetColorVal.violet;
                case PresetColor.Wheat:
                    return ST_PresetColorVal.wheat;
                case PresetColor.White:
                    return ST_PresetColorVal.white;
                case PresetColor.WhiteSmoke:
                    return ST_PresetColorVal.whiteSmoke;
                case PresetColor.Yellow:
                    return ST_PresetColorVal.yellow;
                case PresetColor.YellowGreen:
                    return ST_PresetColorVal.yellowGreen;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}


