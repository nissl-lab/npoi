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

    public enum PresetPattern
    {
        Cross,
        DashDownwardDiagonal,
        DashHorizontal,
        DashUpwardDiagonal,
        DashVertical,
        DiagonalBrick,
        DiagonalCross,
        Divot,
        DarkDownwardDiagonal,
        DarkHorizontal,
        DarkUpwardDiagonal,
        DarkVertical,
        DownwardDiagonal,
        DottedDiamond,
        DottedGrid,
        Horizontal,
        HorizontalBrick,
        LargeCheckerBoard,
        LargeConfetti,
        LargeGrid,
        LightDownwardDiagonal,
        LightHorizontal,
        LightUpwardDiagonal,
        LightVertical,
        NarrowHorizontal,
        NarrowVertical,
        OpenDiamond,
        Percent5,
        Percent10,
        Percent20,
        Percent25,
        Percent30,
        Percent40,
        Percent50,
        Percent60,
        Percent70,
        Percent75,
        Percent80,
        Percent90,
        Plaid,
        Shingle,
        SmallCheckerBoard,
        SmallConfetti,
        SmallGrid,
        SolidDiamond,
        Sphere,
        Trellis,
        UpwardDiagonal,
        Vertical,
        Wave,
        Weave,
        WideDownwardDiagonal,
        WideUpwardDiagonal,
        ZigZag
    }
    public static class PresetPatternExtensions
    {
        private static Dictionary<ST_PresetPatternVal, PresetPattern> reverse = new Dictionary<ST_PresetPatternVal, PresetPattern>()
        {
            { ST_PresetPatternVal.cross, PresetPattern.Cross },
            { ST_PresetPatternVal.dashDnDiag, PresetPattern.DashDownwardDiagonal },
            { ST_PresetPatternVal.dashHorz, PresetPattern.DashHorizontal },
            { ST_PresetPatternVal.dashUpDiag, PresetPattern.DashUpwardDiagonal },
            { ST_PresetPatternVal.dashVert, PresetPattern.DashVertical },
            { ST_PresetPatternVal.diagBrick, PresetPattern.DiagonalBrick },
            { ST_PresetPatternVal.diagCross, PresetPattern.DiagonalCross },
            { ST_PresetPatternVal.divot, PresetPattern.Divot },
            { ST_PresetPatternVal.dkDnDiag, PresetPattern.DarkDownwardDiagonal },
            { ST_PresetPatternVal.dkHorz, PresetPattern.DarkHorizontal },
            { ST_PresetPatternVal.dkUpDiag, PresetPattern.DarkUpwardDiagonal },
            { ST_PresetPatternVal.dkVert, PresetPattern.DarkVertical },
            { ST_PresetPatternVal.dnDiag, PresetPattern.DownwardDiagonal },
            { ST_PresetPatternVal.dotDmnd, PresetPattern.DottedDiamond },
            { ST_PresetPatternVal.dotGrid, PresetPattern.DottedGrid },
            { ST_PresetPatternVal.horz, PresetPattern.Horizontal },
            { ST_PresetPatternVal.horzBrick, PresetPattern.HorizontalBrick },
            { ST_PresetPatternVal.lgCheck, PresetPattern.LargeCheckerBoard },
            { ST_PresetPatternVal.lgConfetti, PresetPattern.LargeConfetti },
            { ST_PresetPatternVal.lgGrid, PresetPattern.LargeGrid },
            { ST_PresetPatternVal.ltDnDiag, PresetPattern.LightDownwardDiagonal },
            { ST_PresetPatternVal.ltHorz, PresetPattern.LightHorizontal },
            { ST_PresetPatternVal.ltUpDiag, PresetPattern.LightUpwardDiagonal },
            { ST_PresetPatternVal.ltVert, PresetPattern.LightVertical },
            { ST_PresetPatternVal.narHorz, PresetPattern.NarrowHorizontal },
            { ST_PresetPatternVal.narVert, PresetPattern.NarrowVertical },
            { ST_PresetPatternVal.openDmnd, PresetPattern.OpenDiamond },
            { ST_PresetPatternVal.pct5, PresetPattern.Percent5 },
            { ST_PresetPatternVal.pct10, PresetPattern.Percent10 },
            { ST_PresetPatternVal.pct20, PresetPattern.Percent20 },
            { ST_PresetPatternVal.pct25, PresetPattern.Percent25 },
            { ST_PresetPatternVal.pct30, PresetPattern.Percent30 },
            { ST_PresetPatternVal.pct40, PresetPattern.Percent40 },
            { ST_PresetPatternVal.pct50, PresetPattern.Percent50 },
            { ST_PresetPatternVal.pct60, PresetPattern.Percent60 },
            { ST_PresetPatternVal.pct70, PresetPattern.Percent70 },
            { ST_PresetPatternVal.pct75, PresetPattern.Percent75 },
            { ST_PresetPatternVal.pct80, PresetPattern.Percent80 },
            { ST_PresetPatternVal.pct90, PresetPattern.Percent90 },
            { ST_PresetPatternVal.plaid, PresetPattern.Plaid },
            { ST_PresetPatternVal.shingle, PresetPattern.Shingle },
            { ST_PresetPatternVal.smCheck, PresetPattern.SmallCheckerBoard },
            { ST_PresetPatternVal.smConfetti, PresetPattern.SmallConfetti },
            { ST_PresetPatternVal.smGrid, PresetPattern.SmallGrid },
            { ST_PresetPatternVal.solidDmnd, PresetPattern.SolidDiamond },
            { ST_PresetPatternVal.sphere, PresetPattern.Sphere },
            { ST_PresetPatternVal.trellis, PresetPattern.Trellis },
            { ST_PresetPatternVal.upDiag, PresetPattern.UpwardDiagonal },
            { ST_PresetPatternVal.vert, PresetPattern.Vertical },
            { ST_PresetPatternVal.wave, PresetPattern.Wave },
            { ST_PresetPatternVal.weave, PresetPattern.Weave },
            { ST_PresetPatternVal.wdDnDiag, PresetPattern.WideDownwardDiagonal },
            { ST_PresetPatternVal.wdUpDiag, PresetPattern.WideUpwardDiagonal },
            { ST_PresetPatternVal.zigZag, PresetPattern.ZigZag },
        };
        public static PresetPattern ValueOf(ST_PresetPatternVal value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_PresetPatternVal", nameof(value));
        }
        public static ST_PresetPatternVal ToST_PresetPatternVal(this PresetPattern value)
        {
            switch(value)
            {
                case PresetPattern.Cross:
                    return ST_PresetPatternVal.cross;
                case PresetPattern.DashDownwardDiagonal:
                    return ST_PresetPatternVal.dashDnDiag;
                case PresetPattern.DashHorizontal:
                    return ST_PresetPatternVal.dashHorz;
                case PresetPattern.DashUpwardDiagonal:
                    return ST_PresetPatternVal.dashUpDiag;
                case PresetPattern.DashVertical:
                    return ST_PresetPatternVal.dashVert;
                case PresetPattern.DiagonalBrick:
                    return ST_PresetPatternVal.diagBrick;
                case PresetPattern.DiagonalCross:
                    return ST_PresetPatternVal.diagCross;
                case PresetPattern.Divot:
                    return ST_PresetPatternVal.divot;
                case PresetPattern.DarkDownwardDiagonal:
                    return ST_PresetPatternVal.dkDnDiag;
                case PresetPattern.DarkHorizontal:
                    return ST_PresetPatternVal.dkHorz;
                case PresetPattern.DarkUpwardDiagonal:
                    return ST_PresetPatternVal.dkUpDiag;
                case PresetPattern.DarkVertical:
                    return ST_PresetPatternVal.dkVert;
                case PresetPattern.DownwardDiagonal:
                    return ST_PresetPatternVal.dnDiag;
                case PresetPattern.DottedDiamond:
                    return ST_PresetPatternVal.dotDmnd;
                case PresetPattern.DottedGrid:
                    return ST_PresetPatternVal.dotGrid;
                case PresetPattern.Horizontal:
                    return ST_PresetPatternVal.horz;
                case PresetPattern.HorizontalBrick:
                    return ST_PresetPatternVal.horzBrick;
                case PresetPattern.LargeCheckerBoard:
                    return ST_PresetPatternVal.lgCheck;
                case PresetPattern.LargeConfetti:
                    return ST_PresetPatternVal.lgConfetti;
                case PresetPattern.LargeGrid:
                    return ST_PresetPatternVal.lgGrid;
                case PresetPattern.LightDownwardDiagonal:
                    return ST_PresetPatternVal.ltDnDiag;
                case PresetPattern.LightHorizontal:
                    return ST_PresetPatternVal.ltHorz;
                case PresetPattern.LightUpwardDiagonal:
                    return ST_PresetPatternVal.ltUpDiag;
                case PresetPattern.LightVertical:
                    return ST_PresetPatternVal.ltVert;
                case PresetPattern.NarrowHorizontal:
                    return ST_PresetPatternVal.narHorz;
                case PresetPattern.NarrowVertical:
                    return ST_PresetPatternVal.narVert;
                case PresetPattern.OpenDiamond:
                    return ST_PresetPatternVal.openDmnd;
                case PresetPattern.Percent5:
                    return ST_PresetPatternVal.pct5;
                case PresetPattern.Percent10:
                    return ST_PresetPatternVal.pct10;
                case PresetPattern.Percent20:
                    return ST_PresetPatternVal.pct20;
                case PresetPattern.Percent25:
                    return ST_PresetPatternVal.pct25;
                case PresetPattern.Percent30:
                    return ST_PresetPatternVal.pct30;
                case PresetPattern.Percent40:
                    return ST_PresetPatternVal.pct40;
                case PresetPattern.Percent50:
                    return ST_PresetPatternVal.pct50;
                case PresetPattern.Percent60:
                    return ST_PresetPatternVal.pct60;
                case PresetPattern.Percent70:
                    return ST_PresetPatternVal.pct70;
                case PresetPattern.Percent75:
                    return ST_PresetPatternVal.pct75;
                case PresetPattern.Percent80:
                    return ST_PresetPatternVal.pct80;
                case PresetPattern.Percent90:
                    return ST_PresetPatternVal.pct90;
                case PresetPattern.Plaid:
                    return ST_PresetPatternVal.plaid;
                case PresetPattern.Shingle:
                    return ST_PresetPatternVal.shingle;
                case PresetPattern.SmallCheckerBoard:
                    return ST_PresetPatternVal.smCheck;
                case PresetPattern.SmallConfetti:
                    return ST_PresetPatternVal.smConfetti;
                case PresetPattern.SmallGrid:
                    return ST_PresetPatternVal.smGrid;
                case PresetPattern.SolidDiamond:
                    return ST_PresetPatternVal.solidDmnd;
                case PresetPattern.Sphere:
                    return ST_PresetPatternVal.sphere;
                case PresetPattern.Trellis:
                    return ST_PresetPatternVal.trellis;
                case PresetPattern.UpwardDiagonal:
                    return ST_PresetPatternVal.upDiag;
                case PresetPattern.Vertical:
                    return ST_PresetPatternVal.vert;
                case PresetPattern.Wave:
                    return ST_PresetPatternVal.wave;
                case PresetPattern.Weave:
                    return ST_PresetPatternVal.weave;
                case PresetPattern.WideDownwardDiagonal:
                    return ST_PresetPatternVal.wdDnDiag;
                case PresetPattern.WideUpwardDiagonal:
                    return ST_PresetPatternVal.wdUpDiag;
                case PresetPattern.ZigZag:
                    return ST_PresetPatternVal.zigZag;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}


