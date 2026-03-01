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

    public enum PresetGeometry
    {
        AccentBorderCallout1,
        AccentBorderCallout2,
        AccentBorderCallout3,
        AccentCallout1,
        AccentCallout2,
        AccentCallout3,
        ActionButtonBackPrevious,
        ActionButtonBeginning,
        ActionButtonBlank,
        ActionButtonDocument,
        ActionButtonEnd,
        ActionButtonForwardNext,
        ActionButtonHelp,
        ActionButtonHome,
        ActionButtonInformation,
        ActionButtonMovie,
        ActionButtonReturn,
        ActionButtonSound,
        Arc,
        BentArrow,
        BentConnector2,
        BentConnector3,
        DarkBlue,
        BentConnector4,
        BentConnector5,
        BentUpArrow,
        Bevel,
        BlockArc,
        BorderCallout1,
        BorderCallout2,
        BorderCallout3,
        BracePair,
        BracketPair,
        Callout1,
        Callout2,
        Callout3,
        Can,
        ChartPlus,
        ChartStar,
        ChartX,
        Chevron,
        Chord,
        CircularArrow,
        Cloud,
        CloudCallout,
        Corner,
        CornerTabs,
        Cube,
        CurvedConnector2,
        CurvedConnector3,
        CurvedConnector4,
        CurvedConnector5,
        CurvedDownArrow,
        CurvedLeftArrow,
        CurvedRightArrow,
        CurvedUpArrow,
        Decagon,
        DiagonalStripe,
        Diamond,
        Dodecagon,
        Donut,
        DoubleWave,
        DownArrow,
        DownArrowCallout,
        Ellipse,
        EllipseRibbon,
        EllipseRibbon2,
        FlowChartAlternateProcess,
        FlowChartCollate,
        FlowChartConnector,
        FlowChartDecision,
        FlowChartDelay,
        FlowChartDisplay,
        FlowChartDocument,
        FlowChartExtract,
        FlowChartInputOutput,
        FlowChartInternalStorage,
        FlowChartMagneticDisk,
        FlowChartMagneticDrum,
        FlowChartMagneticTape,
        FlowChartManualInput,
        FlowChartManualOperation,
        FlowChartMerge,
        FlowChartMultidocument,
        FlowChartOfflineStorage,
        FlowChartOffpageConnector,
        FlowChartOnlineStorage,
        FlowChartOr,
        FlowChartPredefinedProcess,
        FlowChartPreparation,
        FlowChartProcess,
        FlowChartPunchedCard,
        FlowChartPunchedTape,
        FlowChartSort,
        FlowChartSummingJunction,
        FlowChartTerminator,
        FoldedCorner,
        Frame,
        Funnel,
        Gear6,
        Gear9,
        HalfFrame,
        Heart,
        Heptagon,
        Hexagon,
        HomePlate,
        HorizontalScroll,
        IrregularSeal1,
        IrregularSeal2,
        LeftArrow,
        LeftArrowCallout,
        LeftBrace,
        LeftBracket,
        LeftCircularArrow,
        LeftRightArrow,
        LeftRightArrowCallout,
        LeftRightCircularArrow,
        LeftRightRibbon,
        LeftRightUpArrow,
        LeftUpArrow,
        LightningBolt,
        Line,
        LineInverted,
        MathDivide,
        MathEqual,
        MathMinus,
        MathMultiply,
        MathNotEqual,
        MathPlus,
        Moon,
        NoSmoking,
        NonIsoscelesTrapezoid,
        NotchedRightArrow,
        Octagon,
        Parallelogram,
        Pentagon,
        Pie,
        PieWedge,
        Plaque,
        PlaqueTabs,
        Plus,
        QuadArrow,
        QuadArrowCallout,
        Rectangle,
        Ribbon,
        Ribbon2,
        RightArrow,
        RightArrowCallout,
        RightBrace,
        RightBracket,
        RoundRectangle1Corner,
        RoundRectangle2DiagonalCorners,
        RoundRectangle2SameSideCorners,
        RoundRectangle,
        RightTriangle,
        SmileyFace,
        SnipRectangle1Corner,
        SnipRectangle2DiagonalCorners,
        SnipRectangle2SameSideCorners,
        SnipRoundRectangle,
        SquareTabs,
        Star10,
        Star12,
        Star16,
        Star24,
        Star32,
        Star4,
        Star5,
        Star6,
        Star7,
        Star8,
        StraightConnector,
        StripedRightArrow,
        Sun,
        SwooshArrow,
        Teardrop,
        Trapezoid,
        Triangle,
        UpArrow,
        UpArrowCallout,
        UpDownArrow,
        UpDownArrowCallout,
        UturnArrow,
        VerticalScroll,
        Wave,
        WedgeEllipseCallout,
        WedgeRectangleCallout,
        WedgeRoundRectangleCallout
    }
    public static class PresetGeometryExtensions
    {
        private static Dictionary<ST_ShapeType, PresetGeometry> reverse = new Dictionary<ST_ShapeType, PresetGeometry>()
        {
            { ST_ShapeType.accentBorderCallout1, PresetGeometry.AccentBorderCallout1 },
            { ST_ShapeType.accentBorderCallout2, PresetGeometry.AccentBorderCallout2 },
            { ST_ShapeType.accentBorderCallout3, PresetGeometry.AccentBorderCallout3 },
            { ST_ShapeType.accentCallout1, PresetGeometry.AccentCallout1 },
            { ST_ShapeType.accentCallout2, PresetGeometry.AccentCallout2 },
            { ST_ShapeType.accentCallout3, PresetGeometry.AccentCallout3 },
            { ST_ShapeType.actionButtonBackPrevious, PresetGeometry.ActionButtonBackPrevious },
            { ST_ShapeType.actionButtonBeginning, PresetGeometry.ActionButtonBeginning },
            { ST_ShapeType.actionButtonBlank, PresetGeometry.ActionButtonBlank },
            { ST_ShapeType.actionButtonDocument, PresetGeometry.ActionButtonDocument },
            { ST_ShapeType.actionButtonEnd, PresetGeometry.ActionButtonEnd },
            { ST_ShapeType.actionButtonForwardNext, PresetGeometry.ActionButtonForwardNext },
            { ST_ShapeType.actionButtonHelp, PresetGeometry.ActionButtonHelp },
            { ST_ShapeType.actionButtonHome, PresetGeometry.ActionButtonHome },
            { ST_ShapeType.actionButtonInformation, PresetGeometry.ActionButtonInformation },
            { ST_ShapeType.actionButtonMovie, PresetGeometry.ActionButtonMovie },
            { ST_ShapeType.actionButtonReturn, PresetGeometry.ActionButtonReturn },
            { ST_ShapeType.actionButtonSound, PresetGeometry.ActionButtonSound },
            { ST_ShapeType.arc, PresetGeometry.Arc },
            { ST_ShapeType.bentArrow, PresetGeometry.BentArrow },
            { ST_ShapeType.bentConnector2, PresetGeometry.BentConnector2 },
            { ST_ShapeType.bentConnector3, PresetGeometry.BentConnector3 },
            { ST_ShapeType.bentConnector4, PresetGeometry.DarkBlue },
            { ST_ShapeType.bentConnector4, PresetGeometry.BentConnector4 },
            { ST_ShapeType.bentConnector5, PresetGeometry.BentConnector5 },
            { ST_ShapeType.bentUpArrow, PresetGeometry.BentUpArrow },
            { ST_ShapeType.bevel, PresetGeometry.Bevel },
            { ST_ShapeType.blockArc, PresetGeometry.BlockArc },
            { ST_ShapeType.borderCallout1, PresetGeometry.BorderCallout1 },
            { ST_ShapeType.borderCallout2, PresetGeometry.BorderCallout2 },
            { ST_ShapeType.borderCallout3, PresetGeometry.BorderCallout3 },
            { ST_ShapeType.bracePair, PresetGeometry.BracePair },
            { ST_ShapeType.bracketPair, PresetGeometry.BracketPair },
            { ST_ShapeType.callout1, PresetGeometry.Callout1 },
            { ST_ShapeType.callout2, PresetGeometry.Callout2 },
            { ST_ShapeType.callout3, PresetGeometry.Callout3 },
            { ST_ShapeType.can, PresetGeometry.Can },
            { ST_ShapeType.chartPlus, PresetGeometry.ChartPlus },
            { ST_ShapeType.chartStar, PresetGeometry.ChartStar },
            { ST_ShapeType.chartX, PresetGeometry.ChartX },
            { ST_ShapeType.chevron, PresetGeometry.Chevron },
            { ST_ShapeType.chord, PresetGeometry.Chord },
            { ST_ShapeType.circularArrow, PresetGeometry.CircularArrow },
            { ST_ShapeType.cloud, PresetGeometry.Cloud },
            { ST_ShapeType.cloudCallout, PresetGeometry.CloudCallout },
            { ST_ShapeType.corner, PresetGeometry.Corner },
            { ST_ShapeType.cornerTabs, PresetGeometry.CornerTabs },
            { ST_ShapeType.cube, PresetGeometry.Cube },
            { ST_ShapeType.curvedConnector2, PresetGeometry.CurvedConnector2 },
            { ST_ShapeType.curvedConnector3, PresetGeometry.CurvedConnector3 },
            { ST_ShapeType.curvedConnector4, PresetGeometry.CurvedConnector4 },
            { ST_ShapeType.curvedConnector5, PresetGeometry.CurvedConnector5 },
            { ST_ShapeType.curvedDownArrow, PresetGeometry.CurvedDownArrow },
            { ST_ShapeType.curvedLeftArrow, PresetGeometry.CurvedLeftArrow },
            { ST_ShapeType.curvedRightArrow, PresetGeometry.CurvedRightArrow },
            { ST_ShapeType.curvedUpArrow, PresetGeometry.CurvedUpArrow },
            { ST_ShapeType.decagon, PresetGeometry.Decagon },
            { ST_ShapeType.diagStripe, PresetGeometry.DiagonalStripe },
            { ST_ShapeType.diamond, PresetGeometry.Diamond },
            { ST_ShapeType.dodecagon, PresetGeometry.Dodecagon },
            { ST_ShapeType.donut, PresetGeometry.Donut },
            { ST_ShapeType.doubleWave, PresetGeometry.DoubleWave },
            { ST_ShapeType.downArrow, PresetGeometry.DownArrow },
            { ST_ShapeType.downArrowCallout, PresetGeometry.DownArrowCallout },
            { ST_ShapeType.ellipse, PresetGeometry.Ellipse },
            { ST_ShapeType.ellipseRibbon, PresetGeometry.EllipseRibbon },
            { ST_ShapeType.ellipseRibbon2, PresetGeometry.EllipseRibbon2 },
            { ST_ShapeType.flowChartAlternateProcess, PresetGeometry.FlowChartAlternateProcess },
            { ST_ShapeType.flowChartCollate, PresetGeometry.FlowChartCollate },
            { ST_ShapeType.flowChartConnector, PresetGeometry.FlowChartConnector },
            { ST_ShapeType.flowChartDecision, PresetGeometry.FlowChartDecision },
            { ST_ShapeType.flowChartDelay, PresetGeometry.FlowChartDelay },
            { ST_ShapeType.flowChartDisplay, PresetGeometry.FlowChartDisplay },
            { ST_ShapeType.flowChartDocument, PresetGeometry.FlowChartDocument },
            { ST_ShapeType.flowChartExtract, PresetGeometry.FlowChartExtract },
            { ST_ShapeType.flowChartInputOutput, PresetGeometry.FlowChartInputOutput },
            { ST_ShapeType.flowChartInternalStorage, PresetGeometry.FlowChartInternalStorage },
            { ST_ShapeType.flowChartMagneticDisk, PresetGeometry.FlowChartMagneticDisk },
            { ST_ShapeType.flowChartMagneticDrum, PresetGeometry.FlowChartMagneticDrum },
            { ST_ShapeType.flowChartMagneticTape, PresetGeometry.FlowChartMagneticTape },
            { ST_ShapeType.flowChartManualInput, PresetGeometry.FlowChartManualInput },
            { ST_ShapeType.flowChartManualOperation, PresetGeometry.FlowChartManualOperation },
            { ST_ShapeType.flowChartMerge, PresetGeometry.FlowChartMerge },
            { ST_ShapeType.flowChartMultidocument, PresetGeometry.FlowChartMultidocument },
            { ST_ShapeType.flowChartOfflineStorage, PresetGeometry.FlowChartOfflineStorage },
            { ST_ShapeType.flowChartOffpageConnector, PresetGeometry.FlowChartOffpageConnector },
            { ST_ShapeType.flowChartOnlineStorage, PresetGeometry.FlowChartOnlineStorage },
            { ST_ShapeType.flowChartOr, PresetGeometry.FlowChartOr },
            { ST_ShapeType.flowChartPredefinedProcess, PresetGeometry.FlowChartPredefinedProcess },
            { ST_ShapeType.flowChartPreparation, PresetGeometry.FlowChartPreparation },
            { ST_ShapeType.flowChartProcess, PresetGeometry.FlowChartProcess },
            { ST_ShapeType.flowChartPunchedCard, PresetGeometry.FlowChartPunchedCard },
            { ST_ShapeType.flowChartPunchedTape, PresetGeometry.FlowChartPunchedTape },
            { ST_ShapeType.flowChartSort, PresetGeometry.FlowChartSort },
            { ST_ShapeType.flowChartSummingJunction, PresetGeometry.FlowChartSummingJunction },
            { ST_ShapeType.flowChartTerminator, PresetGeometry.FlowChartTerminator },
            { ST_ShapeType.foldedCorner, PresetGeometry.FoldedCorner },
            { ST_ShapeType.frame, PresetGeometry.Frame },
            { ST_ShapeType.funnel, PresetGeometry.Funnel },
            { ST_ShapeType.gear6, PresetGeometry.Gear6 },
            { ST_ShapeType.gear9, PresetGeometry.Gear9 },
            { ST_ShapeType.halfFrame, PresetGeometry.HalfFrame },
            { ST_ShapeType.heart, PresetGeometry.Heart },
            { ST_ShapeType.heptagon, PresetGeometry.Heptagon },
            { ST_ShapeType.hexagon, PresetGeometry.Hexagon },
            { ST_ShapeType.homePlate, PresetGeometry.HomePlate },
            { ST_ShapeType.horizontalScroll, PresetGeometry.HorizontalScroll },
            { ST_ShapeType.irregularSeal1, PresetGeometry.IrregularSeal1 },
            { ST_ShapeType.irregularSeal2, PresetGeometry.IrregularSeal2 },
            { ST_ShapeType.leftArrow, PresetGeometry.LeftArrow },
            { ST_ShapeType.leftArrowCallout, PresetGeometry.LeftArrowCallout },
            { ST_ShapeType.leftBrace, PresetGeometry.LeftBrace },
            { ST_ShapeType.leftBracket, PresetGeometry.LeftBracket },
            { ST_ShapeType.leftCircularArrow, PresetGeometry.LeftCircularArrow },
            { ST_ShapeType.leftRightArrow, PresetGeometry.LeftRightArrow },
            { ST_ShapeType.leftRightArrowCallout, PresetGeometry.LeftRightArrowCallout },
            { ST_ShapeType.leftRightCircularArrow, PresetGeometry.LeftRightCircularArrow },
            { ST_ShapeType.leftRightRibbon, PresetGeometry.LeftRightRibbon },
            { ST_ShapeType.leftRightUpArrow, PresetGeometry.LeftRightUpArrow },
            { ST_ShapeType.leftUpArrow, PresetGeometry.LeftUpArrow },
            { ST_ShapeType.lightningBolt, PresetGeometry.LightningBolt },
            { ST_ShapeType.line, PresetGeometry.Line },
            { ST_ShapeType.lineInv, PresetGeometry.LineInverted },
            { ST_ShapeType.mathDivide, PresetGeometry.MathDivide },
            { ST_ShapeType.mathEqual, PresetGeometry.MathEqual },
            { ST_ShapeType.mathMinus, PresetGeometry.MathMinus },
            { ST_ShapeType.mathMultiply, PresetGeometry.MathMultiply },
            { ST_ShapeType.mathNotEqual, PresetGeometry.MathNotEqual },
            { ST_ShapeType.mathPlus, PresetGeometry.MathPlus },
            { ST_ShapeType.moon, PresetGeometry.Moon },
            { ST_ShapeType.noSmoking, PresetGeometry.NoSmoking },
            { ST_ShapeType.nonIsoscelesTrapezoid, PresetGeometry.NonIsoscelesTrapezoid },
            { ST_ShapeType.notchedRightArrow, PresetGeometry.NotchedRightArrow },
            { ST_ShapeType.octagon, PresetGeometry.Octagon },
            { ST_ShapeType.parallelogram, PresetGeometry.Parallelogram },
            { ST_ShapeType.pentagon, PresetGeometry.Pentagon },
            { ST_ShapeType.pie, PresetGeometry.Pie },
            { ST_ShapeType.pieWedge, PresetGeometry.PieWedge },
            { ST_ShapeType.plaque, PresetGeometry.Plaque },
            { ST_ShapeType.plaqueTabs, PresetGeometry.PlaqueTabs },
            { ST_ShapeType.plus, PresetGeometry.Plus },
            { ST_ShapeType.quadArrow, PresetGeometry.QuadArrow },
            { ST_ShapeType.quadArrowCallout, PresetGeometry.QuadArrowCallout },
            { ST_ShapeType.rect, PresetGeometry.Rectangle },
            { ST_ShapeType.ribbon, PresetGeometry.Ribbon },
            { ST_ShapeType.ribbon2, PresetGeometry.Ribbon2 },
            { ST_ShapeType.rightArrow, PresetGeometry.RightArrow },
            { ST_ShapeType.rightArrowCallout, PresetGeometry.RightArrowCallout },
            { ST_ShapeType.rightBrace, PresetGeometry.RightBrace },
            { ST_ShapeType.rightBracket, PresetGeometry.RightBracket },
            { ST_ShapeType.round1Rect, PresetGeometry.RoundRectangle1Corner },
            { ST_ShapeType.round2DiagRect, PresetGeometry.RoundRectangle2DiagonalCorners },
            { ST_ShapeType.round2SameRect, PresetGeometry.RoundRectangle2SameSideCorners },
            { ST_ShapeType.roundRect, PresetGeometry.RoundRectangle },
            { ST_ShapeType.rtTriangle, PresetGeometry.RightTriangle },
            { ST_ShapeType.smileyFace, PresetGeometry.SmileyFace },
            { ST_ShapeType.snip1Rect, PresetGeometry.SnipRectangle1Corner },
            { ST_ShapeType.snip2DiagRect, PresetGeometry.SnipRectangle2DiagonalCorners },
            { ST_ShapeType.snip2SameRect, PresetGeometry.SnipRectangle2SameSideCorners },
            { ST_ShapeType.snipRoundRect, PresetGeometry.SnipRoundRectangle },
            { ST_ShapeType.squareTabs, PresetGeometry.SquareTabs },
            { ST_ShapeType.star10, PresetGeometry.Star10 },
            { ST_ShapeType.star12, PresetGeometry.Star12 },
            { ST_ShapeType.star16, PresetGeometry.Star16 },
            { ST_ShapeType.star24, PresetGeometry.Star24 },
            { ST_ShapeType.star32, PresetGeometry.Star32 },
            { ST_ShapeType.star4, PresetGeometry.Star4 },
            { ST_ShapeType.star5, PresetGeometry.Star5 },
            { ST_ShapeType.star6, PresetGeometry.Star6 },
            { ST_ShapeType.star7, PresetGeometry.Star7 },
            { ST_ShapeType.star8, PresetGeometry.Star8 },
            { ST_ShapeType.straightConnector1, PresetGeometry.StraightConnector },
            { ST_ShapeType.stripedRightArrow, PresetGeometry.StripedRightArrow },
            { ST_ShapeType.sun, PresetGeometry.Sun },
            { ST_ShapeType.swooshArrow, PresetGeometry.SwooshArrow },
            { ST_ShapeType.teardrop, PresetGeometry.Teardrop },
            { ST_ShapeType.trapezoid, PresetGeometry.Trapezoid },
            { ST_ShapeType.triangle, PresetGeometry.Triangle },
            { ST_ShapeType.upArrow, PresetGeometry.UpArrow },
            { ST_ShapeType.upArrowCallout, PresetGeometry.UpArrowCallout },
            { ST_ShapeType.upDownArrow, PresetGeometry.UpDownArrow },
            { ST_ShapeType.upDownArrowCallout, PresetGeometry.UpDownArrowCallout },
            { ST_ShapeType.uturnArrow, PresetGeometry.UturnArrow },
            { ST_ShapeType.verticalScroll, PresetGeometry.VerticalScroll },
            { ST_ShapeType.wave, PresetGeometry.Wave },
            { ST_ShapeType.wedgeEllipseCallout, PresetGeometry.WedgeEllipseCallout },
            { ST_ShapeType.wedgeRectCallout, PresetGeometry.WedgeRectangleCallout },
            { ST_ShapeType.wedgeRoundRectCallout, PresetGeometry.WedgeRoundRectangleCallout },
        };
        public static PresetGeometry ValueOf(ST_ShapeType value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_ShapeType", nameof(value));
        }
        public static ST_ShapeType ToST_ShapeType(this PresetGeometry value)
        {
            switch(value)
            {
                case PresetGeometry.AccentBorderCallout1:
                    return ST_ShapeType.accentBorderCallout1;
                case PresetGeometry.AccentBorderCallout2:
                    return ST_ShapeType.accentBorderCallout2;
                case PresetGeometry.AccentBorderCallout3:
                    return ST_ShapeType.accentBorderCallout3;
                case PresetGeometry.AccentCallout1:
                    return ST_ShapeType.accentCallout1;
                case PresetGeometry.AccentCallout2:
                    return ST_ShapeType.accentCallout2;
                case PresetGeometry.AccentCallout3:
                    return ST_ShapeType.accentCallout3;
                case PresetGeometry.ActionButtonBackPrevious:
                    return ST_ShapeType.actionButtonBackPrevious;
                case PresetGeometry.ActionButtonBeginning:
                    return ST_ShapeType.actionButtonBeginning;
                case PresetGeometry.ActionButtonBlank:
                    return ST_ShapeType.actionButtonBlank;
                case PresetGeometry.ActionButtonDocument:
                    return ST_ShapeType.actionButtonDocument;
                case PresetGeometry.ActionButtonEnd:
                    return ST_ShapeType.actionButtonEnd;
                case PresetGeometry.ActionButtonForwardNext:
                    return ST_ShapeType.actionButtonForwardNext;
                case PresetGeometry.ActionButtonHelp:
                    return ST_ShapeType.actionButtonHelp;
                case PresetGeometry.ActionButtonHome:
                    return ST_ShapeType.actionButtonHome;
                case PresetGeometry.ActionButtonInformation:
                    return ST_ShapeType.actionButtonInformation;
                case PresetGeometry.ActionButtonMovie:
                    return ST_ShapeType.actionButtonMovie;
                case PresetGeometry.ActionButtonReturn:
                    return ST_ShapeType.actionButtonReturn;
                case PresetGeometry.ActionButtonSound:
                    return ST_ShapeType.actionButtonSound;
                case PresetGeometry.Arc:
                    return ST_ShapeType.arc;
                case PresetGeometry.BentArrow:
                    return ST_ShapeType.bentArrow;
                case PresetGeometry.BentConnector2:
                    return ST_ShapeType.bentConnector2;
                case PresetGeometry.BentConnector3:
                    return ST_ShapeType.bentConnector3;
                case PresetGeometry.DarkBlue:
                    return ST_ShapeType.bentConnector4;
                case PresetGeometry.BentConnector4:
                    return ST_ShapeType.bentConnector4;
                case PresetGeometry.BentConnector5:
                    return ST_ShapeType.bentConnector5;
                case PresetGeometry.BentUpArrow:
                    return ST_ShapeType.bentUpArrow;
                case PresetGeometry.Bevel:
                    return ST_ShapeType.bevel;
                case PresetGeometry.BlockArc:
                    return ST_ShapeType.blockArc;
                case PresetGeometry.BorderCallout1:
                    return ST_ShapeType.borderCallout1;
                case PresetGeometry.BorderCallout2:
                    return ST_ShapeType.borderCallout2;
                case PresetGeometry.BorderCallout3:
                    return ST_ShapeType.borderCallout3;
                case PresetGeometry.BracePair:
                    return ST_ShapeType.bracePair;
                case PresetGeometry.BracketPair:
                    return ST_ShapeType.bracketPair;
                case PresetGeometry.Callout1:
                    return ST_ShapeType.callout1;
                case PresetGeometry.Callout2:
                    return ST_ShapeType.callout2;
                case PresetGeometry.Callout3:
                    return ST_ShapeType.callout3;
                case PresetGeometry.Can:
                    return ST_ShapeType.can;
                case PresetGeometry.ChartPlus:
                    return ST_ShapeType.chartPlus;
                case PresetGeometry.ChartStar:
                    return ST_ShapeType.chartStar;
                case PresetGeometry.ChartX:
                    return ST_ShapeType.chartX;
                case PresetGeometry.Chevron:
                    return ST_ShapeType.chevron;
                case PresetGeometry.Chord:
                    return ST_ShapeType.chord;
                case PresetGeometry.CircularArrow:
                    return ST_ShapeType.circularArrow;
                case PresetGeometry.Cloud:
                    return ST_ShapeType.cloud;
                case PresetGeometry.CloudCallout:
                    return ST_ShapeType.cloudCallout;
                case PresetGeometry.Corner:
                    return ST_ShapeType.corner;
                case PresetGeometry.CornerTabs:
                    return ST_ShapeType.cornerTabs;
                case PresetGeometry.Cube:
                    return ST_ShapeType.cube;
                case PresetGeometry.CurvedConnector2:
                    return ST_ShapeType.curvedConnector2;
                case PresetGeometry.CurvedConnector3:
                    return ST_ShapeType.curvedConnector3;
                case PresetGeometry.CurvedConnector4:
                    return ST_ShapeType.curvedConnector4;
                case PresetGeometry.CurvedConnector5:
                    return ST_ShapeType.curvedConnector5;
                case PresetGeometry.CurvedDownArrow:
                    return ST_ShapeType.curvedDownArrow;
                case PresetGeometry.CurvedLeftArrow:
                    return ST_ShapeType.curvedLeftArrow;
                case PresetGeometry.CurvedRightArrow:
                    return ST_ShapeType.curvedRightArrow;
                case PresetGeometry.CurvedUpArrow:
                    return ST_ShapeType.curvedUpArrow;
                case PresetGeometry.Decagon:
                    return ST_ShapeType.decagon;
                case PresetGeometry.DiagonalStripe:
                    return ST_ShapeType.diagStripe;
                case PresetGeometry.Diamond:
                    return ST_ShapeType.diamond;
                case PresetGeometry.Dodecagon:
                    return ST_ShapeType.dodecagon;
                case PresetGeometry.Donut:
                    return ST_ShapeType.donut;
                case PresetGeometry.DoubleWave:
                    return ST_ShapeType.doubleWave;
                case PresetGeometry.DownArrow:
                    return ST_ShapeType.downArrow;
                case PresetGeometry.DownArrowCallout:
                    return ST_ShapeType.downArrowCallout;
                case PresetGeometry.Ellipse:
                    return ST_ShapeType.ellipse;
                case PresetGeometry.EllipseRibbon:
                    return ST_ShapeType.ellipseRibbon;
                case PresetGeometry.EllipseRibbon2:
                    return ST_ShapeType.ellipseRibbon2;
                case PresetGeometry.FlowChartAlternateProcess:
                    return ST_ShapeType.flowChartAlternateProcess;
                case PresetGeometry.FlowChartCollate:
                    return ST_ShapeType.flowChartCollate;
                case PresetGeometry.FlowChartConnector:
                    return ST_ShapeType.flowChartConnector;
                case PresetGeometry.FlowChartDecision:
                    return ST_ShapeType.flowChartDecision;
                case PresetGeometry.FlowChartDelay:
                    return ST_ShapeType.flowChartDelay;
                case PresetGeometry.FlowChartDisplay:
                    return ST_ShapeType.flowChartDisplay;
                case PresetGeometry.FlowChartDocument:
                    return ST_ShapeType.flowChartDocument;
                case PresetGeometry.FlowChartExtract:
                    return ST_ShapeType.flowChartExtract;
                case PresetGeometry.FlowChartInputOutput:
                    return ST_ShapeType.flowChartInputOutput;
                case PresetGeometry.FlowChartInternalStorage:
                    return ST_ShapeType.flowChartInternalStorage;
                case PresetGeometry.FlowChartMagneticDisk:
                    return ST_ShapeType.flowChartMagneticDisk;
                case PresetGeometry.FlowChartMagneticDrum:
                    return ST_ShapeType.flowChartMagneticDrum;
                case PresetGeometry.FlowChartMagneticTape:
                    return ST_ShapeType.flowChartMagneticTape;
                case PresetGeometry.FlowChartManualInput:
                    return ST_ShapeType.flowChartManualInput;
                case PresetGeometry.FlowChartManualOperation:
                    return ST_ShapeType.flowChartManualOperation;
                case PresetGeometry.FlowChartMerge:
                    return ST_ShapeType.flowChartMerge;
                case PresetGeometry.FlowChartMultidocument:
                    return ST_ShapeType.flowChartMultidocument;
                case PresetGeometry.FlowChartOfflineStorage:
                    return ST_ShapeType.flowChartOfflineStorage;
                case PresetGeometry.FlowChartOffpageConnector:
                    return ST_ShapeType.flowChartOffpageConnector;
                case PresetGeometry.FlowChartOnlineStorage:
                    return ST_ShapeType.flowChartOnlineStorage;
                case PresetGeometry.FlowChartOr:
                    return ST_ShapeType.flowChartOr;
                case PresetGeometry.FlowChartPredefinedProcess:
                    return ST_ShapeType.flowChartPredefinedProcess;
                case PresetGeometry.FlowChartPreparation:
                    return ST_ShapeType.flowChartPreparation;
                case PresetGeometry.FlowChartProcess:
                    return ST_ShapeType.flowChartProcess;
                case PresetGeometry.FlowChartPunchedCard:
                    return ST_ShapeType.flowChartPunchedCard;
                case PresetGeometry.FlowChartPunchedTape:
                    return ST_ShapeType.flowChartPunchedTape;
                case PresetGeometry.FlowChartSort:
                    return ST_ShapeType.flowChartSort;
                case PresetGeometry.FlowChartSummingJunction:
                    return ST_ShapeType.flowChartSummingJunction;
                case PresetGeometry.FlowChartTerminator:
                    return ST_ShapeType.flowChartTerminator;
                case PresetGeometry.FoldedCorner:
                    return ST_ShapeType.foldedCorner;
                case PresetGeometry.Frame:
                    return ST_ShapeType.frame;
                case PresetGeometry.Funnel:
                    return ST_ShapeType.funnel;
                case PresetGeometry.Gear6:
                    return ST_ShapeType.gear6;
                case PresetGeometry.Gear9:
                    return ST_ShapeType.gear9;
                case PresetGeometry.HalfFrame:
                    return ST_ShapeType.halfFrame;
                case PresetGeometry.Heart:
                    return ST_ShapeType.heart;
                case PresetGeometry.Heptagon:
                    return ST_ShapeType.heptagon;
                case PresetGeometry.Hexagon:
                    return ST_ShapeType.hexagon;
                case PresetGeometry.HomePlate:
                    return ST_ShapeType.homePlate;
                case PresetGeometry.HorizontalScroll:
                    return ST_ShapeType.horizontalScroll;
                case PresetGeometry.IrregularSeal1:
                    return ST_ShapeType.irregularSeal1;
                case PresetGeometry.IrregularSeal2:
                    return ST_ShapeType.irregularSeal2;
                case PresetGeometry.LeftArrow:
                    return ST_ShapeType.leftArrow;
                case PresetGeometry.LeftArrowCallout:
                    return ST_ShapeType.leftArrowCallout;
                case PresetGeometry.LeftBrace:
                    return ST_ShapeType.leftBrace;
                case PresetGeometry.LeftBracket:
                    return ST_ShapeType.leftBracket;
                case PresetGeometry.LeftCircularArrow:
                    return ST_ShapeType.leftCircularArrow;
                case PresetGeometry.LeftRightArrow:
                    return ST_ShapeType.leftRightArrow;
                case PresetGeometry.LeftRightArrowCallout:
                    return ST_ShapeType.leftRightArrowCallout;
                case PresetGeometry.LeftRightCircularArrow:
                    return ST_ShapeType.leftRightCircularArrow;
                case PresetGeometry.LeftRightRibbon:
                    return ST_ShapeType.leftRightRibbon;
                case PresetGeometry.LeftRightUpArrow:
                    return ST_ShapeType.leftRightUpArrow;
                case PresetGeometry.LeftUpArrow:
                    return ST_ShapeType.leftUpArrow;
                case PresetGeometry.LightningBolt:
                    return ST_ShapeType.lightningBolt;
                case PresetGeometry.Line:
                    return ST_ShapeType.line;
                case PresetGeometry.LineInverted:
                    return ST_ShapeType.lineInv;
                case PresetGeometry.MathDivide:
                    return ST_ShapeType.mathDivide;
                case PresetGeometry.MathEqual:
                    return ST_ShapeType.mathEqual;
                case PresetGeometry.MathMinus:
                    return ST_ShapeType.mathMinus;
                case PresetGeometry.MathMultiply:
                    return ST_ShapeType.mathMultiply;
                case PresetGeometry.MathNotEqual:
                    return ST_ShapeType.mathNotEqual;
                case PresetGeometry.MathPlus:
                    return ST_ShapeType.mathPlus;
                case PresetGeometry.Moon:
                    return ST_ShapeType.moon;
                case PresetGeometry.NoSmoking:
                    return ST_ShapeType.noSmoking;
                case PresetGeometry.NonIsoscelesTrapezoid:
                    return ST_ShapeType.nonIsoscelesTrapezoid;
                case PresetGeometry.NotchedRightArrow:
                    return ST_ShapeType.notchedRightArrow;
                case PresetGeometry.Octagon:
                    return ST_ShapeType.octagon;
                case PresetGeometry.Parallelogram:
                    return ST_ShapeType.parallelogram;
                case PresetGeometry.Pentagon:
                    return ST_ShapeType.pentagon;
                case PresetGeometry.Pie:
                    return ST_ShapeType.pie;
                case PresetGeometry.PieWedge:
                    return ST_ShapeType.pieWedge;
                case PresetGeometry.Plaque:
                    return ST_ShapeType.plaque;
                case PresetGeometry.PlaqueTabs:
                    return ST_ShapeType.plaqueTabs;
                case PresetGeometry.Plus:
                    return ST_ShapeType.plus;
                case PresetGeometry.QuadArrow:
                    return ST_ShapeType.quadArrow;
                case PresetGeometry.QuadArrowCallout:
                    return ST_ShapeType.quadArrowCallout;
                case PresetGeometry.Rectangle:
                    return ST_ShapeType.rect;
                case PresetGeometry.Ribbon:
                    return ST_ShapeType.ribbon;
                case PresetGeometry.Ribbon2:
                    return ST_ShapeType.ribbon2;
                case PresetGeometry.RightArrow:
                    return ST_ShapeType.rightArrow;
                case PresetGeometry.RightArrowCallout:
                    return ST_ShapeType.rightArrowCallout;
                case PresetGeometry.RightBrace:
                    return ST_ShapeType.rightBrace;
                case PresetGeometry.RightBracket:
                    return ST_ShapeType.rightBracket;
                case PresetGeometry.RoundRectangle1Corner:
                    return ST_ShapeType.round1Rect;
                case PresetGeometry.RoundRectangle2DiagonalCorners:
                    return ST_ShapeType.round2DiagRect;
                case PresetGeometry.RoundRectangle2SameSideCorners:
                    return ST_ShapeType.round2SameRect;
                case PresetGeometry.RoundRectangle:
                    return ST_ShapeType.roundRect;
                case PresetGeometry.RightTriangle:
                    return ST_ShapeType.rtTriangle;
                case PresetGeometry.SmileyFace:
                    return ST_ShapeType.smileyFace;
                case PresetGeometry.SnipRectangle1Corner:
                    return ST_ShapeType.snip1Rect;
                case PresetGeometry.SnipRectangle2DiagonalCorners:
                    return ST_ShapeType.snip2DiagRect;
                case PresetGeometry.SnipRectangle2SameSideCorners:
                    return ST_ShapeType.snip2SameRect;
                case PresetGeometry.SnipRoundRectangle:
                    return ST_ShapeType.snipRoundRect;
                case PresetGeometry.SquareTabs:
                    return ST_ShapeType.squareTabs;
                case PresetGeometry.Star10:
                    return ST_ShapeType.star10;
                case PresetGeometry.Star12:
                    return ST_ShapeType.star12;
                case PresetGeometry.Star16:
                    return ST_ShapeType.star16;
                case PresetGeometry.Star24:
                    return ST_ShapeType.star24;
                case PresetGeometry.Star32:
                    return ST_ShapeType.star32;
                case PresetGeometry.Star4:
                    return ST_ShapeType.star4;
                case PresetGeometry.Star5:
                    return ST_ShapeType.star5;
                case PresetGeometry.Star6:
                    return ST_ShapeType.star6;
                case PresetGeometry.Star7:
                    return ST_ShapeType.star7;
                case PresetGeometry.Star8:
                    return ST_ShapeType.star8;
                case PresetGeometry.StraightConnector:
                    return ST_ShapeType.straightConnector1;
                case PresetGeometry.StripedRightArrow:
                    return ST_ShapeType.stripedRightArrow;
                case PresetGeometry.Sun:
                    return ST_ShapeType.sun;
                case PresetGeometry.SwooshArrow:
                    return ST_ShapeType.swooshArrow;
                case PresetGeometry.Teardrop:
                    return ST_ShapeType.teardrop;
                case PresetGeometry.Trapezoid:
                    return ST_ShapeType.trapezoid;
                case PresetGeometry.Triangle:
                    return ST_ShapeType.triangle;
                case PresetGeometry.UpArrow:
                    return ST_ShapeType.upArrow;
                case PresetGeometry.UpArrowCallout:
                    return ST_ShapeType.upArrowCallout;
                case PresetGeometry.UpDownArrow:
                    return ST_ShapeType.upDownArrow;
                case PresetGeometry.UpDownArrowCallout:
                    return ST_ShapeType.upDownArrowCallout;
                case PresetGeometry.UturnArrow:
                    return ST_ShapeType.uturnArrow;
                case PresetGeometry.VerticalScroll:
                    return ST_ShapeType.verticalScroll;
                case PresetGeometry.Wave:
                    return ST_ShapeType.wave;
                case PresetGeometry.WedgeEllipseCallout:
                    return ST_ShapeType.wedgeEllipseCallout;
                case PresetGeometry.WedgeRectangleCallout:
                    return ST_ShapeType.wedgeRectCallout;
                case PresetGeometry.WedgeRoundRectangleCallout:
                    return ST_ShapeType.wedgeRoundRectCallout;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}


