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

    public enum SystemColor
    {
        ActiveBorder,
        ActiveCaption,
        ApplicationWorkspace,
        Background,
        ButtonFace,
        ButtonHighlight,
        ButtonShadow,
        ButtonText,
        CaptionText,
        GradientActiveCaption,
        GradientInactiveCaption,
        GrayText,
        Highlight,
        HighlightText,
        HotLight,
        InactiveBorder,
        InactiveCaption,
        InactiveCaptionText,
        InfoBackground,
        InfoText,
        Menu,
        MenuBar,
        MenuHighlight,
        MenuText,
        ScrollBar,
        Window,
        WindowFrame,
        WindowText,
        X3dDarkShadow,
        X3dLight
    }
    public static class SystemColorExtensions
    {
        private static Dictionary<ST_SystemColorVal, SystemColor> reverse = new Dictionary<ST_SystemColorVal, SystemColor>()
        {
            { ST_SystemColorVal.activeBorder, SystemColor.ActiveBorder },
            { ST_SystemColorVal.activeCaption, SystemColor.ActiveCaption },
            { ST_SystemColorVal.appWorkspace, SystemColor.ApplicationWorkspace },
            { ST_SystemColorVal.background, SystemColor.Background },
            { ST_SystemColorVal.btnFace, SystemColor.ButtonFace },
            { ST_SystemColorVal.btnHighlight, SystemColor.ButtonHighlight },
            { ST_SystemColorVal.btnShadow, SystemColor.ButtonShadow },
            { ST_SystemColorVal.btnText, SystemColor.ButtonText },
            { ST_SystemColorVal.captionText, SystemColor.CaptionText },
            { ST_SystemColorVal.gradientActiveCaption, SystemColor.GradientActiveCaption },
            { ST_SystemColorVal.gradientInactiveCaption, SystemColor.GradientInactiveCaption },
            { ST_SystemColorVal.grayText, SystemColor.GrayText },
            { ST_SystemColorVal.highlight, SystemColor.Highlight },
            { ST_SystemColorVal.highlightText, SystemColor.HighlightText },
            { ST_SystemColorVal.hotLight, SystemColor.HotLight },
            { ST_SystemColorVal.inactiveBorder, SystemColor.InactiveBorder },
            { ST_SystemColorVal.inactiveCaption, SystemColor.InactiveCaption },
            { ST_SystemColorVal.inactiveCaptionText, SystemColor.InactiveCaptionText },
            { ST_SystemColorVal.infoBk, SystemColor.InfoBackground },
            { ST_SystemColorVal.infoText, SystemColor.InfoText },
            { ST_SystemColorVal.menu, SystemColor.Menu },
            { ST_SystemColorVal.menuBar, SystemColor.MenuBar },
            { ST_SystemColorVal.menuHighlight, SystemColor.MenuHighlight },
            { ST_SystemColorVal.menuText, SystemColor.MenuText },
            { ST_SystemColorVal.scrollBar, SystemColor.ScrollBar },
            { ST_SystemColorVal.window, SystemColor.Window },
            { ST_SystemColorVal.windowFrame, SystemColor.WindowFrame },
            { ST_SystemColorVal.windowText, SystemColor.WindowText },
            { ST_SystemColorVal.x3dDkShadow, SystemColor.X3dDarkShadow },
            { ST_SystemColorVal.x3dLight, SystemColor.X3dLight },
        };
        public static SystemColor ValueOf(ST_SystemColorVal value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_SystemColorVal", nameof(value));
        }
        public static ST_SystemColorVal ToST_SystemColorVal(this SystemColor value)
        {
            switch(value)
            {
                case SystemColor.ActiveBorder:
                    return ST_SystemColorVal.activeBorder;
                case SystemColor.ActiveCaption:
                    return ST_SystemColorVal.activeCaption;
                case SystemColor.ApplicationWorkspace:
                    return ST_SystemColorVal.appWorkspace;
                case SystemColor.Background:
                    return ST_SystemColorVal.background;
                case SystemColor.ButtonFace:
                    return ST_SystemColorVal.btnFace;
                case SystemColor.ButtonHighlight:
                    return ST_SystemColorVal.btnHighlight;
                case SystemColor.ButtonShadow:
                    return ST_SystemColorVal.btnShadow;
                case SystemColor.ButtonText:
                    return ST_SystemColorVal.btnText;
                case SystemColor.CaptionText:
                    return ST_SystemColorVal.captionText;
                case SystemColor.GradientActiveCaption:
                    return ST_SystemColorVal.gradientActiveCaption;
                case SystemColor.GradientInactiveCaption:
                    return ST_SystemColorVal.gradientInactiveCaption;
                case SystemColor.GrayText:
                    return ST_SystemColorVal.grayText;
                case SystemColor.Highlight:
                    return ST_SystemColorVal.highlight;
                case SystemColor.HighlightText:
                    return ST_SystemColorVal.highlightText;
                case SystemColor.HotLight:
                    return ST_SystemColorVal.hotLight;
                case SystemColor.InactiveBorder:
                    return ST_SystemColorVal.inactiveBorder;
                case SystemColor.InactiveCaption:
                    return ST_SystemColorVal.inactiveCaption;
                case SystemColor.InactiveCaptionText:
                    return ST_SystemColorVal.inactiveCaptionText;
                case SystemColor.InfoBackground:
                    return ST_SystemColorVal.infoBk;
                case SystemColor.InfoText:
                    return ST_SystemColorVal.infoText;
                case SystemColor.Menu:
                    return ST_SystemColorVal.menu;
                case SystemColor.MenuBar:
                    return ST_SystemColorVal.menuBar;
                case SystemColor.MenuHighlight:
                    return ST_SystemColorVal.menuHighlight;
                case SystemColor.MenuText:
                    return ST_SystemColorVal.menuText;
                case SystemColor.ScrollBar:
                    return ST_SystemColorVal.scrollBar;
                case SystemColor.Window:
                    return ST_SystemColorVal.window;
                case SystemColor.WindowFrame:
                    return ST_SystemColorVal.windowFrame;
                case SystemColor.WindowText:
                    return ST_SystemColorVal.windowText;
                case SystemColor.X3dDarkShadow:
                    return ST_SystemColorVal.x3dDkShadow;
                case SystemColor.X3dLight:
                    return ST_SystemColorVal.x3dLight;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}


