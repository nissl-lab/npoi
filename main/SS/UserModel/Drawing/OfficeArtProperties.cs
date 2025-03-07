using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.UserModel.Drawing
{
    public static class OfficeArtProperties
    {
        #region FillStyle
        public const int fillType = 0x180;
        public const int fillColor = 0x181;
        public const int fillOpacity = 0x182;
        public const int fillBackColor = 0x183;
        public const int fillBackOpacity = 0x184;
        public const int fillCrMod = 0x185;
        public const int fillBlip = 0x186;
        public const int fillBlipName = 0x187;
        public const int fillBlipFlags = 0x188;
        public const int fillWidth = 0x189;
        public const int fillHeight = 0x18a;
        public const int fillAngle = 0x18b;
        public const int fillFocus = 0x18c;
        public const int fillToLeft = 0x18d;
        public const int fillToTop = 0x18e;
        public const int fillToRight = 0x18f;
        public const int fillToBottom = 0x190;
        public const int fillRectLeft = 0x191;
        public const int fillRectTop = 0x192;
        public const int fillRectRight = 0x193;
        public const int fillRectBottom = 0x194;
        public const int fillDztype = 0x195;
        public const int fillShadePreset = 0x196;
        public const int fillShadeColors = 0x197;
        public const int fillOriginX = 0x198;
        public const int fillOriginY = 0x199;
        public const int fillShapeOriginX = 0x19a;
        public const int fillShapeOriginY = 0x19b;
        public const int fillShadeType = 0x19c;

        public const int fillColorExt = 0x19e;
        public const int reserved415 = 0x19f;
        public const int fillColorExtMod = 0x1a0;
        public const int reserved417 = 0x1a1;
        public const int fillBackColorExt = 0x1a2;
        public const int reserved419 = 0x1a3;
        public const int fillBackColorExtMod = 0x1a4;
        public const int reserved421 = 0x1a5;
        public const int reserved422 = 0x1a6;
        public const int reserved423 = 0x1a7;
        public const int fillStyleBoolean = 0x1bf;

        private static readonly Dictionary<int, string> fillStyle = new Dictionary<int,string>();
        private static void InitFillStyle()
        {
            fillStyle.Add(0x180, "fillType");
            fillStyle.Add(0x181, "fillColor");
            fillStyle.Add(0x182, "fillOpacity");
            fillStyle.Add(0x183, "fillBackColor");
            fillStyle.Add(0x184, "fillBackOpacity");
            fillStyle.Add(0x185, "fillCrMod");
            fillStyle.Add(0x186, "fillBlip");
            fillStyle.Add(0x187, "fillBlipName");
            fillStyle.Add(0x188, "fillBlipFlags");
            fillStyle.Add(0x189, "fillWidth");
            fillStyle.Add(0x18a, "fillHeight");
            fillStyle.Add(0x18b, "fillAngle");
            fillStyle.Add(0x18c, "fillFocus");
            fillStyle.Add(0x18d, "fillToLeft");
            fillStyle.Add(0x18e, "fillToTop");
            fillStyle.Add(0x18f, "fillToRight");
            fillStyle.Add(0x190, "fillToBottom");
            fillStyle.Add(0x191, "fillRectLeft");
            fillStyle.Add(0x192, "fillRectTop");
            fillStyle.Add(0x193, "fillRectRight");
            fillStyle.Add(0x194, "fillRectBottom");
            fillStyle.Add(0x195, "fillDztype");
            fillStyle.Add(0x196, "fillShadePreset");
            fillStyle.Add(0x197, "fillShadeColors");
            fillStyle.Add(0x198, "fillOriginX");
            fillStyle.Add(0x199, "fillOriginY");
            fillStyle.Add(0x19a, "fillShapeOriginX");
            fillStyle.Add(0x19b, "fillShapeOriginY");
            fillStyle.Add(0x19c, "fillShadeType");

            fillStyle.Add(0x19e, "fillColorExt");
            fillStyle.Add(0x19f, "reserved415");
            fillStyle.Add(0x1a0, "fillColorExtMod");
            fillStyle.Add(0x1a1, "reserved417");
            fillStyle.Add(0x1a2, "fillBackColorExt");
            fillStyle.Add(0x1a3, "reserved419");
            fillStyle.Add(0x1a4, "fillBackColorExtMod");
            fillStyle.Add(0x1a5, "reserved421");
            fillStyle.Add(0x1a6, "reserved422");
            fillStyle.Add(0x1a7, "reserved423");
            fillStyle.Add(0x1bf, "fillStyleBoolean");
        }
        public static string GetFillStyleName(int optionId)
        {
            if (fillStyle.TryGetValue(optionId, out string name))
                return name;
            return "Unknown";
        }
        #endregion
    }
}
