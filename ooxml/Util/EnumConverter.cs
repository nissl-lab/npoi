using System;
using System.Collections.Generic;
using System.Text;
using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;
using System.Reflection;

namespace NPOI.Util
{
    public static class EnumConverter
    {       
        public static ST_Jc ValueOf(ParagraphAlignment val)
        {
            switch (val)
            {
                case ParagraphAlignment.BOTH: return ST_Jc.both;
                case ParagraphAlignment.CENTER: return ST_Jc.center;
                case ParagraphAlignment.DISTRIBUTE: return ST_Jc.distribute;
                case ParagraphAlignment.HIGH_KASHIDA: return ST_Jc.highKashida;
                case ParagraphAlignment.LOW_KASHIDA: return ST_Jc.lowKashida;
                case ParagraphAlignment.MEDIUM_KASHIDA: return ST_Jc.mediumKashida;
                case ParagraphAlignment.NUM_TAB: return ST_Jc.numTab;
                case ParagraphAlignment.RIGHT: return ST_Jc.right;
                case ParagraphAlignment.THAI_DISTRIBUTE: return ST_Jc.thaiDistribute;
                default: return ST_Jc.left;
            }
        }
        public static ParagraphAlignment ValueOf(ST_Jc val)
        {
            switch (val)
            {
                case ST_Jc.both: return ParagraphAlignment.BOTH;
                case ST_Jc.center: return ParagraphAlignment.CENTER;
                case ST_Jc.distribute: return ParagraphAlignment.DISTRIBUTE;

                default: return ParagraphAlignment.LEFT;
            }
        }
        public static T ValueOf<T, F>(F val)
        {
            string value = Enum.GetName(val.GetType(), val);
            return (T)Enum.Parse(typeof(T), value, true);
        }
    }
}
