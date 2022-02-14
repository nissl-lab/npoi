using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.OOXML.XSSF.UserModel
{
    public class XSSFBuiltinTableStyle
    {
        private static Dictionary<XSSFBuiltinTableStyle, ITableStyle> styleMap = new Dictionary<XSSFBuiltinTableStyle, ITableStyle>();
        public ITableStyle GetStyle()
        {
            Init();
            return styleMap[this];
        }
        public static bool IsBuiltinStyle(ITableStyle style)
        {
            if (style == null) return false;
            try
            {
                XSSFBuiltinTableStyle.valueOf(style.Name);
                return true;
            }
            catch (ArgumentException)
            {
                return false;
            }
        }
        protected  class XSSFBuiltinTypeStyleStyle : ITableStyle
        {

        private XSSFBuiltinTableStyle builtIn;
        private ITableStyle style;

        /**
         * @param builtIn
         * @param style
         */
        protected XSSFBuiltinTypeStyleStyle(XSSFBuiltinTableStyle builtIn, ITableStyle style)
        {
            this.builtIn = builtIn;
            this.style = style;
        }

        public String Name
        {
            get
            {
                return style.Name;
            }
        }

        public int Index
        {
            get
            {
                return builtIn.ordinal();
            }
        }

        public bool IsBuiltin
        {
            get
            {
                return true;
            }
        }

        public DifferentialStyleProvider GetStyle(TableStyleType type)
        {
            return style.GetStyle(type);
        }

    }
}
}
