using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.OpenXmlFormats.Shared;
using NPOI.WP.UserModel;
using NPOI.XWPF.UserModel;
using W = NPOI.OpenXmlFormats.Wordprocessing;

namespace NPOI.XWPF.Usermodel
{
    public class XWPFSharedRun : ICharacterRun
    {
        private CT_R run;
        private IRunBody parent;

        public XWPFSharedRun(CT_R ctR, IRunBody p)
        {
            this.run = ctR;
            this.parent = p;
            SetFontFamily("Cambria Math", FontCharRange.None);
        }

        public bool IsBold { get; set; }
        public bool IsItalic {
            get {
                W.CT_RPr pr = run.rPr1;
                if (pr == null || !pr.IsSetI())
                    return false;
                return IsCTOnOff(pr.i);
            }
            set
            {
                W.CT_RPr pr = run.IsSetRPr1() ? run.rPr1 : run.AddNewRPr1();
                W.CT_OnOff italic = pr.IsSetI() ? pr.i : pr.AddNewI();
                italic.val = value;
            }
        }

        /**
         * For isBold, isItalic etc
         */
        private bool IsCTOnOff(W.CT_OnOff onoff)
        {
            if (!onoff.IsSetVal())
                return true;
            return onoff.val;
        }

        public bool IsSmallCaps { get; set; }
        public bool IsCapitalized { get; set; }
        public bool IsStrikeThrough { get; set; }
        public bool IsDoubleStrikeThrough { get; set; }
        public bool IsShadowed { get; set; }
        public bool IsEmbossed { get; set; }
        public bool IsImprinted { get; set; }
        public int CharacterSpacing { get; set; }
        public int Kerning { get; set; }
        public bool IsHighlighted { get; set; }

        public string FontName
        {
            get { return FontFamily; }
        }

        public string FontFamily
        {
            get
            {
                return GetFontFamily(FontCharRange.None);
            }
            set
            {
                SetFontFamily(value, FontCharRange.None);
            }
        }
        /// <summary>
        /// Specifies the fonts which shall be used to display the text contents of
        /// this run.The default handling for fcr == null is to overwrite the
        /// ascii font char range with the given font family and also set all not
        /// specified font ranges
        /// </summary>
        /// <param name="fontFamily">fontFamily</param>
        /// <param name="fcr">FontCharRange or null for default handling</param>
        public void SetFontFamily(String fontFamily, FontCharRange fcr)
        {
            W.CT_RPr pr = run.IsSetRPr1() ? run.rPr1 : run.AddNewRPr1();
            W.CT_Fonts fonts = pr.IsSetRFonts() ? pr.rFonts : pr.AddNewRFonts();

            if (fcr == FontCharRange.None)
            {
                fonts.ascii = (fontFamily);
                if (!fonts.IsSetHAnsi())
                {
                    fonts.hAnsi = (fontFamily);
                }
                if (!fonts.IsSetCs())
                {
                    fonts.cs = (fontFamily);
                }
                if (!fonts.IsSetEastAsia())
                {
                    fonts.eastAsia = (fontFamily);
                }
            }
            else
            {
                switch (fcr)
                {
                    case FontCharRange.Ascii:
                        fonts.ascii = (fontFamily);
                        break;
                    case FontCharRange.CS:
                        fonts.cs = (fontFamily);
                        break;
                    case FontCharRange.EastAsia:
                        fonts.eastAsia = (fontFamily);
                        break;
                    case FontCharRange.HAnsi:
                        fonts.hAnsi = (fontFamily);
                        break;
                }
            }
        }

        /// <summary>
        /// Gets the font family for the specified font char range.
        /// If fcr is null, the font char range "ascii" is used
        /// Please use "Cambria Math"(set as default) font otherwise MS Word 
        /// don't open file, LibreOffice Writer open it normaly.
        /// I think this is MS Word bug, because this is not standart.
        /// </summary>
        /// <param name="fcr">the font char range, defaults to "ansi"</param>
        /// <returns>a string representing the font famil</returns>
        public String GetFontFamily(FontCharRange fcr)
        {
            OpenXmlFormats.Wordprocessing.CT_RPr pr = run.rPr1;
            if (pr == null || !pr.IsSetRFonts()) return null;

            OpenXmlFormats.Wordprocessing.CT_Fonts fonts = pr.rFonts;
            switch (fcr == FontCharRange.None ? FontCharRange.Ascii : fcr)
            {
                default:
                case FontCharRange.Ascii:
                    return fonts.ascii;
                case FontCharRange.CS:
                    return fonts.cs;
                case FontCharRange.EastAsia:
                    return fonts.eastAsia;
                case FontCharRange.HAnsi:
                    return fonts.hAnsi;
            }
        }

        public int FontSize { get; set; }

        public string Text
        {
            get
            {
                StringBuilder text = new StringBuilder();
                for (int i = 0; i < run.Items.Count; i++)
                {
                    object o = run.Items[i];
                    if (o is CT_Text1)
                    {
                        text.Append(((CT_Text1)o).Value);
                    }
                }

                return text.ToString();
            }
        }

        /// <summary>
        ///Sets the text of this text run
        /// </summary>
        /// <param name="value">the literal text which shall be displayed in the document</param>
        public XWPFSharedRun SetText(string value)
        {
            SetText(value, 0);
            return this;
        }

        /// <summary>
        /// Sets the text of this text run.in the 
        /// </summary>
        /// <param name="value">the literal text which shall be displayed in the document</param>
        /// <param name="pos">position in the text array (NB: 0 based)</param>
        private XWPFSharedRun SetText(String value, int pos)
        {
            int length = run.SizeOfTArray();
            if (pos > length) throw new IndexOutOfRangeException("Value too large for the parameter position");
            CT_Text1 t = (pos < length && pos >= 0) ? run.GetTArray(pos) : run.AddNewT();
            t.Value = (value);
            preserveSpaces(t);
            return this;
        }

        /// <summary>
        /// Add the xml:spaces="preserve" attribute if the string has leading or trailing white spaces
        /// </summary>
        /// <param name="xs">the string to check</param>
        static void preserveSpaces(CT_Text1 xs)
        {
            String text = xs.Value;
            if (text != null && (text.StartsWith(" ") || text.EndsWith(" ")))
            {
                //    XmlCursor c = xs.NewCursor();
                //    c.ToNextToken();
                //    c.InsertAttributeWithValue(new QName("http://www.w3.org/XML/1998/namespace", "space"), "preserve");
                //    c.Dispose();
                xs.space = "preserve";
            }
        }
    }
}
