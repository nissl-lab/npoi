/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

namespace NPOI.HSLF.Model;

using NPOI.HSLF.record.FontEntityAtom;

/**
 * Represents a Font used in a presenation.
 * <p>
 * In PowerPoint Font is a shared resource and can be shared among text object in the presentation.
 * </p>
 * Some commonly used fonts are predefined in static constants.
 *
 * @author Yegor Kozlov
 */
public class PPFont {
    /**
     * ANSI character Set
     */
    public static byte ANSI_CHARSET = 0;

    /**
     * Default character Set.
     */
    public static byte DEFAULT_CHARSET = 1;

    /**
     * Symbol character Set
     */
    public static byte SYMBOL_CHARSET = 2;


    /**
     * Constants for the pitch and family of the font.
     * The two low-order bits specify the pitch of the font and can be one of the following values
     */
    public static byte DEFAULT_PITCH  = 0;
    public static byte FIXED_PITCH    = 1;
    public static byte VARIABLE_PITCH = 2;

    /**
     * Don't care or don't know.
     */
    public static byte FF_DONTCARE    = 0;
    /**
     * Fonts with variable stroke width (proportional) and with serifs. Times New Roman is an example.
     */
    public static byte FF_ROMAN       = 16;
    /**
     * Fonts with variable stroke width (proportional) and without serifs. Arial is an example.
     */
    public static byte FF_SWISS       = 32;
    /**
     * Fonts designed to look like handwriting. Script and Cursive are examples.
     */
    public static byte FF_SCRIPT      = 64;
    /**
     * Fonts with constant stroke width (monospace), with or without serifs.
     * Monospace fonts are usually modern. CourierNew is an example
     */
    public static byte FF_MODERN      = 48;
    /**
     * Novelty fonts. Old English is an example
     */
    public static byte FF_DECORATIVE  = 80;


    protected int charset;
    protected int type;
    protected int flags;
    protected int pitch;
    protected String name;

    /**
     * Creates a new instance of PPFont
     */
    public PPFont(){

    }

    /**
     * Creates a new instance of PPFont and Initialize it from the supplied font atom
     */
    public PPFont(FontEntityAtom fontAtom){
        name = fontAtom.GetFontName();
        charset = fontAtom.GetCharSet();
        type = fontAtom.GetFontType();
        flags = fontAtom.GetFontFlags();
        pitch = fontAtom.GetPitchAndFamily();
    }

    /**
     * Set the name for the font (i.e. Arial)
     *
     * @param val  String representing the name of the font to use
     */
     public void SetFontName(String val){
        name = val;
    }

    /**
     * Get the name for the font (i.e. Arial)
     *
     * @return String representing the name of the font to use
     */
    public String GetFontName(){
        return name;
    }

    /**
     * Set the character Set
     *
     * @param val - characterset
     */
    public void SetCharSet(int val){
        charset = val;
    }

    /**
     * Get the character Set
     *
     * @return charset - characterset
     */
    public int GetCharSet(){
        return charset;
    }

    /**
     * Set the font flags
     * Bit 1: If Set, font is subsetted
     *
     * @param val - the font flags
     */
    public void SetFontFlags(int val){
        flags = val;
    }

    /**
     * Get the character Set
     * Bit 1: If Set, font is subsetted
     *
     * @return the font flags
     */
    public int GetFontFlags(){
        return flags;
    }

    /**
     * Set the font type
     * <p>
     * Bit 1: Raster Font
     * Bit 2: Device Font
     * Bit 3: TrueType Font
     * </p>
     *
     * @param val - the font type
     */
    public void SetFontType(int val){
        type = val;
    }

    /**
     * Get the font type
     * <p>
     * Bit 1: Raster Font
     * Bit 2: Device Font
     * Bit 3: TrueType Font
     * </p>
     *
     * @return the font type
     */
    public int GetFontType(){
        return type;
    }

    /**
     * Set lfPitchAndFamily
     *
     *
     * @param val - Corresponds to the lfPitchAndFamily field of the Win32 API LOGFONT structure
     */
    public void SetPitchAndFamily(int val){
        pitch = val;
    }

    /**
     * Get lfPitchAndFamily
     *
     * @return corresponds to the lfPitchAndFamily field of the Win32 API LOGFONT structure
     */
    public int GetPitchAndFamily(){
        return pitch;
    }

    public static PPFont ARIAL;
    public static PPFont TIMES_NEW_ROMAN ;
    public static PPFont COURIER_NEW;
    public static PPFont WINGDINGS;
    static {
        ARIAL = new PPFont();
        ARIAL.SetFontName("Arial");
        ARIAL.SetCharSet(ANSI_CHARSET);
        ARIAL.SetFontType(4);
        ARIAL.SetFontFlags(0);
        ARIAL.SetPitchAndFamily(VARIABLE_PITCH | FF_SWISS);

        TIMES_NEW_ROMAN = new PPFont();
        TIMES_NEW_ROMAN.SetFontName("Times New Roman");
        TIMES_NEW_ROMAN.SetCharSet(ANSI_CHARSET);
        TIMES_NEW_ROMAN.SetFontType(4);
        TIMES_NEW_ROMAN.SetFontFlags(0);
        TIMES_NEW_ROMAN.SetPitchAndFamily(VARIABLE_PITCH | FF_ROMAN);

        COURIER_NEW = new PPFont();
        COURIER_NEW.SetFontName("Courier New");
        COURIER_NEW.SetCharSet(ANSI_CHARSET);
        COURIER_NEW.SetFontType(4);
        COURIER_NEW.SetFontFlags(0);
        COURIER_NEW.SetPitchAndFamily(FIXED_PITCH | FF_MODERN);

        WINGDINGS = new PPFont();
        WINGDINGS.SetFontName("Wingdings");
        WINGDINGS.SetCharSet(SYMBOL_CHARSET);
        WINGDINGS.SetFontType(4);
        WINGDINGS.SetFontFlags(0);
        WINGDINGS.SetPitchAndFamily(VARIABLE_PITCH | FF_DONTCARE);
    }
}





