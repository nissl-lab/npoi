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

namespace NPOI.Common.UserModel.Fonts
{
    /// <summary>
    /// <para>
    /// A FontInfo object holds information about a font configuration.
    /// It is roughly an equivalent to the LOGFONT structure in Windows GDI.
    /// </para>
    /// <para>
    /// If an implementation doesn't provide a property, the Getter will return <c>null</c> -
    /// if the value is unset, a default value will be returned.
    /// </para>
    /// <para>
    /// Setting a unsupported property results in an <see cref="UnsupportedOperationException"/>.
    /// </para>
    /// </summary>
    /// 
    /// @see <a href="https://msdn.microsoft.com/en-us/library/dd145037.aspx">LOGFONT structure</a>
    /// <remarks>
    /// @since POI 3.17-beta2
    /// </remarks>
    public interface IFontInfo
    {

        /// <summary>
        /// Get the index within the collection of Font objects
        /// </summary>
        /// <returns>unique index number of the underlying record this Font represents
        /// (probably you don't care unless you're comparing which one is which)
        /// </returns>
        int GetIndex();

        /// <summary>
        /// Sets the index within the collection of Font objects
        /// </summary>
        /// <param name="index">the index within the collection of Font objects</param>
        /// 
        /// <exception cref="UnsupportedOperationException">if unsupported</exception>
        void SetIndex(int index);


        /// <summary>
        /// </summary>
        /// <returns>the full name of the font, i.e. font family + type face</returns>\1string GetTypeface();

        /// <summary>
        /// Sets the font name
        /// </summary>
        /// <param name="typeface">the full name of the font, when <c>null</c> removes the font definition -
        /// removal is implementation specific
        /// </param>
        void SetTypeface(string typeface);

        /// <summary>
        /// </summary>
        /// <returns>the font charset</returns>
        FontCharset GetCharset();

        /// <summary>
        /// Sets the charset
        /// </summary>
        /// <param name="charset">the charset</param>
        void SetCharset(FontCharset charset);

        /// <summary>
        /// </summary>
        /// <returns>the family class</returns>
        FontFamily GetFamily();

        /// <summary>
        /// Sets the font family class
        /// </summary>
        /// <param name="family">the font family class</param>
        void SetFamily(FontFamily family);

        /// <summary>
        /// </summary>
        /// <returns>the font pitch or <c>null</c> if unsupported</returns>
        FontPitch GetPitch();

        /// <summary>
        /// Set the font pitch
        /// </summary>
        /// <param name="pitch">the font pitch</param>
        /// 
        /// <exception cref="UnsupportedOperationException">if unsupported</exception>
        void SetPitch(FontPitch pitch);
    }
}


