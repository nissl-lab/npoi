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
namespace NPOI.SS.UserModel
{
    using NPOI.SS.Util;
    using System;
    /// <summary>
    /// An object that handles instantiating concrete
    ///  classes of the various instances one needs for
    ///  HSSF and XSSF.
    /// Works around a limitation in Java where we
    ///  cannot have static methods on interfaces or abstract
    ///  classes.
    /// This allows you to Get the appropriate class for
    ///  a given interface, without you having to worry
    ///  about if you're dealing with HSSF or XSSF.
    /// </summary>
    public interface ICreationHelper
    {
        /// <summary>
        /// Creates a new RichTextString instance
        /// </summary>
        /// <param name="text">The text to initialise the RichTextString with</param>
        IRichTextString CreateRichTextString(String text);

        /// <summary>
        /// Creates a new DataFormat instance
        /// </summary>
        IDataFormat CreateDataFormat();

        /// <summary>
        /// Creates a new Hyperlink, of the given type
        /// </summary>
        IHyperlink CreateHyperlink(HyperlinkType type);

        /// <summary>
        /// Creates FormulaEvaluator - an object that evaluates formula cells.
        /// </summary>
        /// <return>a FormulaEvaluator instance</return>
        IFormulaEvaluator CreateFormulaEvaluator();

        /// <summary>
        /// Creates a XSSF-style Color object, used for extended sheet
        ///  formattings and conditional formattings
        /// </summary>
        ExtendedColor CreateExtendedColor();
        /// <summary>
        /// Creates a ClientAnchor. Use this object to position Drawing object in a sheet
        /// </summary>
        /// <return>a ClientAnchor instance</return>
        /// <see cref="IDrawing" />
        IClientAnchor CreateClientAnchor();

        /// <summary>
        /// Creates an AreaReference.
        /// </summary>
        /// <param name="reference">cell reference</param>
        /// <return>an AreaReference instance</return>
        AreaReference CreateAreaReference(String reference);

        /// <summary>
        /// Creates an area ref from a pair of Cell References..
        /// </summary>
        /// <param name="topLeft">cell reference</param>
        /// <param name="bottomRight">cell reference</param>
        /// <return>an AreaReference instance</return>
        AreaReference CreateAreaReference(CellReference topLeft, CellReference bottomRight);
    }
}
