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
    using System;

    using NPOI.SS.Util;
    /// <summary>
    /// Error style constants for error box
    /// </summary>
    public static class ERRORSTYLE
    {
        /// <summary>
        /// STOP style */
        /// </summary>
        public const int STOP = 0x00;
        /// <summary>
        /// WARNING style */
        /// </summary>
        public const int WARNING = 0x01;
        /// <summary>
        /// INFO style */
        /// </summary>
        public const int INFO = 0x02;
    }

    public interface IDataValidation
    {
        IDataValidationConstraint ValidationConstraint { get; }

        /// <summary>
        /// get or set the error style for error box
        /// </summary>
        int ErrorStyle { get; set; }
        /// <summary>
        /// Setting this allows an empty object as a valid value. Retrieve the settings for empty cells allowed.
        /// @return True if this object should treats empty as valid value , false otherwise
        /// </summary>
        /// <value><c>true</c> if this object should treats empty as valid value, <c>false</c>  otherwise</value>
        bool EmptyCellAllowed { get; set; }
        /// <summary>
        /// Useful for list validation objects .
        /// Useful only list validation objects . This method always returns false if the object isn't a list validation object
        /// </summary>
        bool SuppressDropDownArrow { get; set; }
        /*
         * Useful for list validation objects .
         * 
         * @param suppress
         *            True if a list should display the values into a drop down list ,
         *            false otherwise . In other words , if a list should display
         *            the arrow sign on its right side
         */
        //void SetSuppressDropDownArrow(bool suppress);

        /*
         * Useful only list validation objects . This method always returns false if
         * the object isn't a list validation object
         * 
         * @return <c>true</c> if a list should display the values into a drop down list ,
         *         <c>false</c> otherwise .
         */
        //bool GetSuppressDropDownArrow();

        /// <summary>
        /// <para>
        /// Sets the behaviour when a cell which belongs to this object is selected
        /// </para>
        /// <para>
        /// <value><c>true</c> if an prompt box should be displayed , <c>false</c> otherwise</value>
        /// </para>
        /// </summary>
        bool ShowPromptBox { get; set; }
        //void SetShowPromptBox(bool Show);

        /*
         * @return <c>true</c> if an prompt box should be displayed , <c>false</c> otherwise
         */
        //bool GetShowPromptBox();

        /// <summary>
        /// <para>
        /// Sets the behaviour when an invalid value is entered
        /// </para>
        /// <para>
        /// <value><c>true</c> if an error box should be displayed , <c>false</c> otherwise</value>
        /// </para>
        /// </summary>
        bool ShowErrorBox { get; set; }
        //void SetShowErrorBox(bool Show);

        /*
         * @return <c>true</c> if an error box should be displayed , <c>false</c> otherwise
         */
        //bool GetShowErrorBox();

        /// <summary>
        /// Sets the title and text for the prompt box . Prompt box is displayed when
        /// the user selects a cell which belongs to this validation object . In
        /// order for a prompt box to be displayed you should also use method
        /// SetShowPromptBox( bool show )
        /// </summary>
        /// <param name="title">The prompt box's title</param>
        /// <param name="text">The prompt box's text</param>
        void CreatePromptBox(String title, String text);
        /// <summary>
        /// </summary>
        /// <returns>Prompt box's title or <c>null</c></returns>
        String PromptBoxTitle { get; }

        /// <summary>
        /// </summary>
        /// <returns>Prompt box's text or <c>null</c></returns>
        String PromptBoxText { get; }

        /// <summary>
        /// Sets the title and text for the error box . Error box is displayed when
        /// the user enters an invalid value int o a cell which belongs to this
        /// validation object . In order for an error box to be displayed you should
        /// also use method SetShowErrorBox( bool show )
        /// </summary>
        /// <param name="title">The error box's title</param>
        /// <param name="text">The error box's text</param>
        void CreateErrorBox(String title, String text);

        /// <summary>
        /// </summary>
        /// <returns>Error box's title or <c>null</c></returns>

        String ErrorBoxTitle { get; }

        /// <summary>
        /// </summary>
        /// <returns>Error box's text or <c>null</c></returns>
        String ErrorBoxText { get; }

        CellRangeAddressList Regions { get; }

    }

}