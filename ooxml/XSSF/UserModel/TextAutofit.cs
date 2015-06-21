/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for Additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */
namespace NPOI.XSSF.UserModel
{
    /**
     * Specifies a list of auto-fit types.
     * <p>
     * Autofit specifies that a shape should be auto-fit to fully contain the text described within it.
     * Auto-fitting is when text within a shape is scaled in order to contain all the text inside
     * </p>
     */
    public enum TextAutofit
    {
        /**
         * Specifies that text within the text body should not be auto-fit to the bounding box.
         * Auto-fitting is when text within a text box is scaled in order to remain inside
         * the text box.
         */
        NONE,
        /**
         * Specifies that text within the text body should be normally auto-fit to the bounding box.
         * Autofitting is when text within a text box is scaled in order to remain inside the text box.
         *
         * <p>
         * <em>Example:</em> Consider the situation where a user is building a diagram and needs
         * to have the text for each shape that they are using stay within the bounds of the shape.
         * An easy way this might be done is by using NORMAL autofit
         * </p>
         */
        NORMAL,
        /**
         * Specifies that a shape should be auto-fit to fully contain the text described within it.
         * Auto-fitting is when text within a shape is scaled in order to contain all the text inside.
         *
         * <p>
         * <em>Example:</em> Consider the situation where a user is building a diagram and needs to have
         * the text for each shape that they are using stay within the bounds of the shape.
         * An easy way this might be done is by using SHAPE autofit
         * </p>
         */
        SHAPE
    }

}