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
     * Vertical Text Types
     */
    public enum TextDirection
    {
        None,
        /**
         * Horizontal text. This should be default.
         */
        HORIZONTAL,
        /**
         * Vertical orientation.
         * (each line is 90 degrees rotated clockwise, so it goes
         * from top to bottom; each next line is to the left from
         * the previous one).
         */
        VERTICAL,
        /**
         * Vertical orientation.
         * (each line is 270 degrees rotated clockwise, so it goes
         * from bottom to top; each next line is to the right from
         * the previous one).
         */
        VERTICAL_270,
        /**
         * Determines if all of the text is vertical
         * ("one letter on top of another").
         */
        STACKED
    }

}