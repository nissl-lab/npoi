/*
* Licensed to the Apache Software Foundation (ASF) under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed under the License is distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations under the License.
*/
/*
 * Created on May 29, 2005
 *
 */

namespace TestCases.SS.Formula.Functions
{
    using System;
    using NUnit.Framework;
    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *
     */
    public abstract class AbstractNumericTestCase
    {

        public static double POS_ZERO = 1E-4;
        public static double DIFF_TOLERANCE_FACTOR = 1E-8;

        public void SetUp()
        {
        }

        public void tearDown()
        {
        }

        /**
         * Why doesnt JUnit have a method like this for doubles? 
         * The current impl (3.8.1) of Junit has a retar*** method
         * for comparing doubles. DO NOT use that.
         * TODO: This class should really be in an abstract super class
         * to avoid code duplication across this project.
         * @param message
         * @param baseval
         * @param Checkval
         */
        public static void AssertEquals(String message, double baseval, double checkval, double almostZero, double diffToleranceFactor)
        {
            double posZero = Math.Abs(almostZero);
            double negZero = -1 * posZero;
            if (Double.IsNaN(baseval))
            {
                Assert.IsTrue(Double.IsNaN(baseval), message + ": Expected " + baseval + " but was " + checkval);
            }
            else if (Double.IsInfinity(baseval))
            {
                Assert.IsTrue(Double.IsInfinity(baseval) && ((baseval < 0) == (checkval < 0)), message + ": Expected " + baseval + " but was " + checkval);
            }
            else
            {
                Assert.IsTrue(baseval != 0 ? Math.Abs(baseval - checkval) <= Math.Abs(diffToleranceFactor * baseval) : checkval < posZero && checkval > negZero, message + ": Expected " + baseval + " but was " + checkval);
            }
        }

        public static void AssertEquals(String msg, double baseval, double checkval)
        {
            AssertEquals(msg, baseval, checkval, POS_ZERO, DIFF_TOLERANCE_FACTOR);
        }

    }

}