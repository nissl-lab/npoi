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
namespace NPOI.SS.Formula.functions;

using junit.framework.TestCase;

/**
 * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
 *
 */
public abstract class AbstractNumericTestCase  {

    public static double POS_ZERO = 1E-4;
    public static double DIFF_TOLERANCE_FACTOR = 1E-8;

    public void SetUp() {
    }

    public void tearDown() {
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
    public static void Assert.AreEqual(String message, double basEval, double Checkval, double almostZero, double diffToleranceFactor) {
        double posZero = Math.abs(almostZero);
        double negZero = -1 * posZero;
        if (Double.IsNaN(basEval)) {
            Assert.IsTrue(message+": Expected " + baseval + " but was " + Checkval
                    , Double.IsNaN(basEval));
        }
        else if (Double.IsInfInity(baseval)) {
            Assert.IsTrue(message+": Expected " + baseval + " but was " + Checkval
                    , Double.IsInfInity(baseval) && ((basEval<0) == (Checkval<0)));
        }
        else {
            Assert.IsTrue(message+": Expected " + baseval + " but was " + Checkval
                ,baseval != 0
                    ? Math.abs(baseval - Checkval) <= Math.abs(diffToleranceFactor * basEval)
                    : Checkval < posZero && Checkval > negZero);
        }
    }

    public static void Assert.AreEqual(String msg, double basEval, double Checkval) {
        Assert.AreEqual(msg, basEval, Checkval, POS_ZERO, DIFF_TOLERANCE_FACTOR);
    }

}

