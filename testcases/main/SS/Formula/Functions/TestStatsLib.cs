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
 * Created on May 30, 2005
 *
 */
namespace NPOI.SS.Formula.functions;

using junit.framework.AssertionFailedError;

using NPOI.SS.Formula.Eval.ErrorEval;
using NPOI.SS.Formula.Eval.EvaluationException;


/**
 * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
 *
 */
public class TestStatsLib : AbstractNumericTestCase {

    public void TestDevsq() {
        double[] v = null;
        double d, x = 0;
        
        v = new double[] {1,2,3,4,5,6,7,8,9,10};
        d = StatsLib.devsq(v);
        x = 82.5;
        Assert.AreEqual("devsq ", x, d);
        
        v = new double[] {1,1,1,1,1,1,1,1,1,1};
        d = StatsLib.devsq(v);
        x = 0;
        Assert.AreEqual("devsq ", x, d);
        
        v = new double[] {0,0,0,0,0,0,0,0,0,0};
        d = StatsLib.devsq(v);
        x = 0;
        Assert.AreEqual("devsq ", x, d);
        
        v = new double[] {1,2,1,2,1,2,1,2,1,2};
        d = StatsLib.devsq(v);
        x = 2.5;
        Assert.AreEqual("devsq ", x, d);
        
        v = new double[] {123.12,33.3333,2d/3d,5.37828,0.999};
        d = StatsLib.devsq(v);
        x = 10953.7416965767;
        Assert.AreEqual("devsq ", x, d);
        
        v = new double[] {-1,-2,-3,-4,-5,-6,-7,-8,-9,-10};
        d = StatsLib.devsq(v);
        x = 82.5;
        Assert.AreEqual("devsq ", x, d);
    }

    public void TestKthLargest() {
        double[] v = null;
        double d, x = 0;
        
        v = new double[] {1,2,3,4,5,6,7,8,9,10};
        d = StatsLib.kthLargest(v, 3);
        x = 8;
        Assert.AreEqual("kthLargest ", x, d);
        
        v = new double[] {1,1,1,1,1,1,1,1,1,1};
        d = StatsLib.kthLargest(v, 3);
        x = 1;
        Assert.AreEqual("kthLargest ", x, d);
        
        v = new double[] {0,0,0,0,0,0,0,0,0,0};
        d = StatsLib.kthLargest(v, 3);
        x = 0;
        Assert.AreEqual("kthLargest ", x, d);
        
        v = new double[] {1,2,1,2,1,2,1,2,1,2};
        d = StatsLib.kthLargest(v, 3);
        x = 2;
        Assert.AreEqual("kthLargest ", x, d);
        
        v = new double[] {123.12,33.3333,2d/3d,5.37828,0.999};
        d = StatsLib.kthLargest(v, 3);
        x = 5.37828;
        Assert.AreEqual("kthLargest ", x, d);
        
        v = new double[] {-1,-2,-3,-4,-5,-6,-7,-8,-9,-10};
        d = StatsLib.kthLargest(v, 3);
        x = -3;
        Assert.AreEqual("kthLargest ", x, d);
    }

    public void TestKthSmallest() {
    }

    public void TestAvedev() {
        double[] v = null;
        double d, x = 0;
        
        v = new double[] {1,2,3,4,5,6,7,8,9,10};
        d = StatsLib.avedev(v);
        x = 2.5;
        Assert.AreEqual("avedev ", x, d);
        
        v = new double[] {1,1,1,1,1,1,1,1,1,1};
        d = StatsLib.avedev(v);
        x = 0;
        Assert.AreEqual("avedev ", x, d);
        
        v = new double[] {0,0,0,0,0,0,0,0,0,0};
        d = StatsLib.avedev(v);
        x = 0;
        Assert.AreEqual("avedev ", x, d);
        
        v = new double[] {1,2,1,2,1,2,1,2,1,2};
        d = StatsLib.avedev(v);
        x = 0.5;
        Assert.AreEqual("avedev ", x, d);
        
        v = new double[] {123.12,33.3333,2d/3d,5.37828,0.999};
        d = StatsLib.avedev(v);
        x = 36.42176053333;
        Assert.AreEqual("avedev ", x, d);
        
        v = new double[] {-1,-2,-3,-4,-5,-6,-7,-8,-9,-10};
        d = StatsLib.avedev(v);
        x = 2.5;
        Assert.AreEqual("avedev ", x, d);
    }

    public void TestMedian() {
        double[] v = null;
        double d, x = 0;
        
        v = new double[] {1,2,3,4,5,6,7,8,9,10};
        d = StatsLib.median(v);
        x = 5.5;
        Assert.AreEqual("median ", x, d);
        
        v = new double[] {1,1,1,1,1,1,1,1,1,1};
        d = StatsLib.median(v);
        x = 1;
        Assert.AreEqual("median ", x, d);
        
        v = new double[] {0,0,0,0,0,0,0,0,0,0};
        d = StatsLib.median(v);
        x = 0;
        Assert.AreEqual("median ", x, d);
        
        v = new double[] {1,2,1,2,1,2,1,2,1,2};
        d = StatsLib.median(v);
        x = 1.5;
        Assert.AreEqual("median ", x, d);
        
        v = new double[] {123.12,33.3333,2d/3d,5.37828,0.999};
        d = StatsLib.median(v);
        x = 5.37828;
        Assert.AreEqual("median ", x, d);
        
        v = new double[] {-1,-2,-3,-4,-5,-6,-7,-8,-9,-10};
        d = StatsLib.median(v);
        x = -5.5;
        Assert.AreEqual("median ", x, d);
        
        v = new double[] {-2,-3,-4,-5,-6,-7,-8,-9,-10};
        d = StatsLib.median(v);
        x = -6;
        Assert.AreEqual("median ", x, d);
        
        v = new double[] {1,2,3,4,5,6,7,8,9};
        d = StatsLib.median(v);
        x = 5;
        Assert.AreEqual("median ", x, d);
    }

    public void TestMode() {
        double[] v;
        double d, x = 0;
        
        v = new double[] {1,2,3,4,5,6,7,8,9,10};
        ConfirmMode(v, null);
        
        v = new double[] {1,1,1,1,1,1,1,1,1,1};
        ConfirmMode(v, 1.0);
        
        v = new double[] {0,0,0,0,0,0,0,0,0,0};
        ConfirmMode(v, 0.0);
        
        v = new double[] {1,2,1,2,1,2,1,2,1,2};
        ConfirmMode(v, 1.0);
        
        v = new double[] {123.12,33.3333,2d/3d,5.37828,0.999};
        ConfirmMode(v, null);
        
        v = new double[] {-1,-2,-3,-4,-5,-6,-7,-8,-9,-10};
        ConfirmMode(v, null);
        
        v = new double[] {1,2,3,4,1,1,1,1,0,0,0,0,0};
        ConfirmMode(v, 1.0);
        
        v = new double[] {0,1,2,3,4,1,1,1,0,0,0,0,1};
        ConfirmMode(v, 0.0);
    }
    private static void ConfirmMode(double[] v, double expectedResult) {
    	ConfirmMode(v, expectedResult);
    }
    private static void ConfirmMode(double[] v, Double expectedResult) {
    	double actual;
		try {
			actual = Mode.Evaluate(v);
			if (expectedResult == null) {
				throw new AssertionFailedError("Expected N/A exception was not thrown");
			}
		} catch (EvaluationException e) {
			if (expectedResult == null) {
				Assert.AreEqual(ErrorEval.NA, e.GetErrorEval());
				return;
			}
			throw new RuntimeException(e);
		}
    	Assert.AreEqual("mode", expectedResult , actual);
    }
    

    public void TestStddev() {
        double[] v = null;
        double d, x = 0;
        
        v = new double[] {1,2,3,4,5,6,7,8,9,10};
        d = StatsLib.stdev(v);
        x = 3.02765035410;
        Assert.AreEqual("stdev ", x, d);
        
        v = new double[] {1,1,1,1,1,1,1,1,1,1};
        d = StatsLib.stdev(v);
        x = 0;
        Assert.AreEqual("stdev ", x, d);
        
        v = new double[] {0,0,0,0,0,0,0,0,0,0};
        d = StatsLib.stdev(v);
        x = 0;
        Assert.AreEqual("stdev ", x, d);
        
        v = new double[] {1,2,1,2,1,2,1,2,1,2};
        d = StatsLib.stdev(v);
        x = 0.52704627669;
        Assert.AreEqual("stdev ", x, d);
        
        v = new double[] {123.12,33.3333,2d/3d,5.37828,0.999};
        d = StatsLib.stdev(v);
        x = 52.33006233652;
        Assert.AreEqual("stdev ", x, d);
        
        v = new double[] {-1,-2,-3,-4,-5,-6,-7,-8,-9,-10};
        d = StatsLib.stdev(v);
        x = 3.02765035410;
        Assert.AreEqual("stdev ", x, d);
    }

    public void TestVar() {
        double[] v = null;
        double d, x = 0;

        v = new double[] {3.50, 5.00, 7.23, 2.99};
        d = StatsLib.var(v);
        x = 3.6178;
        Assert.AreEqual("var ", x, d);

        v = new double[] {34.5, 2.0, 8.9, -4.0};
        d = StatsLib.var(v);
        x = 286.99;
        Assert.AreEqual("var ", x, d);

        v = new double[] {7.0, 25.0, 21.69};
        d = StatsLib.var(v);
        x = 91.79203333;
        Assert.AreEqual("var ", x, d);

        v = new double[] {1345,1301,1368,1322,1310,1370,1318,1350,1303,1299};
        d = StatsLib.var(v);
        x = 754.2666667;
        Assert.AreEqual("var ", x, d);
    }

    public void TestVarp() {
        double[] v = null;
        double d, x = 0;

        v = new double[] {3.50, 5.00, 7.23, 2.99};
        d = StatsLib.varp(v);
        x = 2.71335;
        Assert.AreEqual("varp ", x, d);

        v = new double[] {34.5, 2.0, 8.9, -4.0};
        d = StatsLib.varp(v);
        x = 215.2425;
        Assert.AreEqual("varp ", x, d);

        v = new double[] {7.0, 25.0, 21.69};
        d = StatsLib.varp(v);
        x = 61.19468889;
        Assert.AreEqual("varp ", x, d);

        v = new double[] {1345,1301,1368,1322,1310,1370,1318,1350,1303,1299};
        d = StatsLib.varp(v);
        x = 678.84;
        Assert.AreEqual("varp ", x, d);
    }
}

