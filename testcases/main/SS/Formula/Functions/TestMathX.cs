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
 * Created on May 23, 2005
 *
 */
namespace NPOI.SS.Formula.functions;

using NPOI.SS.Formula.functions.XYNumericFunction.Accumulator;


/**
 * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
 *  
 */
public class TestMathX : AbstractNumericTestCase {

    public void TestAcosh() {
        double d = 0;

        d = MathX.acosh(0);
        Assert.IsTrue("Acosh 0 is NaN", Double.IsNaN(d));

        d = MathX.acosh(1);
        Assert.AreEqual("Acosh 1 ", 0, d);

        d = MathX.acosh(-1);
        Assert.IsTrue("Acosh -1 is NaN", Double.IsNaN(d));

        d = MathX.acosh(100);
        Assert.AreEqual("Acosh 100 ", 5.298292366d, d);

        d = MathX.acosh(101.001);
        Assert.AreEqual("Acosh 101.001 ", 5.308253091d, d);

        d = MathX.acosh(200000);
        Assert.AreEqual("Acosh 200000 ", 12.89921983d, d);

    }

    public void TestAsinh() {
        double d = 0;

        d = MathX.asinh(0);
        Assert.AreEqual("asinh 0", d, 0);

        d = MathX.asinh(1);
        Assert.AreEqual("asinh 1 ", 0.881373587, d);

        d = MathX.asinh(-1);
        Assert.AreEqual("asinh -1 ", -0.881373587, d);

        d = MathX.asinh(-100);
        Assert.AreEqual("asinh -100 ", -5.298342366, d);

        d = MathX.asinh(100);
        Assert.AreEqual("asinh 100 ", 5.298342366, d);

        d = MathX.asinh(200000);
        Assert.AreEqual("asinh 200000", 12.899219826096400, d);

        d = MathX.asinh(-200000);
        Assert.AreEqual("asinh -200000 ", -12.899223853137, d);

    }

    public void TestAtanh() {
        double d = 0;
        d = MathX.atanh(0);
        Assert.AreEqual("atanh 0", d, 0);

        d = MathX.atanh(1);
        Assert.AreEqual("atanh 1 ", Double.POSITIVE_INFINITY, d);

        d = MathX.atanh(-1);
        Assert.AreEqual("atanh -1 ", Double.NEGATIVE_INFINITY, d);

        d = MathX.atanh(-100);
        Assert.AreEqual("atanh -100 ", Double.NaN, d);

        d = MathX.atanh(100);
        Assert.AreEqual("atanh 100 ", Double.NaN, d);

        d = MathX.atanh(200000);
        Assert.AreEqual("atanh 200000", Double.NaN, d);

        d = MathX.atanh(-200000);
        Assert.AreEqual("atanh -200000 ", Double.NaN, d);

        d = MathX.atanh(0.1);
        Assert.AreEqual("atanh 0.1", 0.100335348, d);

        d = MathX.atanh(-0.1);
        Assert.AreEqual("atanh -0.1 ", -0.100335348, d);

    }

    public void TestCosh() {
        double d = 0;
        d = MathX.cosh(0);
        Assert.AreEqual("cosh 0", 1, d);

        d = MathX.cosh(1);
        Assert.AreEqual("cosh 1 ", 1.543080635, d);

        d = MathX.cosh(-1);
        Assert.AreEqual("cosh -1 ", 1.543080635, d);

        d = MathX.cosh(-100);
        Assert.AreEqual("cosh -100 ", 1.344058570908070E+43, d);

        d = MathX.cosh(100);
        Assert.AreEqual("cosh 100 ", 1.344058570908070E+43, d);

        d = MathX.cosh(15);
        Assert.AreEqual("cosh 15", 1634508.686, d);

        d = MathX.cosh(-15);
        Assert.AreEqual("cosh -15 ", 1634508.686, d);

        d = MathX.cosh(0.1);
        Assert.AreEqual("cosh 0.1", 1.005004168, d);

        d = MathX.cosh(-0.1);
        Assert.AreEqual("cosh -0.1 ", 1.005004168, d);

    }

    public void TestTanh() {
        double d = 0;
        d = MathX.tanh(0);
        Assert.AreEqual("tanh 0", 0, d);

        d = MathX.tanh(1);
        Assert.AreEqual("tanh 1 ", 0.761594156, d);

        d = MathX.tanh(-1);
        Assert.AreEqual("tanh -1 ", -0.761594156, d);

        d = MathX.tanh(-100);
        Assert.AreEqual("tanh -100 ", -1, d);

        d = MathX.tanh(100);
        Assert.AreEqual("tanh 100 ", 1, d);

        d = MathX.tanh(15);
        Assert.AreEqual("tanh 15", 1, d);

        d = MathX.tanh(-15);
        Assert.AreEqual("tanh -15 ", -1, d);

        d = MathX.tanh(0.1);
        Assert.AreEqual("tanh 0.1", 0.099667995, d);

        d = MathX.tanh(-0.1);
        Assert.AreEqual("tanh -0.1 ", -0.099667995, d);

    }

    public void TestMax() {
        double[] d = new double[100];
        d[0] = 1.1;     d[1] = 2.1;     d[2] = 3.1;     d[3] = 4.1; 
        d[4] = 5.1;     d[5] = 6.1;     d[6] = 7.1;     d[7] = 8.1;
        d[8] = 9.1;     d[9] = 10.1;    d[10] = 11.1;   d[11] = 12.1;
        d[12] = 13.1;   d[13] = 14.1;   d[14] = 15.1;   d[15] = 16.1;
        d[16] = 17.1;   d[17] = 18.1;   d[18] = 19.1;   d[19] = 20.1; 
        
        double m = MathX.max(d);
        Assert.AreEqual("Max ", 20.1, m);
        
        d = new double[1000];
        m = MathX.max(d);
        Assert.AreEqual("Max ", 0, m);
        
        d[0] = -1.1;     d[1] = 2.1;     d[2] = -3.1;     d[3] = 4.1; 
        d[4] = -5.1;     d[5] = 6.1;     d[6] = -7.1;     d[7] = 8.1;
        d[8] = -9.1;     d[9] = 10.1;    d[10] = -11.1;   d[11] = 12.1;
        d[12] = -13.1;   d[13] = 14.1;   d[14] = -15.1;   d[15] = 16.1;
        d[16] = -17.1;   d[17] = 18.1;   d[18] = -19.1;   d[19] = 20.1; 
        m = MathX.max(d);
        Assert.AreEqual("Max ", 20.1, m);
        
        d = new double[20];
        d[0] = -1.1;     d[1] = -2.1;     d[2] = -3.1;     d[3] = -4.1; 
        d[4] = -5.1;     d[5] = -6.1;     d[6] = -7.1;     d[7] = -8.1;
        d[8] = -9.1;     d[9] = -10.1;    d[10] = -11.1;   d[11] = -12.1;
        d[12] = -13.1;   d[13] = -14.1;   d[14] = -15.1;   d[15] = -16.1;
        d[16] = -17.1;   d[17] = -18.1;   d[18] = -19.1;   d[19] = -20.1; 
        m = MathX.max(d);
        Assert.AreEqual("Max ", -1.1, m);
        
    }

    public void TestMin() {
        double[] d = new double[100];
        d[0] = 1.1;     d[1] = 2.1;     d[2] = 3.1;     d[3] = 4.1; 
        d[4] = 5.1;     d[5] = 6.1;     d[6] = 7.1;     d[7] = 8.1;
        d[8] = 9.1;     d[9] = 10.1;    d[10] = 11.1;   d[11] = 12.1;
        d[12] = 13.1;   d[13] = 14.1;   d[14] = 15.1;   d[15] = 16.1;
        d[16] = 17.1;   d[17] = 18.1;   d[18] = 19.1;   d[19] = 20.1; 
        
        double m = MathX.min(d);
        Assert.AreEqual("Min ", 0, m);
        
        d = new double[20];
        d[0] = 1.1;     d[1] = 2.1;     d[2] = 3.1;     d[3] = 4.1; 
        d[4] = 5.1;     d[5] = 6.1;     d[6] = 7.1;     d[7] = 8.1;
        d[8] = 9.1;     d[9] = 10.1;    d[10] = 11.1;   d[11] = 12.1;
        d[12] = 13.1;   d[13] = 14.1;   d[14] = 15.1;   d[15] = 16.1;
        d[16] = 17.1;   d[17] = 18.1;   d[18] = 19.1;   d[19] = 20.1; 
        
        m = MathX.min(d);
        Assert.AreEqual("Min ", 1.1, m);
        
        d = new double[1000];
        m = MathX.min(d);
        Assert.AreEqual("Min ", 0, m);
        
        d[0] = -1.1;     d[1] = 2.1;     d[2] = -3.1;     d[3] = 4.1; 
        d[4] = -5.1;     d[5] = 6.1;     d[6] = -7.1;     d[7] = 8.1;
        d[8] = -9.1;     d[9] = 10.1;    d[10] = -11.1;   d[11] = 12.1;
        d[12] = -13.1;   d[13] = 14.1;   d[14] = -15.1;   d[15] = 16.1;
        d[16] = -17.1;   d[17] = 18.1;   d[18] = -19.1;   d[19] = 20.1; 
        m = MathX.min(d);
        Assert.AreEqual("Min ", -19.1, m);
        
        d = new double[20];
        d[0] = -1.1;     d[1] = -2.1;     d[2] = -3.1;     d[3] = -4.1; 
        d[4] = -5.1;     d[5] = -6.1;     d[6] = -7.1;     d[7] = -8.1;
        d[8] = -9.1;     d[9] = -10.1;    d[10] = -11.1;   d[11] = -12.1;
        d[12] = -13.1;   d[13] = -14.1;   d[14] = -15.1;   d[15] = -16.1;
        d[16] = -17.1;   d[17] = -18.1;   d[18] = -19.1;   d[19] = -20.1; 
        m = MathX.min(d);
        Assert.AreEqual("Min ", -20.1, m);
    }

    public void TestProduct() {
        double[] d = new double[100];
        d[0] = 1.1;     d[1] = 2.1;     d[2] = 3.1;     d[3] = 4.1; 
        d[4] = 5.1;     d[5] = 6.1;     d[6] = 7.1;     d[7] = 8.1;
        d[8] = 9.1;     d[9] = 10.1;    d[10] = 11.1;   d[11] = 12.1;
        d[12] = 13.1;   d[13] = 14.1;   d[14] = 15.1;   d[15] = 16.1;
        d[16] = 17.1;   d[17] = 18.1;   d[18] = 19.1;   d[19] = 20.1; 
        
        double m = MathX.min(d);
        Assert.AreEqual("Min ", 0, m);
        
        d = new double[20];
        d[0] = 1.1;     d[1] = 2.1;     d[2] = 3.1;     d[3] = 4.1; 
        d[4] = 5.1;     d[5] = 6.1;     d[6] = 7.1;     d[7] = 8.1;
        d[8] = 9.1;     d[9] = 10.1;    d[10] = 11.1;   d[11] = 12.1;
        d[12] = 13.1;   d[13] = 14.1;   d[14] = 15.1;   d[15] = 16.1;
        d[16] = 17.1;   d[17] = 18.1;   d[18] = 19.1;   d[19] = 20.1; 
        
        m = MathX.min(d);
        Assert.AreEqual("Min ", 1.1, m);
        
        d = new double[1000];
        m = MathX.min(d);
        Assert.AreEqual("Min ", 0, m);
        
        d[0] = -1.1;     d[1] = 2.1;     d[2] = -3.1;     d[3] = 4.1; 
        d[4] = -5.1;     d[5] = 6.1;     d[6] = -7.1;     d[7] = 8.1;
        d[8] = -9.1;     d[9] = 10.1;    d[10] = -11.1;   d[11] = 12.1;
        d[12] = -13.1;   d[13] = 14.1;   d[14] = -15.1;   d[15] = 16.1;
        d[16] = -17.1;   d[17] = 18.1;   d[18] = -19.1;   d[19] = 20.1; 
        m = MathX.min(d);
        Assert.AreEqual("Min ", -19.1, m);
        
        d = new double[20];
        d[0] = -1.1;     d[1] = -2.1;     d[2] = -3.1;     d[3] = -4.1; 
        d[4] = -5.1;     d[5] = -6.1;     d[6] = -7.1;     d[7] = -8.1;
        d[8] = -9.1;     d[9] = -10.1;    d[10] = -11.1;   d[11] = -12.1;
        d[12] = -13.1;   d[13] = -14.1;   d[14] = -15.1;   d[15] = -16.1;
        d[16] = -17.1;   d[17] = -18.1;   d[18] = -19.1;   d[19] = -20.1; 
        m = MathX.min(d);
        Assert.AreEqual("Min ", -20.1, m);
    }

    public void TestMod() {

        //example from Excel help
        Assert.AreEqual(1.0, MathX.mod(3, 2));
        Assert.AreEqual(1.0, MathX.mod(-3, 2));
        Assert.AreEqual(-1.0, MathX.mod(3, -2));
        Assert.AreEqual(-1.0, MathX.mod(-3, -2));

        Assert.AreEqual((double) 1.4, MathX.mod(3.4, 2));
        Assert.AreEqual((double) -1.4, MathX.mod(-3.4, -2));
        Assert.AreEqual((double) 0.6000000000000001, MathX.mod(-3.4, 2.0));// should actually be 0.6
        Assert.AreEqual((double) -0.6000000000000001, MathX.mod(3.4, -2.0));// should actually be -0.6

        // Bugzilla 50033
        Assert.AreEqual(1.0, MathX.mod(13, 12));
    }

    public void TestNChooseK() {
        int n=100;
        int k=50;
        double d = MathX.nChooseK(n, k);
        Assert.AreEqual("NChooseK ", 1.00891344545564E29, d);
        
        n = -1; k = 1;
        d = MathX.nChooseK(n, k);
        Assert.AreEqual("NChooseK ", Double.NaN, d);
        
        n = 1; k = -1;
        d = MathX.nChooseK(n, k);
        Assert.AreEqual("NChooseK ", Double.NaN, d);
        
        n = 0; k = 1;
        d = MathX.nChooseK(n, k);
        Assert.AreEqual("NChooseK ", Double.NaN, d);
        
        n = 1; k = 0;
        d = MathX.nChooseK(n, k);
        Assert.AreEqual("NChooseK ", 1, d);
        
        n = 10; k = 9;
        d = MathX.nChooseK(n, k);
        Assert.AreEqual("NChooseK ", 10, d);
        
        n = 10; k = 10;
        d = MathX.nChooseK(n, k);
        Assert.AreEqual("NChooseK ", 1, d);
        
        n = 10; k = 1;
        d = MathX.nChooseK(n, k);
        Assert.AreEqual("NChooseK ", 10, d);
        
        n = 1000; k = 1;
        d = MathX.nChooseK(n, k);
        Assert.AreEqual("NChooseK ", 1000, d); // awesome ;)
        
        n = 1000; k = 2;
        d = MathX.nChooseK(n, k);
        Assert.AreEqual("NChooseK ", 499500, d); // awesome ;)
        
        n = 13; k = 7;
        d = MathX.nChooseK(n, k);
        Assert.AreEqual("NChooseK ", 1716, d);
        
    }

    public void TestSign() {
        short minus = -1;
        short zero = 0;
        short plus = 1;
        double d = 0;
        
        
        Assert.AreEqual("Sign ", minus, MathX.sign(minus));
        Assert.AreEqual("Sign ", plus, MathX.sign(plus));
        Assert.AreEqual("Sign ", zero, MathX.sign(zero));
        
        d = 0;
        Assert.AreEqual("Sign ", zero, MathX.sign(d));
        
        d = -1.000001;
        Assert.AreEqual("Sign ", minus, MathX.sign(d));
        
        d = -.000001;
        Assert.AreEqual("Sign ", minus, MathX.sign(d));
        
        d = -1E-200;
        Assert.AreEqual("Sign ", minus, MathX.sign(d));
        
        d = Double.NEGATIVE_INFINITY;
        Assert.AreEqual("Sign ", minus, MathX.sign(d));
        
        d = -200.11;
        Assert.AreEqual("Sign ", minus, MathX.sign(d));
        
        d = -2000000000000.11;
        Assert.AreEqual("Sign ", minus, MathX.sign(d));
        
        d = 1.000001;
        Assert.AreEqual("Sign ", plus, MathX.sign(d));
        
        d = .000001;
        Assert.AreEqual("Sign ", plus, MathX.sign(d));
        
        d = 1E-200;
        Assert.AreEqual("Sign ", plus, MathX.sign(d));
        
        d = Double.POSITIVE_INFINITY;
        Assert.AreEqual("Sign ", plus, MathX.sign(d));
        
        d = 200.11;
        Assert.AreEqual("Sign ", plus, MathX.sign(d));
        
        d = 2000000000000.11;
        Assert.AreEqual("Sign ", plus, MathX.sign(d));
        
    }

    public void TestSinh() {
        double d = 0;
        d = MathX.sinh(0);
        Assert.AreEqual("sinh 0", 0, d);

        d = MathX.sinh(1);
        Assert.AreEqual("sinh 1 ", 1.175201194, d);

        d = MathX.sinh(-1);
        Assert.AreEqual("sinh -1 ", -1.175201194, d);

        d = MathX.sinh(-100);
        Assert.AreEqual("sinh -100 ", -1.344058570908070E+43, d);

        d = MathX.sinh(100);
        Assert.AreEqual("sinh 100 ", 1.344058570908070E+43, d);

        d = MathX.sinh(15);
        Assert.AreEqual("sinh 15", 1634508.686, d);

        d = MathX.sinh(-15);
        Assert.AreEqual("sinh -15 ", -1634508.686, d);

        d = MathX.sinh(0.1);
        Assert.AreEqual("sinh 0.1", 0.10016675, d);

        d = MathX.sinh(-0.1);
        Assert.AreEqual("sinh -0.1 ", -0.10016675, d);

    }

    public void TestSum() {
        double[] d = new double[100];
        d[0] = 1.1;     d[1] = 2.1;     d[2] = 3.1;     d[3] = 4.1; 
        d[4] = 5.1;     d[5] = 6.1;     d[6] = 7.1;     d[7] = 8.1;
        d[8] = 9.1;     d[9] = 10.1;    d[10] = 11.1;   d[11] = 12.1;
        d[12] = 13.1;   d[13] = 14.1;   d[14] = 15.1;   d[15] = 16.1;
        d[16] = 17.1;   d[17] = 18.1;   d[18] = 19.1;   d[19] = 20.1; 
        
        double s = MathX.sum(d);
        Assert.AreEqual("Sum ", 212, s);
        
        d = new double[1000];
        s = MathX.sum(d);
        Assert.AreEqual("Sum ", 0, s);
        
        d[0] = -1.1;     d[1] = 2.1;     d[2] = -3.1;     d[3] = 4.1; 
        d[4] = -5.1;     d[5] = 6.1;     d[6] = -7.1;     d[7] = 8.1;
        d[8] = -9.1;     d[9] = 10.1;    d[10] = -11.1;   d[11] = 12.1;
        d[12] = -13.1;   d[13] = 14.1;   d[14] = -15.1;   d[15] = 16.1;
        d[16] = -17.1;   d[17] = 18.1;   d[18] = -19.1;   d[19] = 20.1; 
        s = MathX.sum(d);
        Assert.AreEqual("Sum ", 10, s);
        
        d[0] = -1.1;     d[1] = -2.1;     d[2] = -3.1;     d[3] = -4.1; 
        d[4] = -5.1;     d[5] = -6.1;     d[6] = -7.1;     d[7] = -8.1;
        d[8] = -9.1;     d[9] = -10.1;    d[10] = -11.1;   d[11] = -12.1;
        d[12] = -13.1;   d[13] = -14.1;   d[14] = -15.1;   d[15] = -16.1;
        d[16] = -17.1;   d[17] = -18.1;   d[18] = -19.1;   d[19] = -20.1; 
        s = MathX.sum(d);
        Assert.AreEqual("Sum ", -212, s);
        
    }

    public void TestSumsq() {
        double[] d = new double[100];
        d[0] = 1.1;     d[1] = 2.1;     d[2] = 3.1;     d[3] = 4.1; 
        d[4] = 5.1;     d[5] = 6.1;     d[6] = 7.1;     d[7] = 8.1;
        d[8] = 9.1;     d[9] = 10.1;    d[10] = 11.1;   d[11] = 12.1;
        d[12] = 13.1;   d[13] = 14.1;   d[14] = 15.1;   d[15] = 16.1;
        d[16] = 17.1;   d[17] = 18.1;   d[18] = 19.1;   d[19] = 20.1; 
        
        double s = MathX.sumsq(d);
        Assert.AreEqual("Sumsq ", 2912.2, s);
        
        d = new double[1000];
        s = MathX.sumsq(d);
        Assert.AreEqual("Sumsq ", 0, s);
        
        d[0] = -1.1;     d[1] = 2.1;     d[2] = -3.1;     d[3] = 4.1; 
        d[4] = -5.1;     d[5] = 6.1;     d[6] = -7.1;     d[7] = 8.1;
        d[8] = -9.1;     d[9] = 10.1;    d[10] = -11.1;   d[11] = 12.1;
        d[12] = -13.1;   d[13] = 14.1;   d[14] = -15.1;   d[15] = 16.1;
        d[16] = -17.1;   d[17] = 18.1;   d[18] = -19.1;   d[19] = 20.1; 
        s = MathX.sumsq(d);
        Assert.AreEqual("Sumsq ", 2912.2, s);
        
        d[0] = -1.1;     d[1] = -2.1;     d[2] = -3.1;     d[3] = -4.1; 
        d[4] = -5.1;     d[5] = -6.1;     d[6] = -7.1;     d[7] = -8.1;
        d[8] = -9.1;     d[9] = -10.1;    d[10] = -11.1;   d[11] = -12.1;
        d[12] = -13.1;   d[13] = -14.1;   d[14] = -15.1;   d[15] = -16.1;
        d[16] = -17.1;   d[17] = -18.1;   d[18] = -19.1;   d[19] = -20.1; 
        s = MathX.sumsq(d);
        Assert.AreEqual("Sumsq ", 2912.2, s);
    }

    public void TestFactorial() {
        int n = 0;
        double s = 0;
        
        n = 0;
        s = MathX.factorial(n);
        Assert.AreEqual("Factorial ", 1, s);
        
        n = 1;
        s = MathX.factorial(n);
        Assert.AreEqual("Factorial ", 1, s);
        
        n = 10;
        s = MathX.factorial(n);
        Assert.AreEqual("Factorial ", 3628800, s);
        
        n = 99;
        s = MathX.factorial(n);
        Assert.AreEqual("Factorial ", 9.33262154439E+155, s);
        
        n = -1;
        s = MathX.factorial(n);
        Assert.AreEqual("Factorial ", Double.NaN, s);
        
        n = Int32.MaxValue;
        s = MathX.factorial(n);
        Assert.AreEqual("Factorial ", Double.POSITIVE_INFINITY, s);
    }

    public void TestSumx2my2() {
        double[] xarr = null;
        double[] yarr = null;
        
        xarr = new double[]{1, 2, 3, 4, 5, 6, 7, 8, 9, 10};
        yarr = new double[]{0, 1, 2, 3, 4, 5, 6, 7, 8, 9};
        ConfirmSumx2my2(xarr, yarr, 100);
        
        xarr = new double[]{-1, -2, -3, -4, -5, -6, -7, -8, -9, -10};
        yarr = new double[]{0, 1, 2, 3, 4, 5, 6, 7, 8, 9};
        ConfirmSumx2my2(xarr, yarr, 100);
        
        xarr = new double[]{0, 1, 2, 3, 4, 5, 6, 7, 8, 9};
        yarr = new double[]{1, 2, 3, 4, 5, 6, 7, 8, 9, 10};
        ConfirmSumx2my2(xarr, yarr, -100);
        
        xarr = new double[]{10};
        yarr = new double[]{9};
        ConfirmSumx2my2(xarr, yarr, 19);
        
        xarr = new double[]{1, 2, 3, 4, 5, 6, 7, 8, 9, 10};
        yarr = new double[]{1, 2, 3, 4, 5, 6, 7, 8, 9, 10};
        ConfirmSumx2my2(xarr, yarr, 0);
    }

    public void TestSumx2py2() {
        double[] xarr = null;
        double[] yarr = null;
        
        xarr = new double[]{1, 2, 3, 4, 5, 6, 7, 8, 9, 10};
        yarr = new double[]{0, 1, 2, 3, 4, 5, 6, 7, 8, 9};
        ConfirmSumx2py2(xarr, yarr, 670);
        
        xarr = new double[]{-1, -2, -3, -4, -5, -6, -7, -8, -9, -10};
        yarr = new double[]{0, 1, 2, 3, 4, 5, 6, 7, 8, 9};
        ConfirmSumx2py2(xarr, yarr, 670);
        
        xarr = new double[]{0, 1, 2, 3, 4, 5, 6, 7, 8, 9};
        yarr = new double[]{1, 2, 3, 4, 5, 6, 7, 8, 9, 10};
        ConfirmSumx2py2(xarr, yarr, 670);
        
        xarr = new double[]{10};
        yarr = new double[]{9};
        ConfirmSumx2py2(xarr, yarr, 181);
        
        xarr = new double[]{1, 2, 3, 4, 5, 6, 7, 8, 9, 10};
        yarr = new double[]{1, 2, 3, 4, 5, 6, 7, 8, 9, 10};
        ConfirmSumx2py2(xarr, yarr, 770);
    }

    public void TestSumxmy2() {
        double[] xarr = null;
        double[] yarr = null;
        
        xarr = new double[]{1, 2, 3, 4, 5, 6, 7, 8, 9, 10};
        yarr = new double[]{0, 1, 2, 3, 4, 5, 6, 7, 8, 9};
        ConfirmSumxmy2(xarr, yarr, 10);
        
        xarr = new double[]{-1, -2, -3, -4, -5, -6, -7, -8, -9, -10};
        yarr = new double[]{0, 1, 2, 3, 4, 5, 6, 7, 8, 9};
        ConfirmSumxmy2(xarr, yarr, 1330);
        
        xarr = new double[]{0, 1, 2, 3, 4, 5, 6, 7, 8, 9};
        yarr = new double[]{1, 2, 3, 4, 5, 6, 7, 8, 9, 10};
        ConfirmSumxmy2(xarr, yarr, 10);
        
        xarr = new double[]{10};
        yarr = new double[]{9};
        ConfirmSumxmy2(xarr, yarr, 1);
        
        xarr = new double[]{1, 2, 3, 4, 5, 6, 7, 8, 9, 10};
        yarr = new double[]{1, 2, 3, 4, 5, 6, 7, 8, 9, 10};
        ConfirmSumxmy2(xarr, yarr, 0);
    }

    private static void ConfirmSumx2my2(double[] xarr, double[] yarr, double expectedResult) {
        ConfirmXY(new Sumx2my2().CreateAccumulator(), xarr, yarr, expectedResult);
    }
    private static void ConfirmSumx2py2(double[] xarr, double[] yarr, double expectedResult) {
        ConfirmXY(new Sumx2py2().CreateAccumulator(), xarr, yarr, expectedResult);
    }
    private static void ConfirmSumxmy2(double[] xarr, double[] yarr, double expectedResult) {
        ConfirmXY(new Sumxmy2().CreateAccumulator(), xarr, yarr, expectedResult);
    }

    private static void ConfirmXY(Accumulator acc, double[] xarr, double[] yarr,
            double expectedResult) {
        double result = 0.0;
        for (int i = 0; i < xarr.Length; i++) {
            result += acc.accumulate(xarr[i], yarr[i]);
        }
        Assert.AreEqual(expectedResult, result, 0.0);
    }
    
    public void TestRound() {
        double d = 0;
        int p = 0;
        
        d = 0; p = 0;
        Assert.AreEqual("round ", 0, MathX.round(d, p));
        
        d = 10; p = 0;
        Assert.AreEqual("round ", 10, MathX.round(d, p));
        
        d = 123.23; p = 0;
        Assert.AreEqual("round ", 123, MathX.round(d, p));
        
        d = -123.23; p = 0;
        Assert.AreEqual("round ", -123, MathX.round(d, p));
        
        d = 123.12; p = 2;
        Assert.AreEqual("round ", 123.12, MathX.round(d, p));
        
        d = 88.123459; p = 5;
        Assert.AreEqual("round ", 88.12346, MathX.round(d, p));
        
        d = 0; p = 2;
        Assert.AreEqual("round ", 0, MathX.round(d, p));
        
        d = 0; p = -1;
        Assert.AreEqual("round ", 0, MathX.round(d, p));
        
        d = 0.01; p = -1;
        Assert.AreEqual("round ", 0, MathX.round(d, p));

        d = 123.12; p = -2;
        Assert.AreEqual("round ", 100, MathX.round(d, p));
        
        d = 88.123459; p = -3;
        Assert.AreEqual("round ", 0, MathX.round(d, p));
        
        d = 49.00000001; p = -1;
        Assert.AreEqual("round ", 50, MathX.round(d, p));
        
        d = 149.999999; p = -2;
        Assert.AreEqual("round ", 100, MathX.round(d, p));
        
        d = 150.0; p = -2;
        Assert.AreEqual("round ", 200, MathX.round(d, p));

        d = 2162.615d; p = 2;
        Assert.AreEqual("round ", 2162.62d, MathX.round(d, p));
    }

    public void TestRoundDown() {
        double d = 0;
        int p = 0;
        
        d = 0; p = 0;
        Assert.AreEqual("roundDown ", 0, MathX.roundDown(d, p));
        
        d = 10; p = 0;
        Assert.AreEqual("roundDown ", 10, MathX.roundDown(d, p));
        
        d = 123.99; p = 0;
        Assert.AreEqual("roundDown ", 123, MathX.roundDown(d, p));
        
        d = -123.99; p = 0;
        Assert.AreEqual("roundDown ", -123, MathX.roundDown(d, p));
        
        d = 123.99; p = 2;
        Assert.AreEqual("roundDown ", 123.99, MathX.roundDown(d, p));
        
        d = 88.123459; p = 5;
        Assert.AreEqual("roundDown ", 88.12345, MathX.roundDown(d, p));
        
        d = 0; p = 2;
        Assert.AreEqual("roundDown ", 0, MathX.roundDown(d, p));
        
        d = 0; p = -1;
        Assert.AreEqual("roundDown ", 0, MathX.roundDown(d, p));
        
        d = 0.01; p = -1;
        Assert.AreEqual("roundDown ", 0, MathX.roundDown(d, p));

        d = 199.12; p = -2;
        Assert.AreEqual("roundDown ", 100, MathX.roundDown(d, p));
        
        d = 88.123459; p = -3;
        Assert.AreEqual("roundDown ", 0, MathX.roundDown(d, p));
        
        d = 99.00000001; p = -1;
        Assert.AreEqual("roundDown ", 90, MathX.roundDown(d, p));
        
        d = 100.00001; p = -2;
        Assert.AreEqual("roundDown ", 100, MathX.roundDown(d, p));
        
        d = 150.0; p = -2;
        Assert.AreEqual("roundDown ", 100, MathX.roundDown(d, p));
    }

    public void TestRoundUp() {
        double d = 0;
        int p = 0;
        
        d = 0; p = 0;
        Assert.AreEqual("roundUp ", 0, MathX.roundUp(d, p));
        
        d = 10; p = 0;
        Assert.AreEqual("roundUp ", 10, MathX.roundUp(d, p));
        
        d = 123.23; p = 0;
        Assert.AreEqual("roundUp ", 124, MathX.roundUp(d, p));
        
        d = -123.23; p = 0;
        Assert.AreEqual("roundUp ", -124, MathX.roundUp(d, p));
        
        d = 123.12; p = 2;
        Assert.AreEqual("roundUp ", 123.12, MathX.roundUp(d, p));
        
        d = 88.123459; p = 5;
        Assert.AreEqual("roundUp ", 88.12346, MathX.roundUp(d, p));
        
        d = 0; p = 2;
        Assert.AreEqual("roundUp ", 0, MathX.roundUp(d, p));
        
        d = 0; p = -1;
        Assert.AreEqual("roundUp ", 0, MathX.roundUp(d, p));
        
        d = 0.01; p = -1;
        Assert.AreEqual("roundUp ", 10, MathX.roundUp(d, p));

        d = 123.12; p = -2;
        Assert.AreEqual("roundUp ", 200, MathX.roundUp(d, p));
        
        d = 88.123459; p = -3;
        Assert.AreEqual("roundUp ", 1000, MathX.roundUp(d, p));
        
        d = 49.00000001; p = -1;
        Assert.AreEqual("roundUp ", 50, MathX.roundUp(d, p));
        
        d = 149.999999; p = -2;
        Assert.AreEqual("roundUp ", 200, MathX.roundUp(d, p));
        
        d = 150.0; p = -2;
        Assert.AreEqual("roundUp ", 200, MathX.roundUp(d, p));
    }

    public void TestCeiling() {
        double d = 0;
        double s = 0;
        
        d = 0; s = 0;
        Assert.AreEqual("ceiling ", 0, MathX.ceiling(d, s));
        
        d = 1; s = 0;
        Assert.AreEqual("ceiling ", 0, MathX.ceiling(d, s));
        
        d = 0; s = 1;
        Assert.AreEqual("ceiling ", 0, MathX.ceiling(d, s));
        
        d = -1; s = 0;
        Assert.AreEqual("ceiling ", 0, MathX.ceiling(d, s));
        
        d = 0; s = -1;
        Assert.AreEqual("ceiling ", 0, MathX.ceiling(d, s));
        
        d = 10; s = 1.11;
        Assert.AreEqual("ceiling ", 11.1, MathX.ceiling(d, s));
        
        d = 11.12333; s = 0.03499;
        Assert.AreEqual("ceiling ", 11.12682, MathX.ceiling(d, s));
        
        d = -11.12333; s = 0.03499;
        Assert.AreEqual("ceiling ", Double.NaN, MathX.ceiling(d, s));
        
        d = 11.12333; s = -0.03499;
        Assert.AreEqual("ceiling ", Double.NaN, MathX.ceiling(d, s));
        
        d = -11.12333; s = -0.03499;
        Assert.AreEqual("ceiling ", -11.12682, MathX.ceiling(d, s));
        
        d = 100; s = 0.001;
        Assert.AreEqual("ceiling ", 100, MathX.ceiling(d, s));
        
        d = -0.001; s = -9.99;
        Assert.AreEqual("ceiling ", -9.99, MathX.ceiling(d, s));
        
        d = 4.42; s = 0.05;
        Assert.AreEqual("ceiling ", 4.45, MathX.ceiling(d, s));
        
        d = 0.05; s = 4.42;
        Assert.AreEqual("ceiling ", 4.42, MathX.ceiling(d, s));
        
        d = 0.6666; s = 3.33;
        Assert.AreEqual("ceiling ", 3.33, MathX.ceiling(d, s));
        
        d = 2d/3; s = 3.33;
        Assert.AreEqual("ceiling ", 3.33, MathX.ceiling(d, s));
    }

    public void TestFloor() {
        double d = 0;
        double s = 0;
        
        d = 0; s = 0;
        Assert.AreEqual("floor ", 0, MathX.floor(d, s));
        
        d = 1; s = 0;
        Assert.AreEqual("floor ", Double.NaN, MathX.floor(d, s));
        
        d = 0; s = 1;
        Assert.AreEqual("floor ", 0, MathX.floor(d, s));
        
        d = -1; s = 0;
        Assert.AreEqual("floor ", Double.NaN, MathX.floor(d, s));
        
        d = 0; s = -1;
        Assert.AreEqual("floor ", 0, MathX.floor(d, s));
        
        d = 10; s = 1.11;
        Assert.AreEqual("floor ", 9.99, MathX.floor(d, s));
        
        d = 11.12333; s = 0.03499;
        Assert.AreEqual("floor ", 11.09183, MathX.floor(d, s));
        
        d = -11.12333; s = 0.03499;
        Assert.AreEqual("floor ", Double.NaN, MathX.floor(d, s));
        
        d = 11.12333; s = -0.03499;
        Assert.AreEqual("floor ", Double.NaN, MathX.floor(d, s));
        
        d = -11.12333; s = -0.03499;
        Assert.AreEqual("floor ", -11.09183, MathX.floor(d, s));
        
        d = 100; s = 0.001;
        Assert.AreEqual("floor ", 100, MathX.floor(d, s));
        
        d = -0.001; s = -9.99;
        Assert.AreEqual("floor ", 0, MathX.floor(d, s));
        
        d = 4.42; s = 0.05;
        Assert.AreEqual("floor ", 4.4, MathX.floor(d, s));
        
        d = 0.05; s = 4.42;
        Assert.AreEqual("floor ", 0, MathX.floor(d, s));
        
        d = 0.6666; s = 3.33;
        Assert.AreEqual("floor ", 0, MathX.floor(d, s));
        
        d = 2d/3; s = 3.33;
        Assert.AreEqual("floor ", 0, MathX.floor(d, s));
    }

}

