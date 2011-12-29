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
namespace TestCases.SS.Formula.Functions
{

    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.SS.Formula.Functions;


    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *  
     */
    [TestClass]
    public class TestMathX : AbstractNumericTestCase
    {
        [TestMethod]
        public void TestAcosh()
        {
            double d = 0;

            d = MathX.acosh(0);
            Assert.IsTrue(Double.IsNaN(d), "Acosh 0 is NaN");

            d = MathX.acosh(1);
            Assert.AreEqual(0, d, "Acosh 1 ");

            d = MathX.acosh(-1);
            Assert.IsTrue(Double.IsNaN(d), "Acosh -1 is NaN");

            d = MathX.acosh(100);
            Assert.AreEqual( 5.298292366d, d,"Acosh 100 ");

            d = MathX.acosh(101.001);
            Assert.AreEqual(5.308253091d, d, "Acosh 101.001 ");

            d = MathX.acosh(200000);
            Assert.AreEqual(12.89921983d, d, "Acosh 200000 ");

        }
        [TestMethod]
        public void TestAsinh()
        {
            double d = 0;

            d = MathX.asinh(0);
            Assert.AreEqual(d, 0, "asinh 0");

            d = MathX.asinh(1);
            Assert.AreEqual(0.881373587, d, "asinh 1 ");

            d = MathX.asinh(-1);
            Assert.AreEqual(-0.881373587, d, "asinh -1 ");

            d = MathX.asinh(-100);
            Assert.AreEqual(-5.298342366, d, "asinh -100 ");

            d = MathX.asinh(100);
            Assert.AreEqual(5.298342366, d, "asinh 100 ");

            d = MathX.asinh(200000);
            Assert.AreEqual(12.899219826096400, d, "asinh 200000");

            d = MathX.asinh(-200000);
            Assert.AreEqual(-12.899223853137, d, "asinh -200000 ");

        }
        [TestMethod]
        public void TestAtanh()
        {
            double d = 0;
            d = MathX.atanh(0);
            Assert.AreEqual( d, 0,"atanh 0");

            d = MathX.atanh(1);
            Assert.AreEqual( Double.PositiveInfinity, d,"atanh 1 ");

            d = MathX.atanh(-1);
            Assert.AreEqual( Double.NegativeInfinity, d,"atanh -1 ");

            d = MathX.atanh(-100);
            Assert.AreEqual( Double.NaN, d,"atanh -100 ");
            
            d = MathX.atanh(100);
            Assert.AreEqual( Double.NaN, d,"atanh 100 ");

            d = MathX.atanh(200000);
            Assert.AreEqual( Double.NaN, d,"atanh 200000");

            d = MathX.atanh(-200000);
            Assert.AreEqual( Double.NaN, d,"atanh -200000 ");

            d = MathX.atanh(0.1);
            Assert.AreEqual( 0.100335348, d,"atanh 0.1");

            d = MathX.atanh(-0.1);
            Assert.AreEqual( -0.100335348, d,"atanh -0.1 ");

        }
        [TestMethod]
        public void TestCosh()
        {
            double d = 0;
            d = MathX.cosh(0);
            Assert.AreEqual( 1, d,"cosh 0");

            d = MathX.cosh(1);
            Assert.AreEqual( 1.543080635, d,"cosh 1 ");

            d = MathX.cosh(-1);
            Assert.AreEqual( 1.543080635, d,"cosh -1 ");

            d = MathX.cosh(-100);
            Assert.AreEqual( 1.344058570908070E+43, d,"cosh -100 ");

            d = MathX.cosh(100);
            Assert.AreEqual( 1.344058570908070E+43, d,"cosh 100 ");

            d = MathX.cosh(15);
            Assert.AreEqual( 1634508.686, d,"cosh 15");

            d = MathX.cosh(-15);
            Assert.AreEqual( 1634508.686, d,"cosh -15 ");

            d = MathX.cosh(0.1);
            Assert.AreEqual( 1.005004168, d,"cosh 0.1");

            d = MathX.cosh(-0.1);
            Assert.AreEqual( 1.005004168, d,"cosh -0.1 ");

        }
        [TestMethod]
        public void TestTanh()
        {
            double d = 0;
            d = MathX.tanh(0);
            Assert.AreEqual( 0, d,"tanh 0");

            d = MathX.tanh(1);
            Assert.AreEqual( 0.761594156, d,"tanh 1 ");

            d = MathX.tanh(-1);
            Assert.AreEqual( -0.761594156, d,"tanh -1 ");

            d = MathX.tanh(-100);
            Assert.AreEqual( -1, d,"tanh -100 ");

            d = MathX.tanh(100);
            Assert.AreEqual( 1, d,"tanh 100 ");

            d = MathX.tanh(15);
            Assert.AreEqual( 1, d,"tanh 15");

            d = MathX.tanh(-15);
            Assert.AreEqual( -1, d,"tanh -15 ");

            d = MathX.tanh(0.1);
            Assert.AreEqual( 0.099667995, d,"tanh 0.1");

            d = MathX.tanh(-0.1);
            Assert.AreEqual( -0.099667995, d,"tanh -0.1 ");

        }
        [TestMethod]
        public void TestMax()
        {
            double[] d = new double[100];
            d[0] = 1.1; d[1] = 2.1; d[2] = 3.1; d[3] = 4.1;
            d[4] = 5.1; d[5] = 6.1; d[6] = 7.1; d[7] = 8.1;
            d[8] = 9.1; d[9] = 10.1; d[10] = 11.1; d[11] = 12.1;
            d[12] = 13.1; d[13] = 14.1; d[14] = 15.1; d[15] = 16.1;
            d[16] = 17.1; d[17] = 18.1; d[18] = 19.1; d[19] = 20.1;

            double m = MathX.max(d);
            Assert.AreEqual( 20.1, m,"Max ");

            d = new double[1000];
            m = MathX.max(d);
            Assert.AreEqual( 0, m,"Max ");

            d[0] = -1.1; d[1] = 2.1; d[2] = -3.1; d[3] = 4.1;
            d[4] = -5.1; d[5] = 6.1; d[6] = -7.1; d[7] = 8.1;
            d[8] = -9.1; d[9] = 10.1; d[10] = -11.1; d[11] = 12.1;
            d[12] = -13.1; d[13] = 14.1; d[14] = -15.1; d[15] = 16.1;
            d[16] = -17.1; d[17] = 18.1; d[18] = -19.1; d[19] = 20.1;
            m = MathX.max(d);
            Assert.AreEqual( 20.1, m,"Max ");

            d = new double[20];
            d[0] = -1.1; d[1] = -2.1; d[2] = -3.1; d[3] = -4.1;
            d[4] = -5.1; d[5] = -6.1; d[6] = -7.1; d[7] = -8.1;
            d[8] = -9.1; d[9] = -10.1; d[10] = -11.1; d[11] = -12.1;
            d[12] = -13.1; d[13] = -14.1; d[14] = -15.1; d[15] = -16.1;
            d[16] = -17.1; d[17] = -18.1; d[18] = -19.1; d[19] = -20.1;
            m = MathX.max(d);
            Assert.AreEqual( -1.1, m,"Max ");

        }
        [TestMethod]
        public void TestMin()
        {
            double[] d = new double[100];
            d[0] = 1.1; d[1] = 2.1; d[2] = 3.1; d[3] = 4.1;
            d[4] = 5.1; d[5] = 6.1; d[6] = 7.1; d[7] = 8.1;
            d[8] = 9.1; d[9] = 10.1; d[10] = 11.1; d[11] = 12.1;
            d[12] = 13.1; d[13] = 14.1; d[14] = 15.1; d[15] = 16.1;
            d[16] = 17.1; d[17] = 18.1; d[18] = 19.1; d[19] = 20.1;

            double m = MathX.min(d);
            Assert.AreEqual( 0, m,"Min ");

            d = new double[20];
            d[0] = 1.1; d[1] = 2.1; d[2] = 3.1; d[3] = 4.1;
            d[4] = 5.1; d[5] = 6.1; d[6] = 7.1; d[7] = 8.1;
            d[8] = 9.1; d[9] = 10.1; d[10] = 11.1; d[11] = 12.1;
            d[12] = 13.1; d[13] = 14.1; d[14] = 15.1; d[15] = 16.1;
            d[16] = 17.1; d[17] = 18.1; d[18] = 19.1; d[19] = 20.1;

            m = MathX.min(d);
            Assert.AreEqual( 1.1, m,"Min ");

            d = new double[1000];
            m = MathX.min(d);
            Assert.AreEqual( 0, m,"Min ");

            d[0] = -1.1; d[1] = 2.1; d[2] = -3.1; d[3] = 4.1;
            d[4] = -5.1; d[5] = 6.1; d[6] = -7.1; d[7] = 8.1;
            d[8] = -9.1; d[9] = 10.1; d[10] = -11.1; d[11] = 12.1;
            d[12] = -13.1; d[13] = 14.1; d[14] = -15.1; d[15] = 16.1;
            d[16] = -17.1; d[17] = 18.1; d[18] = -19.1; d[19] = 20.1;
            m = MathX.min(d);
            Assert.AreEqual( -19.1, m,"Min ");

            d = new double[20];
            d[0] = -1.1; d[1] = -2.1; d[2] = -3.1; d[3] = -4.1;
            d[4] = -5.1; d[5] = -6.1; d[6] = -7.1; d[7] = -8.1;
            d[8] = -9.1; d[9] = -10.1; d[10] = -11.1; d[11] = -12.1;
            d[12] = -13.1; d[13] = -14.1; d[14] = -15.1; d[15] = -16.1;
            d[16] = -17.1; d[17] = -18.1; d[18] = -19.1; d[19] = -20.1;
            m = MathX.min(d);
            Assert.AreEqual( -20.1, m,"Min ");
        }
        [TestMethod]
        public void TestProduct()
        {
            double[] d = new double[100];
            d[0] = 1.1; d[1] = 2.1; d[2] = 3.1; d[3] = 4.1;
            d[4] = 5.1; d[5] = 6.1; d[6] = 7.1; d[7] = 8.1;
            d[8] = 9.1; d[9] = 10.1; d[10] = 11.1; d[11] = 12.1;
            d[12] = 13.1; d[13] = 14.1; d[14] = 15.1; d[15] = 16.1;
            d[16] = 17.1; d[17] = 18.1; d[18] = 19.1; d[19] = 20.1;

            double m = MathX.min(d);
            Assert.AreEqual( 0, m,"Min ");

            d = new double[20];
            d[0] = 1.1; d[1] = 2.1; d[2] = 3.1; d[3] = 4.1;
            d[4] = 5.1; d[5] = 6.1; d[6] = 7.1; d[7] = 8.1;
            d[8] = 9.1; d[9] = 10.1; d[10] = 11.1; d[11] = 12.1;
            d[12] = 13.1; d[13] = 14.1; d[14] = 15.1; d[15] = 16.1;
            d[16] = 17.1; d[17] = 18.1; d[18] = 19.1; d[19] = 20.1;

            m = MathX.min(d);
            Assert.AreEqual( 1.1, m,"Min ");

            d = new double[1000];
            m = MathX.min(d);
            Assert.AreEqual( 0, m,"Min ");

            d[0] = -1.1; d[1] = 2.1; d[2] = -3.1; d[3] = 4.1;
            d[4] = -5.1; d[5] = 6.1; d[6] = -7.1; d[7] = 8.1;
            d[8] = -9.1; d[9] = 10.1; d[10] = -11.1; d[11] = 12.1;
            d[12] = -13.1; d[13] = 14.1; d[14] = -15.1; d[15] = 16.1;
            d[16] = -17.1; d[17] = 18.1; d[18] = -19.1; d[19] = 20.1;
            m = MathX.min(d);
            Assert.AreEqual( -19.1, m,"Min ");

            d = new double[20];
            d[0] = -1.1; d[1] = -2.1; d[2] = -3.1; d[3] = -4.1;
            d[4] = -5.1; d[5] = -6.1; d[6] = -7.1; d[7] = -8.1;
            d[8] = -9.1; d[9] = -10.1; d[10] = -11.1; d[11] = -12.1;
            d[12] = -13.1; d[13] = -14.1; d[14] = -15.1; d[15] = -16.1;
            d[16] = -17.1; d[17] = -18.1; d[18] = -19.1; d[19] = -20.1;
            m = MathX.min(d);
            Assert.AreEqual( -20.1, m,"Min ");
        }
        [TestMethod]
        public void TestMod()
        {

            //example from Excel help
            Assert.AreEqual(1.0, MathX.mod(3, 2));
            Assert.AreEqual(1.0, MathX.mod(-3, 2));
            Assert.AreEqual(-1.0, MathX.mod(3, -2));
            Assert.AreEqual(-1.0, MathX.mod(-3, -2));

            Assert.AreEqual((double)1.4, MathX.mod(3.4, 2));
            Assert.AreEqual((double)-1.4, MathX.mod(-3.4, -2));
            Assert.AreEqual((double)0.6000000000000001, MathX.mod(-3.4, 2.0));// should actually be 0.6
            Assert.AreEqual((double)-0.6000000000000001, MathX.mod(3.4, -2.0));// should actually be -0.6

            // Bugzilla 50033
            Assert.AreEqual(1.0, MathX.mod(13, 12));
        }
        [TestMethod]
        public void TestNChooseK()
        {
            int n = 100;
            int k = 50;
            double d = MathX.nChooseK(n, k);
            Assert.AreEqual( 1.00891344545564E29, d,"NChooseK ");

            n = -1; k = 1;
            d = MathX.nChooseK(n, k);
            Assert.AreEqual( Double.NaN, d,"NChooseK ");

            n = 1; k = -1;
            d = MathX.nChooseK(n, k);
            Assert.AreEqual( Double.NaN, d,"NChooseK ");

            n = 0; k = 1;
            d = MathX.nChooseK(n, k);
            Assert.AreEqual( Double.NaN, d,"NChooseK ");

            n = 1; k = 0;
            d = MathX.nChooseK(n, k);
            Assert.AreEqual( 1, d,"NChooseK ");

            n = 10; k = 9;
            d = MathX.nChooseK(n, k);
            Assert.AreEqual( 10, d,"NChooseK ");

            n = 10; k = 10;
            d = MathX.nChooseK(n, k);
            Assert.AreEqual( 1, d,"NChooseK ");

            n = 10; k = 1;
            d = MathX.nChooseK(n, k);
            Assert.AreEqual( 10, d,"NChooseK ");

            n = 1000; k = 1;
            d = MathX.nChooseK(n, k);
            Assert.AreEqual( 1000, d,"NChooseK "); // awesome ;)

            n = 1000; k = 2;
            d = MathX.nChooseK(n, k);
            Assert.AreEqual( 499500, d,"NChooseK "); // awesome ;)

            n = 13; k = 7;
            d = MathX.nChooseK(n, k);
            Assert.AreEqual( 1716, d,"NChooseK ");

        }
        [TestMethod]
        public void TestSign()
        {
            short minus = -1;
            short zero = 0;
            short plus = 1;
            double d = 0;


            Assert.AreEqual( minus, MathX.sign(minus),"Sign ");
            Assert.AreEqual( plus, MathX.sign(plus),"Sign ");
            Assert.AreEqual( zero, MathX.sign(zero),"Sign ");

            d = 0;
            Assert.AreEqual( zero, MathX.sign(d),"Sign ");

            d = -1.000001;
            Assert.AreEqual( minus, MathX.sign(d),"Sign ");

            d = -.000001;
            Assert.AreEqual( minus, MathX.sign(d),"Sign ");

            d = -1E-200;
            Assert.AreEqual( minus, MathX.sign(d),"Sign ");

            d = Double.NegativeInfinity;
            Assert.AreEqual( minus, MathX.sign(d),"Sign ");

            d = -200.11;
            Assert.AreEqual( minus, MathX.sign(d),"Sign ");

            d = -2000000000000.11;
            Assert.AreEqual( minus, MathX.sign(d),"Sign ");

            d = 1.000001;
            Assert.AreEqual( plus, MathX.sign(d),"Sign ");

            d = .000001;
            Assert.AreEqual( plus, MathX.sign(d),"Sign ");

            d = 1E-200;
            Assert.AreEqual( plus, MathX.sign(d),"Sign ");

            d = Double.PositiveInfinity;
            Assert.AreEqual( plus, MathX.sign(d),"Sign ");

            d = 200.11;
            Assert.AreEqual( plus, MathX.sign(d),"Sign ");

            d = 2000000000000.11;
            Assert.AreEqual( plus, MathX.sign(d),"Sign ");

        }
        [TestMethod]
        public void TestSinh()
        {
            double d = 0;
            d = MathX.sinh(0);
            Assert.AreEqual( 0, d,"sinh 0");

            d = MathX.sinh(1);
            Assert.AreEqual( 1.175201194, d,"sinh 1 ");

            d = MathX.sinh(-1);
            Assert.AreEqual( -1.175201194, d,"sinh -1 ");

            d = MathX.sinh(-100);
            Assert.AreEqual( -1.344058570908070E+43, d,"sinh -100 ");

            d = MathX.sinh(100);
            Assert.AreEqual( 1.344058570908070E+43, d,"sinh 100 ");

            d = MathX.sinh(15);
            Assert.AreEqual( 1634508.686, d,"sinh 15");

            d = MathX.sinh(-15);
            Assert.AreEqual( -1634508.686, d,"sinh -15 ");

            d = MathX.sinh(0.1);
            Assert.AreEqual( 0.10016675, d,"sinh 0.1");

            d = MathX.sinh(-0.1);
            Assert.AreEqual( -0.10016675, d,"sinh -0.1 ");

        }
        [TestMethod]
        public void TestSum()
        {
            double[] d = new double[100];
            d[0] = 1.1; d[1] = 2.1; d[2] = 3.1; d[3] = 4.1;
            d[4] = 5.1; d[5] = 6.1; d[6] = 7.1; d[7] = 8.1;
            d[8] = 9.1; d[9] = 10.1; d[10] = 11.1; d[11] = 12.1;
            d[12] = 13.1; d[13] = 14.1; d[14] = 15.1; d[15] = 16.1;
            d[16] = 17.1; d[17] = 18.1; d[18] = 19.1; d[19] = 20.1;

            double s = MathX.sum(d);
            Assert.AreEqual( 212, s,"Sum ");

            d = new double[1000];
            s = MathX.sum(d);
            Assert.AreEqual( 0, s,"Sum ");

            d[0] = -1.1; d[1] = 2.1; d[2] = -3.1; d[3] = 4.1;
            d[4] = -5.1; d[5] = 6.1; d[6] = -7.1; d[7] = 8.1;
            d[8] = -9.1; d[9] = 10.1; d[10] = -11.1; d[11] = 12.1;
            d[12] = -13.1; d[13] = 14.1; d[14] = -15.1; d[15] = 16.1;
            d[16] = -17.1; d[17] = 18.1; d[18] = -19.1; d[19] = 20.1;
            s = MathX.sum(d);
            Assert.AreEqual( 10, s,"Sum ");

            d[0] = -1.1; d[1] = -2.1; d[2] = -3.1; d[3] = -4.1;
            d[4] = -5.1; d[5] = -6.1; d[6] = -7.1; d[7] = -8.1;
            d[8] = -9.1; d[9] = -10.1; d[10] = -11.1; d[11] = -12.1;
            d[12] = -13.1; d[13] = -14.1; d[14] = -15.1; d[15] = -16.1;
            d[16] = -17.1; d[17] = -18.1; d[18] = -19.1; d[19] = -20.1;
            s = MathX.sum(d);
            Assert.AreEqual( -212, s,"Sum ");

        }
        [TestMethod]
        public void TestSumsq()
        {
            double[] d = new double[100];
            d[0] = 1.1; d[1] = 2.1; d[2] = 3.1; d[3] = 4.1;
            d[4] = 5.1; d[5] = 6.1; d[6] = 7.1; d[7] = 8.1;
            d[8] = 9.1; d[9] = 10.1; d[10] = 11.1; d[11] = 12.1;
            d[12] = 13.1; d[13] = 14.1; d[14] = 15.1; d[15] = 16.1;
            d[16] = 17.1; d[17] = 18.1; d[18] = 19.1; d[19] = 20.1;

            double s = MathX.sumsq(d);
            Assert.AreEqual( 2912.2, s,"Sumsq ");

            d = new double[1000];
            s = MathX.sumsq(d);
            Assert.AreEqual( 0, s,"Sumsq ");

            d[0] = -1.1; d[1] = 2.1; d[2] = -3.1; d[3] = 4.1;
            d[4] = -5.1; d[5] = 6.1; d[6] = -7.1; d[7] = 8.1;
            d[8] = -9.1; d[9] = 10.1; d[10] = -11.1; d[11] = 12.1;
            d[12] = -13.1; d[13] = 14.1; d[14] = -15.1; d[15] = 16.1;
            d[16] = -17.1; d[17] = 18.1; d[18] = -19.1; d[19] = 20.1;
            s = MathX.sumsq(d);
            Assert.AreEqual( 2912.2, s,"Sumsq ");

            d[0] = -1.1; d[1] = -2.1; d[2] = -3.1; d[3] = -4.1;
            d[4] = -5.1; d[5] = -6.1; d[6] = -7.1; d[7] = -8.1;
            d[8] = -9.1; d[9] = -10.1; d[10] = -11.1; d[11] = -12.1;
            d[12] = -13.1; d[13] = -14.1; d[14] = -15.1; d[15] = -16.1;
            d[16] = -17.1; d[17] = -18.1; d[18] = -19.1; d[19] = -20.1;
            s = MathX.sumsq(d);
            Assert.AreEqual( 2912.2, s,"Sumsq ");
        }
        [TestMethod]
        public void TestFactorial()
        {
            int n = 0;
            double s = 0;

            n = 0;
            s = MathX.factorial(n);
            Assert.AreEqual( 1, s,"Factorial ");

            n = 1;
            s = MathX.factorial(n);
            Assert.AreEqual( 1, s,"Factorial ");

            n = 10;
            s = MathX.factorial(n);
            Assert.AreEqual( 3628800, s,"Factorial ");

            n = 99;
            s = MathX.factorial(n);
            Assert.AreEqual( 9.33262154439E+155, s,"Factorial ");

            n = -1;
            s = MathX.factorial(n);
            Assert.AreEqual( Double.NaN, s,"Factorial ");

            n = Int32.MaxValue;
            s = MathX.factorial(n);
            Assert.AreEqual( Double.PositiveInfinity, s,"Factorial ");
        }
        [TestMethod]
        public void TestSumx2my2()
        {
            double[] xarr = null;
            double[] yarr = null;

            xarr = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            yarr = new double[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };
            ConfirmSumx2my2(xarr, yarr, 100);

            xarr = new double[] { -1, -2, -3, -4, -5, -6, -7, -8, -9, -10 };
            yarr = new double[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };
            ConfirmSumx2my2(xarr, yarr, 100);

            xarr = new double[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };
            yarr = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            ConfirmSumx2my2(xarr, yarr, -100);

            xarr = new double[] { 10 };
            yarr = new double[] { 9 };
            ConfirmSumx2my2(xarr, yarr, 19);

            xarr = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            yarr = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            ConfirmSumx2my2(xarr, yarr, 0);
        }
        [TestMethod]
        public void TestSumx2py2()
        {
            double[] xarr = null;
            double[] yarr = null;

            xarr = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            yarr = new double[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };
            ConfirmSumx2py2(xarr, yarr, 670);

            xarr = new double[] { -1, -2, -3, -4, -5, -6, -7, -8, -9, -10 };
            yarr = new double[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };
            ConfirmSumx2py2(xarr, yarr, 670);

            xarr = new double[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };
            yarr = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            ConfirmSumx2py2(xarr, yarr, 670);

            xarr = new double[] { 10 };
            yarr = new double[] { 9 };
            ConfirmSumx2py2(xarr, yarr, 181);

            xarr = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            yarr = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            ConfirmSumx2py2(xarr, yarr, 770);
        }
        [TestMethod]
        public void TestSumxmy2()
        {
            double[] xarr = null;
            double[] yarr = null;

            xarr = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            yarr = new double[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };
            ConfirmSumxmy2(xarr, yarr, 10);

            xarr = new double[] { -1, -2, -3, -4, -5, -6, -7, -8, -9, -10 };
            yarr = new double[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };
            ConfirmSumxmy2(xarr, yarr, 1330);

            xarr = new double[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };
            yarr = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            ConfirmSumxmy2(xarr, yarr, 10);

            xarr = new double[] { 10 };
            yarr = new double[] { 9 };
            ConfirmSumxmy2(xarr, yarr, 1);

            xarr = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            yarr = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            ConfirmSumxmy2(xarr, yarr, 0);
        }

        private static void ConfirmSumx2my2(double[] xarr, double[] yarr, double expectedResult)
        {
            ConfirmXY(new Sumx2my2().CreateAccumulator(), xarr, yarr, expectedResult);
        }
        private static void ConfirmSumx2py2(double[] xarr, double[] yarr, double expectedResult)
        {
            ConfirmXY(new Sumx2py2().CreateAccumulator(), xarr, yarr, expectedResult);
        }
        private static void ConfirmSumxmy2(double[] xarr, double[] yarr, double expectedResult)
        {
            ConfirmXY(new Sumxmy2().CreateAccumulator(), xarr, yarr, expectedResult);
        }

        private static void ConfirmXY(Accumulator acc, double[] xarr, double[] yarr,
                double expectedResult)
        {
            double result = 0.0;
            for (int i = 0; i < xarr.Length; i++)
            {
                result += acc.Accumulate(xarr[i], yarr[i]);
            }
            Assert.AreEqual(expectedResult, result, 0.0);
        }
        [TestMethod]
        public void TestRound()
        {
            double d = 0;
            int p = 0;

            d = 0; p = 0;
            Assert.AreEqual( 0, MathX.round(d, p),"round ");

            d = 10; p = 0;
            Assert.AreEqual( 10, MathX.round(d, p),"round ");

            d = 123.23; p = 0;
            Assert.AreEqual( 123, MathX.round(d, p),"round ");

            d = -123.23; p = 0;
            Assert.AreEqual( -123, MathX.round(d, p),"round ");

            d = 123.12; p = 2;
            Assert.AreEqual( 123.12, MathX.round(d, p),"round ");

            d = 88.123459; p = 5;
            Assert.AreEqual( 88.12346, MathX.round(d, p),"round ");

            d = 0; p = 2;
            Assert.AreEqual( 0, MathX.round(d, p),"round ");

            d = 0; p = -1;
            Assert.AreEqual( 0, MathX.round(d, p),"round ");

            d = 0.01; p = -1;
            Assert.AreEqual( 0, MathX.round(d, p),"round ");

            d = 123.12; p = -2;
            Assert.AreEqual( 100, MathX.round(d, p),"round ");

            d = 88.123459; p = -3;
            Assert.AreEqual( 0, MathX.round(d, p),"round ");

            d = 49.00000001; p = -1;
            Assert.AreEqual( 50, MathX.round(d, p),"round ");

            d = 149.999999; p = -2;
            Assert.AreEqual( 100, MathX.round(d, p),"round ");

            d = 150.0; p = -2;
            Assert.AreEqual( 200, MathX.round(d, p),"round ");

            d = 2162.615d; p = 2;
            Assert.AreEqual( 2162.62d, MathX.round(d, p),"round ");
        }
        [TestMethod]
        public void TestRoundDown()
        {
            double d = 0;
            int p = 0;

            d = 0; p = 0;
            Assert.AreEqual( 0, MathX.roundDown(d, p),"roundDown ");

            d = 10; p = 0;
            Assert.AreEqual( 10, MathX.roundDown(d, p),"roundDown ");

            d = 123.99; p = 0;
            Assert.AreEqual( 123, MathX.roundDown(d, p),"roundDown ");

            d = -123.99; p = 0;
            Assert.AreEqual( -123, MathX.roundDown(d, p),"roundDown ");

            d = 123.99; p = 2;
            Assert.AreEqual( 123.99, MathX.roundDown(d, p),"roundDown ");

            d = 88.123459; p = 5;
            Assert.AreEqual( 88.12345, MathX.roundDown(d, p),"roundDown ");

            d = 0; p = 2;
            Assert.AreEqual( 0, MathX.roundDown(d, p),"roundDown ");

            d = 0; p = -1;
            Assert.AreEqual( 0, MathX.roundDown(d, p),"roundDown ");

            d = 0.01; p = -1;
            Assert.AreEqual( 0, MathX.roundDown(d, p),"roundDown ");

            d = 199.12; p = -2;
            Assert.AreEqual( 100, MathX.roundDown(d, p),"roundDown ");

            d = 88.123459; p = -3;
            Assert.AreEqual( 0, MathX.roundDown(d, p),"roundDown ");

            d = 99.00000001; p = -1;
            Assert.AreEqual( 90, MathX.roundDown(d, p),"roundDown ");

            d = 100.00001; p = -2;
            Assert.AreEqual( 100, MathX.roundDown(d, p),"roundDown ");

            d = 150.0; p = -2;
            Assert.AreEqual( 100, MathX.roundDown(d, p),"roundDown ");
        }
        [TestMethod]
        public void TestRoundUp()
        {
            double d = 0;
            int p = 0;

            d = 0; p = 0;
            Assert.AreEqual( 0, MathX.roundUp(d, p),"roundUp ");

            d = 10; p = 0;
            Assert.AreEqual( 10, MathX.roundUp(d, p),"roundUp ");

            d = 123.23; p = 0;
            Assert.AreEqual( 124, MathX.roundUp(d, p),"roundUp ");

            d = -123.23; p = 0;
            Assert.AreEqual( -124, MathX.roundUp(d, p),"roundUp ");

            d = 123.12; p = 2;
            Assert.AreEqual( 123.12, MathX.roundUp(d, p),"roundUp ");

            d = 88.123459; p = 5;
            Assert.AreEqual( 88.12346, MathX.roundUp(d, p),"roundUp ");

            d = 0; p = 2;
            Assert.AreEqual( 0, MathX.roundUp(d, p),"roundUp ");

            d = 0; p = -1;
            Assert.AreEqual( 0, MathX.roundUp(d, p),"roundUp ");

            d = 0.01; p = -1;
            Assert.AreEqual( 10, MathX.roundUp(d, p),"roundUp ");

            d = 123.12; p = -2;
            Assert.AreEqual( 200, MathX.roundUp(d, p),"roundUp ");

            d = 88.123459; p = -3;
            Assert.AreEqual( 1000, MathX.roundUp(d, p),"roundUp ");

            d = 49.00000001; p = -1;
            Assert.AreEqual( 50, MathX.roundUp(d, p),"roundUp ");

            d = 149.999999; p = -2;
            Assert.AreEqual( 200, MathX.roundUp(d, p),"roundUp ");

            d = 150.0; p = -2;
            Assert.AreEqual( 200, MathX.roundUp(d, p),"roundUp ");
        }
        [TestMethod]
        public void TestCeiling()
        {
            double d = 0;
            double s = 0;

            d = 0; s = 0;
            Assert.AreEqual( 0, MathX.ceiling(d, s),"ceiling ");

            d = 1; s = 0;
            Assert.AreEqual( 0, MathX.ceiling(d, s),"ceiling ");

            d = 0; s = 1;
            Assert.AreEqual( 0, MathX.ceiling(d, s),"ceiling ");

            d = -1; s = 0;
            Assert.AreEqual( 0, MathX.ceiling(d, s),"ceiling ");

            d = 0; s = -1;
            Assert.AreEqual( 0, MathX.ceiling(d, s),"ceiling ");

            d = 10; s = 1.11;
            Assert.AreEqual( 11.1, MathX.ceiling(d, s),"ceiling ");

            d = 11.12333; s = 0.03499;
            Assert.AreEqual( 11.12682, MathX.ceiling(d, s),"ceiling ");

            d = -11.12333; s = 0.03499;
            Assert.AreEqual( Double.NaN, MathX.ceiling(d, s),"ceiling ");

            d = 11.12333; s = -0.03499;
            Assert.AreEqual( Double.NaN, MathX.ceiling(d, s),"ceiling ");

            d = -11.12333; s = -0.03499;
            Assert.AreEqual( -11.12682, MathX.ceiling(d, s),"ceiling ");

            d = 100; s = 0.001;
            Assert.AreEqual( 100, MathX.ceiling(d, s),"ceiling ");

            d = -0.001; s = -9.99;
            Assert.AreEqual( -9.99, MathX.ceiling(d, s),"ceiling ");

            d = 4.42; s = 0.05;
            Assert.AreEqual( 4.45, MathX.ceiling(d, s),"ceiling ");

            d = 0.05; s = 4.42;
            Assert.AreEqual( 4.42, MathX.ceiling(d, s),"ceiling ");

            d = 0.6666; s = 3.33;
            Assert.AreEqual( 3.33, MathX.ceiling(d, s),"ceiling ");

            d = 2d / 3; s = 3.33;
            Assert.AreEqual( 3.33, MathX.ceiling(d, s),"ceiling ");
        }
        [TestMethod]
        public void TestFloor()
        {
            double d = 0;
            double s = 0;

            d = 0; s = 0;
            Assert.AreEqual( 0, MathX.floor(d, s),"floor ");

            d = 1; s = 0;
            Assert.AreEqual( Double.NaN, MathX.floor(d, s),"floor ");

            d = 0; s = 1;
            Assert.AreEqual( 0, MathX.floor(d, s),"floor ");

            d = -1; s = 0;
            Assert.AreEqual( Double.NaN, MathX.floor(d, s),"floor ");

            d = 0; s = -1;
            Assert.AreEqual( 0, MathX.floor(d, s),"floor ");

            d = 10; s = 1.11;
            Assert.AreEqual( 9.99, MathX.floor(d, s),"floor ");

            d = 11.12333; s = 0.03499;
            Assert.AreEqual( 11.09183, MathX.floor(d, s),"floor ");

            d = -11.12333; s = 0.03499;
            Assert.AreEqual( Double.NaN, MathX.floor(d, s),"floor ");

            d = 11.12333; s = -0.03499;
            Assert.AreEqual( Double.NaN, MathX.floor(d, s),"floor ");

            d = -11.12333; s = -0.03499;
            Assert.AreEqual( -11.09183, MathX.floor(d, s),"floor ");

            d = 100; s = 0.001;
            Assert.AreEqual( 100, MathX.floor(d, s),"floor ");

            d = -0.001; s = -9.99;
            Assert.AreEqual( 0, MathX.floor(d, s),"floor ");

            d = 4.42; s = 0.05;
            Assert.AreEqual( 4.4, MathX.floor(d, s),"floor ");

            d = 0.05; s = 4.42;
            Assert.AreEqual( 0, MathX.floor(d, s),"floor ");

            d = 0.6666; s = 3.33;
            Assert.AreEqual( 0, MathX.floor(d, s),"floor ");

            d = 2d / 3; s = 3.33;
            Assert.AreEqual( 0, MathX.floor(d, s),"floor ");
        }

    }

}