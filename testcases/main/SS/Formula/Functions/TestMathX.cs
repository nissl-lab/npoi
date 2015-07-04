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
    using NUnit.Framework;
    using NPOI.SS.Formula.Functions;


    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *  
     */
    [TestFixture]
    public class TestMathX : AbstractNumericTestCase
    {
        [Test]
        public void TestAcosh()
        {
            double d = 0;

            d = MathX.Acosh(0);
            Assert.IsTrue(Double.IsNaN(d), "Acosh 0 is NaN");

            d = MathX.Acosh(1);
            AssertEqual("Acosh 1 ",0, d);

            d = MathX.Acosh(-1);
            Assert.IsTrue(Double.IsNaN(d), "Acosh -1 is NaN");

            d = MathX.Acosh(100);
            AssertEqual("Acosh 100 ", 5.298292366d, d);

            d = MathX.Acosh(101.001);
            AssertEqual("Acosh 101.001 ",5.308253091d, d);

            d = MathX.Acosh(200000);
            AssertEqual("Acosh 200000 ",12.89921983d, d);

        }
        [Test]
        public void TestAsinh()
        {
            double d = 0;

            d = MathX.Asinh(0);
            AssertEqual("asinh 0",d, 0);

            d = MathX.Asinh(1);
            AssertEqual("asinh 1 ",0.881373587, d);

            d = MathX.Asinh(-1);
            AssertEqual("asinh -1 ",-0.881373587, d);

            d = MathX.Asinh(-100);
            AssertEqual("asinh -100 ",-5.298342366, d);

            d = MathX.Asinh(100);
            AssertEqual("asinh 100 ",5.298342366, d);

            d = MathX.Asinh(200000);
            AssertEqual("asinh 200000",12.899219826096400, d);

            d = MathX.Asinh(-200000);
            AssertEqual("asinh -200000 ",-12.899223853137, d);

        }
        [Test]
        public void TestAtanh()
        {
            double d = 0;
            d = MathX.Atanh(0);
            AssertEqual("atanh 0", d, 0);

            d = MathX.Atanh(1);
            AssertEqual("atanh 1 ", Double.PositiveInfinity, d);

            d = MathX.Atanh(-1);
            AssertEqual("atanh -1 ", Double.NegativeInfinity, d);

            d = MathX.Atanh(-100);
            AssertEqual("atanh -100 ", Double.NaN, d);
            
            d = MathX.Atanh(100);
            AssertEqual("atanh 100 ", Double.NaN, d);

            d = MathX.Atanh(200000);
            AssertEqual("atanh 200000", Double.NaN, d);

            d = MathX.Atanh(-200000);
            AssertEqual("atanh -200000 ", Double.NaN, d);

            d = MathX.Atanh(0.1);
            AssertEqual("atanh 0.1", 0.100335348, d);

            d = MathX.Atanh(-0.1);
            AssertEqual("atanh -0.1 ", -0.100335348, d);

        }
        [Test]
        public void TestCosh()
        {
            double d = 0;
            d = MathX.Cosh(0);
            AssertEqual("cosh 0", 1, d);

            d = MathX.Cosh(1);
            AssertEqual("cosh 1 ", 1.543080635, d);

            d = MathX.Cosh(-1);
            AssertEqual("cosh -1 ", 1.543080635, d);

            d = MathX.Cosh(-100);
            AssertEqual("cosh -100 ", 1.344058570908070E+43, d);

            d = MathX.Cosh(100);
            AssertEqual("cosh 100 ", 1.344058570908070E+43, d);

            d = MathX.Cosh(15);
            AssertEqual("cosh 15", 1634508.686, d);

            d = MathX.Cosh(-15);
            AssertEqual("cosh -15 ", 1634508.686, d);

            d = MathX.Cosh(0.1);
            AssertEqual("cosh 0.1", 1.005004168, d);

            d = MathX.Cosh(-0.1);
            AssertEqual("cosh -0.1 ", 1.005004168, d);

        }
        [Test]
        public void TestTanh()
        {
            double d = 0;
            d = MathX.Tanh(0);
            AssertEqual("tanh 0", 0, d);

            d = MathX.Tanh(1);
            AssertEqual("tanh 1 ", 0.761594156, d);

            d = MathX.Tanh(-1);
            AssertEqual("tanh -1 ", -0.761594156, d);

            d = MathX.Tanh(-100);
            AssertEqual("tanh -100 ", -1, d);

            d = MathX.Tanh(100);
            AssertEqual("tanh 100 ", 1, d);

            d = MathX.Tanh(15);
            AssertEqual("tanh 15", 1, d);

            d = MathX.Tanh(-15);
            AssertEqual("tanh -15 ", -1, d);

            d = MathX.Tanh(0.1);
            AssertEqual("tanh 0.1", 0.099667995, d);

            d = MathX.Tanh(-0.1);
            AssertEqual("tanh -0.1 ", -0.099667995, d);

        }
        [Test]
        public void TestMax()
        {
            double[] d = new double[100];
            d[0] = 1.1; d[1] = 2.1; d[2] = 3.1; d[3] = 4.1;
            d[4] = 5.1; d[5] = 6.1; d[6] = 7.1; d[7] = 8.1;
            d[8] = 9.1; d[9] = 10.1; d[10] = 11.1; d[11] = 12.1;
            d[12] = 13.1; d[13] = 14.1; d[14] = 15.1; d[15] = 16.1;
            d[16] = 17.1; d[17] = 18.1; d[18] = 19.1; d[19] = 20.1;

            double m = MathX.Max(d);
            AssertEqual("Max ", 20.1, m);

            d = new double[1000];
            m = MathX.Max(d);
            AssertEqual("Max ", 0, m);

            d[0] = -1.1; d[1] = 2.1; d[2] = -3.1; d[3] = 4.1;
            d[4] = -5.1; d[5] = 6.1; d[6] = -7.1; d[7] = 8.1;
            d[8] = -9.1; d[9] = 10.1; d[10] = -11.1; d[11] = 12.1;
            d[12] = -13.1; d[13] = 14.1; d[14] = -15.1; d[15] = 16.1;
            d[16] = -17.1; d[17] = 18.1; d[18] = -19.1; d[19] = 20.1;
            m = MathX.Max(d);
            AssertEqual("Max ", 20.1, m);

            d = new double[20];
            d[0] = -1.1; d[1] = -2.1; d[2] = -3.1; d[3] = -4.1;
            d[4] = -5.1; d[5] = -6.1; d[6] = -7.1; d[7] = -8.1;
            d[8] = -9.1; d[9] = -10.1; d[10] = -11.1; d[11] = -12.1;
            d[12] = -13.1; d[13] = -14.1; d[14] = -15.1; d[15] = -16.1;
            d[16] = -17.1; d[17] = -18.1; d[18] = -19.1; d[19] = -20.1;
            m = MathX.Max(d);
            AssertEqual("Max ", -1.1, m);

        }
        [Test]
        public void TestMin()
        {
            double[] d = new double[100];
            d[0] = 1.1; d[1] = 2.1; d[2] = 3.1; d[3] = 4.1;
            d[4] = 5.1; d[5] = 6.1; d[6] = 7.1; d[7] = 8.1;
            d[8] = 9.1; d[9] = 10.1; d[10] = 11.1; d[11] = 12.1;
            d[12] = 13.1; d[13] = 14.1; d[14] = 15.1; d[15] = 16.1;
            d[16] = 17.1; d[17] = 18.1; d[18] = 19.1; d[19] = 20.1;

            double m = MathX.Min(d);
            AssertEqual("Min ", 0, m);

            d = new double[20];
            d[0] = 1.1; d[1] = 2.1; d[2] = 3.1; d[3] = 4.1;
            d[4] = 5.1; d[5] = 6.1; d[6] = 7.1; d[7] = 8.1;
            d[8] = 9.1; d[9] = 10.1; d[10] = 11.1; d[11] = 12.1;
            d[12] = 13.1; d[13] = 14.1; d[14] = 15.1; d[15] = 16.1;
            d[16] = 17.1; d[17] = 18.1; d[18] = 19.1; d[19] = 20.1;

            m = MathX.Min(d);
            AssertEqual("Min ", 1.1, m);

            d = new double[1000];
            m = MathX.Min(d);
            AssertEqual("Min ", 0, m);

            d[0] = -1.1; d[1] = 2.1; d[2] = -3.1; d[3] = 4.1;
            d[4] = -5.1; d[5] = 6.1; d[6] = -7.1; d[7] = 8.1;
            d[8] = -9.1; d[9] = 10.1; d[10] = -11.1; d[11] = 12.1;
            d[12] = -13.1; d[13] = 14.1; d[14] = -15.1; d[15] = 16.1;
            d[16] = -17.1; d[17] = 18.1; d[18] = -19.1; d[19] = 20.1;
            m = MathX.Min(d);
            AssertEqual("Min ", -19.1, m);

            d = new double[20];
            d[0] = -1.1; d[1] = -2.1; d[2] = -3.1; d[3] = -4.1;
            d[4] = -5.1; d[5] = -6.1; d[6] = -7.1; d[7] = -8.1;
            d[8] = -9.1; d[9] = -10.1; d[10] = -11.1; d[11] = -12.1;
            d[12] = -13.1; d[13] = -14.1; d[14] = -15.1; d[15] = -16.1;
            d[16] = -17.1; d[17] = -18.1; d[18] = -19.1; d[19] = -20.1;
            m = MathX.Min(d);
            AssertEqual("Min ", -20.1, m);
        }
        [Test]
        public void TestProduct()
        {
            Assert.AreEqual(0, MathX.Product(null), "Product ");
            Assert.AreEqual(0, MathX.Product(new double[] { }), "Product ");
            Assert.AreEqual(0, MathX.Product(new double[] { 1, 0 }), "Product ");

            Assert.AreEqual(1, MathX.Product(new double[] { 1 }), "Product ");
            Assert.AreEqual(1, MathX.Product(new double[] { 1, 1 }), "Product ");
            Assert.AreEqual(10, MathX.Product(new double[] { 10, 1 }), "Product ");
            Assert.AreEqual(-2, MathX.Product(new double[] { 2, -1 }), "Product ");
            Assert.AreEqual(99988000209999d, MathX.Product(new double[] { 99999, 99999, 9999 }), "Product ");
        

            double[] d = new double[100];
            d[0] = 1.1; d[1] = 2.1; d[2] = 3.1; d[3] = 4.1;
            d[4] = 5.1; d[5] = 6.1; d[6] = 7.1; d[7] = 8.1;
            d[8] = 9.1; d[9] = 10.1; d[10] = 11.1; d[11] = 12.1;
            d[12] = 13.1; d[13] = 14.1; d[14] = 15.1; d[15] = 16.1;
            d[16] = 17.1; d[17] = 18.1; d[18] = 19.1; d[19] = 20.1;

            double m = MathX.Product(d);
            AssertEqual("Product", 0, m);

            d = new double[20];
            d[0] = 1.1; d[1] = 2.1; d[2] = 3.1; d[3] = 4.1;
            d[4] = 5.1; d[5] = 6.1; d[6] = 7.1; d[7] = 8.1;
            d[8] = 9.1; d[9] = 10.1; d[10] = 11.1; d[11] = 12.1;
            d[12] = 13.1; d[13] = 14.1; d[14] = 15.1; d[15] = 16.1;
            d[16] = 17.1; d[17] = 18.1; d[18] = 19.1; d[19] = 20.1;

            m = MathX.Product(d);
            AssertEqual("Product ", 3459946360003355534d, m);

            d = new double[1000];
            m = MathX.Product(d);
            AssertEqual("Product ", 0, m);

            d = new double[20];
            d[0] = -1.1; d[1] = -2.1; d[2] = -3.1; d[3] = -4.1;
            d[4] = -5.1; d[5] = -6.1; d[6] = -7.1; d[7] = -8.1;
            d[8] = -9.1; d[9] = -10.1; d[10] = -11.1; d[11] = -12.1;
            d[12] = -13.1; d[13] = -14.1; d[14] = -15.1; d[15] = -16.1;
            d[16] = -17.1; d[17] = -18.1; d[18] = -19.1; d[19] = -20.1;
            m = MathX.Product(d);
            AssertEqual("Product ", 3459946360003355534d, m);
        }
        [Test]
        public void TestMod()
        {

            //example from Excel help
            Assert.AreEqual(1.0, MathX.Mod(3, 2));
            Assert.AreEqual(1.0, MathX.Mod(-3, 2));
            Assert.AreEqual(-1.0, MathX.Mod(3, -2));
            Assert.AreEqual(-1.0, MathX.Mod(-3, -2));

            Assert.AreEqual(0.0, MathX.Mod(0, 2));
            Assert.AreEqual(Double.NaN, MathX.Mod(3, 0));

            Assert.AreEqual((double)1.4, MathX.Mod(3.4, 2));
            Assert.AreEqual((double)-1.4, MathX.Mod(-3.4, -2));
            Assert.AreEqual((double)0.6000000000000001, MathX.Mod(-3.4, 2.0));// should actually be 0.6
            Assert.AreEqual((double)-0.6000000000000001, MathX.Mod(3.4, -2.0));// should actually be -0.6

            Assert.AreEqual(3.0, MathX.Mod(3, Double.MaxValue));
            Assert.AreEqual(2.0, MathX.Mod(Double.MaxValue, 3));

            // Bugzilla 50033
            Assert.AreEqual(1.0, MathX.Mod(13, 12));
        }
        [Test]
        public void TestNChooseK()
        {
            int n = 100;
            int k = 50;
            double d = MathX.NChooseK(n, k);
            AssertEqual("NChooseK ", 1.00891344545564E29, d);

            n = -1; k = 1;
            d = MathX.NChooseK(n, k);
            AssertEqual("NChooseK ", Double.NaN, d);

            n = 1; k = -1;
            d = MathX.NChooseK(n, k);
            AssertEqual("NChooseK ", Double.NaN, d);

            n = 0; k = 1;
            d = MathX.NChooseK(n, k);
            AssertEqual("NChooseK ", Double.NaN, d);

            n = 1; k = 0;
            d = MathX.NChooseK(n, k);
            AssertEqual("NChooseK ", 1, d);

            n = 10; k = 9;
            d = MathX.NChooseK(n, k);
            AssertEqual("NChooseK ", 10, d);

            n = 10; k = 10;
            d = MathX.NChooseK(n, k);
            AssertEqual("NChooseK ", 1, d);

            n = 10; k = 1;
            d = MathX.NChooseK(n, k);
            AssertEqual("NChooseK ", 10, d);

            n = 1000; k = 1;
            d = MathX.NChooseK(n, k);
            AssertEqual("NChooseK ", 1000, d); // awesome ;)

            n = 1000; k = 2;
            d = MathX.NChooseK(n, k);
            AssertEqual("NChooseK ", 499500, d); // awesome ;)

            n = 13; k = 7;
            d = MathX.NChooseK(n, k);
            AssertEqual("NChooseK ", 1716, d);

        }
        [Test]
        public void TestSign()
        {
            short minus = -1;
            short zero = 0;
            short plus = 1;
            double d = 0;


            AssertEqual("Sign ", minus, MathX.Sign(minus));
            AssertEqual("Sign ", plus, MathX.Sign(plus));
            AssertEqual("Sign ", zero, MathX.Sign(zero));

            d = 0;
            AssertEqual("Sign ", zero, MathX.Sign(d));

            d = -1.000001;
            AssertEqual("Sign ", minus, MathX.Sign(d));

            d = -.000001;
            AssertEqual("Sign ", minus, MathX.Sign(d));

            d = -1E-200;
            AssertEqual("Sign ", minus, MathX.Sign(d));

            d = Double.NegativeInfinity;
            AssertEqual("Sign ", minus, MathX.Sign(d));

            d = -200.11;
            AssertEqual("Sign ", minus, MathX.Sign(d));

            d = -2000000000000.11;
            AssertEqual("Sign ", minus, MathX.Sign(d));

            d = 1.000001;
            AssertEqual("Sign ", plus, MathX.Sign(d));

            d = .000001;
            AssertEqual("Sign ", plus, MathX.Sign(d));

            d = 1E-200;
            AssertEqual("Sign ", plus, MathX.Sign(d));

            d = Double.PositiveInfinity;
            AssertEqual("Sign ", plus, MathX.Sign(d));

            d = 200.11;
            AssertEqual("Sign ", plus, MathX.Sign(d));

            d = 2000000000000.11;
            AssertEqual("Sign ", plus, MathX.Sign(d));

        }
        [Test]
        public void TestSinh()
        {
            double d = 0;
            d = MathX.Sinh(0);
            AssertEqual("sinh 0", 0, d);

            d = MathX.Sinh(1);
            AssertEqual("sinh 1 ", 1.175201194, d);

            d = MathX.Sinh(-1);
            AssertEqual("sinh -1 ", -1.175201194, d);

            d = MathX.Sinh(-100);
            AssertEqual("sinh -100 ", -1.344058570908070E+43, d);

            d = MathX.Sinh(100);
            AssertEqual("sinh 100 ", 1.344058570908070E+43, d);

            d = MathX.Sinh(15);
            AssertEqual("sinh 15", 1634508.686, d);

            d = MathX.Sinh(-15);
            AssertEqual("sinh -15 ", -1634508.686, d);

            d = MathX.Sinh(0.1);
            AssertEqual("sinh 0.1", 0.10016675, d);

            d = MathX.Sinh(-0.1);
            AssertEqual("sinh -0.1 ", -0.10016675, d);

        }
        [Test]
        public void TestSum()
        {
            double[] d = new double[100];
            d[0] = 1.1; d[1] = 2.1; d[2] = 3.1; d[3] = 4.1;
            d[4] = 5.1; d[5] = 6.1; d[6] = 7.1; d[7] = 8.1;
            d[8] = 9.1; d[9] = 10.1; d[10] = 11.1; d[11] = 12.1;
            d[12] = 13.1; d[13] = 14.1; d[14] = 15.1; d[15] = 16.1;
            d[16] = 17.1; d[17] = 18.1; d[18] = 19.1; d[19] = 20.1;

            double s = MathX.Sum(d);
            AssertEqual( "Sum ", 212, s );

            d = new double[1000];
            s = MathX.Sum(d);
            AssertEqual("Sum ", 0d, s);

            d[0] = -1.1; d[1] = 2.1; d[2] = -3.1; d[3] = 4.1;
            d[4] = -5.1; d[5] = 6.1; d[6] = -7.1; d[7] = 8.1;
            d[8] = -9.1; d[9] = 10.1; d[10] = -11.1; d[11] = 12.1;
            d[12] = -13.1; d[13] = 14.1; d[14] = -15.1; d[15] = 16.1;
            d[16] = -17.1; d[17] = 18.1; d[18] = -19.1; d[19] = 20.1;
            s = MathX.Sum(d);
            AssertEqual("Sum ", 10d, s);

            d[0] = -1.1; d[1] = -2.1; d[2] = -3.1; d[3] = -4.1;
            d[4] = -5.1; d[5] = -6.1; d[6] = -7.1; d[7] = -8.1;
            d[8] = -9.1; d[9] = -10.1; d[10] = -11.1; d[11] = -12.1;
            d[12] = -13.1; d[13] = -14.1; d[14] = -15.1; d[15] = -16.1;
            d[16] = -17.1; d[17] = -18.1; d[18] = -19.1; d[19] = -20.1;
            s = MathX.Sum(d);
            AssertEqual("Sum ", -212d, s);

        }
        [Test]
        public void TestSumsq()
        {
            double[] d = new double[100];
            d[0] = 1.1; d[1] = 2.1; d[2] = 3.1; d[3] = 4.1;
            d[4] = 5.1; d[5] = 6.1; d[6] = 7.1; d[7] = 8.1;
            d[8] = 9.1; d[9] = 10.1; d[10] = 11.1; d[11] = 12.1;
            d[12] = 13.1; d[13] = 14.1; d[14] = 15.1; d[15] = 16.1;
            d[16] = 17.1; d[17] = 18.1; d[18] = 19.1; d[19] = 20.1;

            double s = MathX.Sumsq(d);
            AssertEqual("Sumsq ", 2912.2, s);

            d = new double[1000];
            s = MathX.Sumsq(d);
            AssertEqual("Sumsq ", 0, s);

            d[0] = -1.1; d[1] = 2.1; d[2] = -3.1; d[3] = 4.1;
            d[4] = -5.1; d[5] = 6.1; d[6] = -7.1; d[7] = 8.1;
            d[8] = -9.1; d[9] = 10.1; d[10] = -11.1; d[11] = 12.1;
            d[12] = -13.1; d[13] = 14.1; d[14] = -15.1; d[15] = 16.1;
            d[16] = -17.1; d[17] = 18.1; d[18] = -19.1; d[19] = 20.1;
            s = MathX.Sumsq(d);
            AssertEqual("Sumsq ", 2912.2, s);

            d[0] = -1.1; d[1] = -2.1; d[2] = -3.1; d[3] = -4.1;
            d[4] = -5.1; d[5] = -6.1; d[6] = -7.1; d[7] = -8.1;
            d[8] = -9.1; d[9] = -10.1; d[10] = -11.1; d[11] = -12.1;
            d[12] = -13.1; d[13] = -14.1; d[14] = -15.1; d[15] = -16.1;
            d[16] = -17.1; d[17] = -18.1; d[18] = -19.1; d[19] = -20.1;
            s = MathX.Sumsq(d);
            AssertEqual("Sumsq ", 2912.2, s);
        }
        [Test]
        public void TestFactorial()
        {
            int n = 0;
            double s = 0;

            n = 0;
            s = MathX.Factorial(n);
            AssertEqual("Factorial ", 1, s);

            n = 1;
            s = MathX.Factorial(n);
            AssertEqual("Factorial ", 1, s);

            n = 10;
            s = MathX.Factorial(n);
            AssertEqual("Factorial ", 3628800, s);

            n = 99;
            s = MathX.Factorial(n);
            AssertEqual("Factorial ", 9.33262154439E+155, s);

            n = -1;
            s = MathX.Factorial(n);
            AssertEqual("Factorial ", Double.NaN, s);

            n = Int32.MaxValue;
            s = MathX.Factorial(n);
            AssertEqual("Factorial ", Double.PositiveInfinity, s);
        }
        [Test]
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
        [Test]
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
        [Test]
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
        [Test]
        public void TestRound()
        {
            double d = 0;
            int p = 0;

            d = 0; p = 0;
            AssertEqual("round ", 0, MathX.Round(d, p));

            d = 10; p = 0;
            AssertEqual("round ", 10, MathX.Round(d, p));

            d = 123.23; p = 0;
            AssertEqual("round ", 123, MathX.Round(d, p));

            d = -123.23; p = 0;
            AssertEqual("round ", -123, MathX.Round(d, p));

            d = 123.12; p = 2;
            AssertEqual("round ", 123.12, MathX.Round(d, p));

            d = 88.123459; p = 5;
            AssertEqual("round ", 88.12346, MathX.Round(d, p));

            d = 0; p = 2;
            AssertEqual("round ", 0, MathX.Round(d, p));

            d = 0; p = -1;
            AssertEqual("round ", 0, MathX.Round(d, p));

            d = 0.01; p = -1;
            AssertEqual("round ", 0, MathX.Round(d, p));

            d = 123.12; p = -2;
            AssertEqual("round ", 100, MathX.Round(d, p));

            d = 88.123459; p = -3;
            AssertEqual("round ", 0, MathX.Round(d, p));

            d = 49.00000001; p = -1;
            AssertEqual("round ", 50, MathX.Round(d, p));

            d = 149.999999; p = -2;
            AssertEqual("round ", 100, MathX.Round(d, p));

            d = 150.0; p = -2;
            AssertEqual("round ", 200, MathX.Round(d, p));

            d = 2162.615d; p = 2;
            AssertEqual("round ", 2162.62d, MathX.Round(d, p));


            d = 0.049999999999999975d; p = 2;
            AssertEqual("round ", 0.05d, MathX.Round(d, p));

            d = 0.049999999999999975d; p = 1;
            AssertEqual("round ", 0.1d, MathX.Round(d, p));

            d = Double.NaN; p = 1;
            AssertEqual("round ", Double.NaN, MathX.Round(d, p));

            d = Double.PositiveInfinity; p = 1;
            AssertEqual("round ", Double.NaN, MathX.Round(d, p));

            d = Double.NegativeInfinity; p = 1;
            AssertEqual("round ", Double.NaN, MathX.Round(d, p));

            d = Double.MaxValue; p = 1;
            AssertEqual("round ", Double.MaxValue, MathX.Round(d, p));

            d = Double.MinValue; p = 1;
            AssertEqual("round ", 0.0d, MathX.Round(d, p));
        }
        [Test]
        public void TestRoundDown()
        {
            double d = 0;
            int p = 0;

            d = 0; p = 0;
            AssertEqual("roundDown ", 0, MathX.RoundDown(d, p));

            d = 10; p = 0;
            AssertEqual("roundDown ", 10, MathX.RoundDown(d, p));

            d = 123.99; p = 0;
            AssertEqual("roundDown ", 123, MathX.RoundDown(d, p));

            d = -123.99; p = 0;
            AssertEqual("roundDown ", -123, MathX.RoundDown(d, p));

            d = 123.99; p = 2;
            AssertEqual("roundDown ", 123.99, MathX.RoundDown(d, p));

            d = 88.123459; p = 5;
            AssertEqual("roundDown ", 88.12345, MathX.RoundDown(d, p));

            d = 0; p = 2;
            AssertEqual("roundDown ", 0, MathX.RoundDown(d, p));

            d = 0; p = -1;
            AssertEqual("roundDown ", 0, MathX.RoundDown(d, p));

            d = 0.01; p = -1;
            AssertEqual("roundDown ", 0, MathX.RoundDown(d, p));

            d = 199.12; p = -2;
            AssertEqual("roundDown ", 100, MathX.RoundDown(d, p));

            d = 88.123459; p = -3;
            AssertEqual("roundDown ", 0, MathX.RoundDown(d, p));

            d = 99.00000001; p = -1;
            AssertEqual("roundDown ", 90, MathX.RoundDown(d, p));

            d = 100.00001; p = -2;
            AssertEqual("roundDown ", 100, MathX.RoundDown(d, p));

            d = 150.0; p = -2;
            AssertEqual("roundDown ", 100, MathX.RoundDown(d, p));

            d = 0.049999999999999975d; p = 2;
            AssertEqual("roundDown ", 0.04d, MathX.RoundDown(d, p));

            d = 0.049999999999999975d; p = 1;
            AssertEqual("roundDown ", 0.0d, MathX.RoundDown(d, p));

            d = Double.NaN; p = 1;
            AssertEqual("roundDown ", Double.NaN, MathX.RoundDown(d, p));

            d = Double.PositiveInfinity; p = 1;
            AssertEqual("roundDown ", Double.NaN, MathX.RoundDown(d, p));

            d = Double.NegativeInfinity; p = 1;
            AssertEqual("roundDown ", Double.NaN, MathX.RoundDown(d, p));

            d = Double.MaxValue; p = 1;
            AssertEqual("roundDown ", Double.MaxValue, MathX.RoundDown(d, p));

            d = Double.MinValue; p = 1;
            AssertEqual("roundDown ", 0.0d, MathX.RoundDown(d, p));
        }
        [Test]
        public void TestRoundUp()
        {
            double d = 0;
            int p = 0;

            d = 0; p = 0;
            AssertEqual("roundUp ", 0, MathX.RoundUp(d, p));

            d = 10; p = 0;
            AssertEqual("roundUp ", 10, MathX.RoundUp(d, p));

            d = 123.23; p = 0;
            AssertEqual("roundUp ", 124, MathX.RoundUp(d, p));

            d = -123.23; p = 0;
            AssertEqual("roundUp ", -124, MathX.RoundUp(d, p));

            d = 123.12; p = 2;
            AssertEqual("roundUp ", 123.12, MathX.RoundUp(d, p));

            d = 88.123459; p = 5;
            AssertEqual("roundUp ", 88.12346, MathX.RoundUp(d, p));

            d = 0; p = 2;
            AssertEqual("roundUp ", 0, MathX.RoundUp(d, p));

            d = 0; p = -1;
            AssertEqual("roundUp ", 0, MathX.RoundUp(d, p));

            d = 0.01; p = -1;
            AssertEqual("roundUp ", 10, MathX.RoundUp(d, p));

            d = 123.12; p = -2;
            AssertEqual("roundUp ", 200, MathX.RoundUp(d, p));

            d = 88.123459; p = -3;
            AssertEqual("roundUp ", 1000, MathX.RoundUp(d, p));

            d = 49.00000001; p = -1;
            AssertEqual("roundUp ", 50, MathX.RoundUp(d, p));

            d = 149.999999; p = -2;
            AssertEqual("roundUp ", 200, MathX.RoundUp(d, p));

            d = 150.0; p = -2;
            AssertEqual("roundUp ", 200, MathX.RoundUp(d, p));

            d = 0.049999999999999975d; p = 2;
            AssertEqual("round ", 0.05d, MathX.RoundUp(d, p));

            d = 0.049999999999999975d; p = 1;
            AssertEqual("round ", 0.1d, MathX.RoundUp(d, p));

            d = Double.NaN; p = 1;
            AssertEqual("round ", Double.NaN, MathX.RoundUp(d, p));

            d = Double.PositiveInfinity; p = 1;
            AssertEqual("round ", Double.NaN, MathX.RoundUp(d, p));

            d = Double.NegativeInfinity; p = 1;
            AssertEqual("round ", Double.NaN, MathX.RoundUp(d, p));

            d = Double.MaxValue; p = 1;
            AssertEqual("round ", Double.MaxValue, MathX.RoundUp(d, p));

            d = Double.MinValue; p = 1;
            AssertEqual("round ", 0.1d, MathX.RoundUp(d, p));
        }
        [Test]
        public void TestCeiling()
        {
            double d = 0;
            double s = 0;

            d = 0; s = 0;
            AssertEqual("ceiling ", 0, MathX.Ceiling(d, s));

            d = 1; s = 0;
            AssertEqual("ceiling ", 0, MathX.Ceiling(d, s));

            d = 0; s = 1;
            AssertEqual("ceiling ", 0, MathX.Ceiling(d, s));

            d = -1; s = 0;
            AssertEqual("ceiling ", 0, MathX.Ceiling(d, s));

            d = 0; s = -1;
            AssertEqual("ceiling ", 0, MathX.Ceiling(d, s));

            d = 10; s = 1.11;
            AssertEqual("ceiling ", 11.1, MathX.Ceiling(d, s));

            d = 11.12333; s = 0.03499;
            AssertEqual("ceiling ", 11.12682, MathX.Ceiling(d, s));

            d = -11.12333; s = 0.03499;
            AssertEqual("ceiling ", Double.NaN, MathX.Ceiling(d, s));

            d = 11.12333; s = -0.03499;
            AssertEqual("ceiling ", Double.NaN, MathX.Ceiling(d, s));

            d = -11.12333; s = -0.03499;
            AssertEqual("ceiling ", -11.12682, MathX.Ceiling(d, s));

            d = 100; s = 0.001;
            AssertEqual("ceiling ", 100, MathX.Ceiling(d, s));

            d = -0.001; s = -9.99;
            AssertEqual("ceiling ", -9.99, MathX.Ceiling(d, s));

            d = 4.42; s = 0.05;
            AssertEqual("ceiling ", 4.45, MathX.Ceiling(d, s));

            d = 0.05; s = 4.42;
            AssertEqual("ceiling ", 4.42, MathX.Ceiling(d, s));

            d = 0.6666; s = 3.33;
            AssertEqual("ceiling ", 3.33, MathX.Ceiling(d, s));

            d = 2d / 3; s = 3.33;
            AssertEqual("ceiling ", 3.33, MathX.Ceiling(d, s));
        }
        [Test]
        public void TestFloor()
        {
            double d = 0;
            double s = 0;

            d = 0; s = 0;
            AssertEqual("floor ", 0, MathX.Floor(d, s));

            d = 1; s = 0;
            AssertEqual("floor ", Double.NaN, MathX.Floor(d, s));

            d = 0; s = 1;
            AssertEqual("floor ", 0, MathX.Floor(d, s));

            d = -1; s = 0;
            AssertEqual("floor ", Double.NaN, MathX.Floor(d, s));

            d = 0; s = -1;
            AssertEqual("floor ", 0, MathX.Floor(d, s));

            d = 10; s = 1.11;
            AssertEqual("floor ", 9.99, MathX.Floor(d, s));

            d = 11.12333; s = 0.03499;
            AssertEqual("floor ", 11.09183, MathX.Floor(d, s));

            d = -11.12333; s = 0.03499;
            AssertEqual("floor ", Double.NaN, MathX.Floor(d, s));

            d = 11.12333; s = -0.03499;
            AssertEqual("floor ", Double.NaN, MathX.Floor(d, s));

            d = -11.12333; s = -0.03499;
            AssertEqual("floor ", -11.09183, MathX.Floor(d, s));

            d = 100; s = 0.001;
            AssertEqual("floor ", 100, MathX.Floor(d, s));

            d = -0.001; s = -9.99;
            AssertEqual("floor ", 0, MathX.Floor(d, s));

            d = 4.42; s = 0.05;
            AssertEqual("floor ", 4.4, MathX.Floor(d, s));

            d = 0.05; s = 4.42;
            AssertEqual("floor ", 0, MathX.Floor(d, s));

            d = 0.6666; s = 3.33;
            AssertEqual("floor ", 0, MathX.Floor(d, s));

            d = 2d / 3; s = 3.33;
            AssertEqual("floor ", 0, MathX.Floor(d, s));
        }
        [Ignore("not implement")]
        [Test]
        public void TestCoverage()
        {
            //// get the default constructor
            //final Constructor<MathX> c = MathX.class.getDeclaredConstructor(new Class[] {});

            //// make it callable from the outside
            //c.setAccessible(true);

            //// call it
            //c.newInstance((Object[]) null);
        }
    }

}