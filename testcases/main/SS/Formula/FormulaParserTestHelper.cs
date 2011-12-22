/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace TestCases.SS.Formula
{

    using System;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    /**
     * Avoids making {@link FormulaParser#FormulaParseException} public
     * 
     * @author Josh Micich
     */
    public class FormulaParserTestHelper
    {
        public static void ConfirmParseException(Exception e, String expectedMessage)
        {
            CheckType(e);
            Assert.AreEqual(expectedMessage, e.Message);
        }
        public static void ConfirmParseException(Exception e)
        {
            CheckType(e);
            Assert.IsNotNull(e.Message);
        }
        private static void CheckType(Exception e)
        {
            if (!(e is FormulaParseException))
            {
                String failMsg = "Expected FormulaParseException, but got ("
                    + e.GetType().Name + "):";
                Console.WriteLine(failMsg);
                throw new AssertFailedException(failMsg);
            }
        }
    }
}
