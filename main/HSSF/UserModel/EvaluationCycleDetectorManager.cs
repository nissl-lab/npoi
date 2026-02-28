/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.HSSF.UserModel
{
    using System;

    /**
     * This class makes an <c>EvaluationCycleDetector</c> instance available to
     * each thRead via a <c>ThReadLocal</c> in order to avoid Adding a parameter
     * to a few protected methods within <c>HSSFFormulaEvaluator</c>.
     * 
     * @author Josh Micich
     */
    class EvaluationCycleDetectorManager
    {
        [ThreadStatic]
        static EvaluationCycleDetector ecd = new EvaluationCycleDetector();

        /**
         * @return
         */
        public static EvaluationCycleDetector GetTracker()
        {
            return ecd;
        }

        private EvaluationCycleDetectorManager()
        {
            // no instances of this class
        }
    }
}
