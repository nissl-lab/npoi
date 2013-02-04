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
    using NPOI.SS.Formula.Eval;

    using System.Collections;
    using System.Reflection;
    using NPOI.SS.Formula.PTG;


    /**
     * This class Creates <c>OperationEval</c> instances to help evaluate <c>OperationPtg</c>
     * formula tokens.
     * 
     * @author Josh Micich
     */
    class OperationEvaluatorFactory
    {
        private static Type[] OPERATION_CONSTRUCTOR_CLASS_ARRAY = new Type[] { typeof(Ptg) };

        private static Hashtable _constructorsByPtgClass = InitialiseConstructorsMap();

        private OperationEvaluatorFactory()
        {
            // no instances of this class
        }

        private static Hashtable InitialiseConstructorsMap()
        {
            Hashtable m = new Hashtable(32);
            Add(m, typeof(AddPtg), typeof(AddEval));
            Add(m, typeof(DividePtg), typeof(DivideEval));
            Add(m, typeof(EqualPtg), typeof(EqualEval));
            Add(m, typeof(MultiplyPtg), typeof(MultiplyEval));
            Add(m, typeof(ConcatPtg), typeof(ConcatEval));

            Add(m, typeof(GreaterEqualPtg), typeof(GreaterEqualEval));
            Add(m, typeof(GreaterThanPtg), typeof(GreaterThanEval));
            Add(m, typeof(LessEqualPtg), typeof(LessEqualEval));
            Add(m, typeof(LessThanPtg), typeof(LessThanEval));
            Add(m, typeof(NotEqualPtg), typeof(NotEqualEval));
            Add(m, typeof(PercentPtg), typeof(PercentEval));
            Add(m, typeof(PowerPtg), typeof(PowerEval));
            Add(m, typeof(SubtractPtg), typeof(SubtractEval));
            Add(m, typeof(UnaryMinusPtg), typeof(UnaryMinusEval));
            Add(m, typeof(UnaryPlusPtg), typeof(UnaryPlusEval));
            return m;
        }

        private static void Add(Hashtable m, Type ptgClass, Type evalClass)
        {

            // perform some validation now, to keep later exception handlers simple
            if (!typeof(Ptg).IsAssignableFrom(ptgClass))
            {
                throw new ArgumentException("Expected Ptg subclass");
            }
            if (!typeof(OperationEval).IsAssignableFrom(evalClass))
            {
                throw new ArgumentException("Expected OperationEval subclass");
            }
            if (!evalClass.IsPublic)
            {
                throw new Exception("Eval class must be public");
            }
            if (evalClass.IsAbstract)
            {
                throw new Exception("Eval class must not be abstract");
            }

            ConstructorInfo constructor;
            try
            {
                constructor = evalClass.GetConstructor(OPERATION_CONSTRUCTOR_CLASS_ARRAY);
            }
            catch (Exception)
            {
                throw;
            }
            if (!constructor.IsPublic)
            {
                throw new Exception("Eval constructor must be public");
            }
            m[ptgClass] = constructor;
        }

        /**
         * returns the OperationEval concrete impl instance corresponding
         * to the supplied operationPtg
         */
        public static OperationEval Create(OperationPtg ptg)
        {
            if (ptg == null)
            {
                throw new ArgumentException("ptg must not be null");
            }

            Type ptgClass = ptg.GetType();

            ConstructorInfo constructor = (ConstructorInfo)_constructorsByPtgClass[ptgClass];
            if (constructor == null)
            {
                if (ptgClass == typeof(ExpPtg))
                {
                    // ExpPtg Is used for array formulas and shared formulas.
                    // it Is currently Unsupported, and may not even Get implemented here
                    throw new Exception("ExpPtg currently not supported");
                }
                throw new Exception("Unexpected operation ptg class (" + ptgClass.Name + ")");
            }

            Object result;
            Object[] initargs = { ptg };
            try
            {
                result = constructor.Invoke(initargs);
            }
            catch (Exception)
            {
                throw;
            }
            return (OperationEval)result;
        }
    }
}