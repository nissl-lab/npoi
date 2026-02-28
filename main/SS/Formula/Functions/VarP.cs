/*
 *  ====================================================================
 *    Licensed to the collaborators of the NPOI project under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The collaborators license this file to You under the Apache License, Version 2.0
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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SS.Formula.Functions
{
    public sealed class VarP : NumberListFuncBase
    {
        public static readonly VarP Instance = new();
        private VarP()
        {
        }
        public override double CalculateFromNumberList(List<double> list)
        {
            if (list.Count == 1) return 0.0;
            var average = list.Average();
            var sum = 0.0;

            foreach (var item in list)
            {
                sum += Math.Pow(item - average, 2);
            }

            return sum / list.Count;
        }
    }
}
