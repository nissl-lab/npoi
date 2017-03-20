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

using System;
using System.IO;
using System.Threading;
using NUnit.Framework;
namespace TestCases
{
    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    public sealed class RunSerialyAndSweepTmpFilesAttribute : Attribute, ITestAction
    {
        private static object syncSequential = new object();

        public ActionTargets Targets { get { return ActionTargets.Test; } }

        public void BeforeTest(TestDetails testDetails)
        {
            Monitor.Enter(syncSequential);
            SweepTemporaryFiles();
        }

        public void AfterTest(TestDetails testDetails)
        {
            SweepTemporaryFiles();
            Monitor.Exit(syncSequential);
        }

        private static void SweepTemporaryFiles()
        {
            foreach (var tempFilePath in Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, "*.tmp"))
            {
                File.Delete(tempFilePath);
            }
        }
    }
}
