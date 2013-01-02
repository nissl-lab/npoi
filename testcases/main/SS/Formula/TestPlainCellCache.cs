/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for Additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
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

namespace TestCases.SS.Formula
{

    using NUnit.Framework;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;




    /**
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestPlainCellCache
    {

        /**
         *
         */
        [Test]
        public void TestLoc()
        {
            PlainCellCache cache = new PlainCellCache();
            for (int bookIndex = 0; bookIndex < 0x1000; bookIndex += 0x100)
            {
                for (int sheetIndex = 0; sheetIndex < 0x1000; sheetIndex += 0x100)
                {
                    for (int rowIndex = 0; rowIndex < 0x100000; rowIndex += 0x1000)
                    {
                        for (int columnIndex = 0; columnIndex < 0x4000; columnIndex += 0x100)
                        {
                            Loc loc = new Loc(bookIndex, sheetIndex, rowIndex, columnIndex);
                            Assert.AreEqual(bookIndex, loc.BookIndex);
                            Assert.AreEqual(sheetIndex, loc.SheetIndex);
                            Assert.AreEqual(rowIndex, loc.RowIndex);
                            Assert.AreEqual(columnIndex, loc.ColumnIndex);

                            Loc sameLoc = new Loc(bookIndex, sheetIndex, rowIndex, columnIndex);
                            Assert.AreEqual(loc.GetHashCode(), sameLoc.GetHashCode());
                            Assert.IsTrue(loc.Equals(sameLoc));

                            Assert.IsNull(cache.Get(loc));
                            PlainValueCellCacheEntry entry = new PlainValueCellCacheEntry(new NumberEval(0));
                            cache.Put(loc, entry);

                            Assert.AreSame(entry, cache.Get(loc));
                            cache.Remove(loc);
                            Assert.IsNull(cache.Get(loc));

                            cache.Put(loc, entry);
                        }
                        cache.Clear();
                    }
                }

            }
        }
    }

}