/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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


namespace TestCases.HWPF
{
    using NPOI.POIFS.FileSystem;
    using NPOI.HWPF.Model;
    using System;

    public class HWPFDocFixture
    {
        public byte[] _tableStream;
        public byte[] _mainStream;
        public FileInformationBlock _fib;

        public HWPFDocFixture(Object obj)
        {

        }

        public void SetUp()
        {
            POIFSFileSystem filesystem = new POIFSFileSystem(
                    POIDataSamples.GetDocumentInstance().OpenResourceAsStream("test.doc"));

            DocumentEntry documentProps =
              (DocumentEntry)filesystem.Root.GetEntry("WordDocument");
            _mainStream = new byte[documentProps.Size];
            filesystem.CreateDocumentInputStream("WordDocument").Read(_mainStream);

            // use the fib to determine the name of the table stream.
            _fib = new FileInformationBlock(_mainStream);

            String name = "0Table";
            if (_fib.IsFWhichTblStm())
            {
                name = "1Table";
            }

            // read in the table stream.
            DocumentEntry tableProps =
              (DocumentEntry)filesystem.Root.GetEntry(name);
            _tableStream = new byte[tableProps.Size];
            filesystem.CreateDocumentInputStream(name).Read(_tableStream);

            _fib.FillVariableFields(_mainStream, _tableStream);

        }

    }
}

