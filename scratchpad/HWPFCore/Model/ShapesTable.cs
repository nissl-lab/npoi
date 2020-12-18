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

namespace NPOI.HWPF.Model
{


    using NPOI.HWPF.UserModel;
    using System.Collections;

    public class ShapesTable
    {
        private ArrayList _shapes;
        private ArrayList _shapesVisibili;  //holds visible shapes

        public ShapesTable(byte[] tblStream, FileInformationBlock fib)
        {
            PlexOfCps binTable = new PlexOfCps(tblStream,
                 fib.GetFcPlcspaMom(), fib.GetLcbPlcspaMom(), 26);

            _shapes = new ArrayList();
            _shapesVisibili = new ArrayList();


            for (int i = 0; i < binTable.Length; i++)
            {
                GenericPropertyNode nodo = binTable.GetProperty(i);

                Shape sh = new Shape(nodo);
                _shapes.Add(sh);
                if (sh.IsWithinDocument)
                    _shapesVisibili.Add(sh);
            }
        }

        public ArrayList GetAllShapes()
        {
            return _shapes;
        }

        public ArrayList GetVisibleShapes()
        {
            return _shapesVisibili;
        }
    }
}

