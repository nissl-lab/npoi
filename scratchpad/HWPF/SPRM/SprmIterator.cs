
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License Is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */


namespace NPOI.HWPF.SPRM
{
    using System;
    using System.Collections;


    /**
     * This class Is used to iterate through a list of sprms from a Word 97/2000/XP
     * document.
     * @author Ryan Ackley
     * @version 1.0
     */

    public class SprmIterator
    {
        private byte[] _grpprl;
        int _offset;

        public SprmIterator(byte[] grpprl, int offset)
        {
            _grpprl = grpprl;
            _offset = offset;
        }

        public bool HasNext()
        {
            // A Sprm is at least 2 bytes long
            return _offset < _grpprl.Length-1;
        }

        public SprmOperation Next()
        {
            SprmOperation op = new SprmOperation(_grpprl, _offset);
            _offset += op.Size;
            return op;
        }


    }

}