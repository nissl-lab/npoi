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
    using NPOI.DDF;
    using System;

    public class HSSFChildAnchor : HSSFAnchor
    {
        private EscherChildAnchorRecord _escherChildAnchor;
        /**
         * create anchor from existing file
         * @param escherChildAnchorRecord
         */
        public HSSFChildAnchor(EscherChildAnchorRecord escherChildAnchorRecord)
        {
            this._escherChildAnchor = escherChildAnchorRecord;
        }
        public HSSFChildAnchor()
        {
            _escherChildAnchor = new EscherChildAnchorRecord();
        }
        /**
        * create anchor from scratch
        * @param dx1 x coordinate of the left up corner
        * @param dy1 y coordinate of the left up corner
        * @param dx2 x coordinate of the right down corner
        * @param dy2 y coordinate of the right down corner
        */
        public HSSFChildAnchor(int dx1, int dy1, int dx2, int dy2)
            : base(Math.Min(dx1, dx2), Math.Min(dy1, dy2), Math.Max(dx1, dx2), Math.Max(dy1, dy2))
        {
            if (dx1 > dx2)
            {
                _isHorizontallyFlipped = true;
            }
            if (dy1 > dy2)
            {
                _isVerticallyFlipped = true;
            }
        }
        /**
         * @param dx1 x coordinate of the left up corner
         * @param dy1 y coordinate of the left up corner
         * @param dx2 x coordinate of the right down corner
         * @param dy2 y coordinate of the right down corner
         */
        public void SetAnchor(int dx1, int dy1, int dx2, int dy2)
        {
            this.Dx1 = Math.Min(dx1, dx2);
            this.Dy1 = Math.Min(dy1, dy2);
            this.Dx2 = Math.Max(dx1, dx2);
            this.Dy2 = Math.Max(dy1, dy2);
        }

        public override bool IsHorizontallyFlipped
        {
            get 
            { 
                return _isHorizontallyFlipped; 
            }
        }

        public override bool IsVerticallyFlipped
        {
            get 
            { 
                return _isVerticallyFlipped; 
            }
        }
        public override int Dx1
        {
            get
            {
                return _escherChildAnchor.Dx1;
            }
            set
            {
                _escherChildAnchor.Dx1 = (short)value;
            }
        }
        public override int Dx2
        {
            get
            {
                return _escherChildAnchor.Dx2;
            }
            set
            {
                _escherChildAnchor.Dx2 = (short)value;
            }
        }
        public override int Dy1
        {
            get
            {
                return _escherChildAnchor.Dy1;
            }
            set
            {
                _escherChildAnchor.Dy1 = (short)value;
            }
        }
        public override int Dy2
        {
            get
            {
                return _escherChildAnchor.Dy2;
            }
            set
            {
                _escherChildAnchor.Dy2 = (short)value;
            }
        }
        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;
            if (obj == this)
                return true;
            if (obj.GetType() != GetType())
                return false;
            HSSFChildAnchor anchor = (HSSFChildAnchor)obj;

            return anchor.Dx1 == Dx1 && anchor.Dx2 == Dx2 && anchor.Dy1 == Dy1
                    && anchor.Dy2 == Dy2;
        }

        public override int GetHashCode()
        {
            return Dx1.GetHashCode() ^ Dx2.GetHashCode() ^ Dy1.GetHashCode()
                    ^ Dy2.GetHashCode();
        }

        internal override EscherRecord GetEscherAnchor()
        {
            return _escherChildAnchor;
        }
        protected override void CreateEscherAnchor()
        {
            _escherChildAnchor = new EscherChildAnchorRecord();
        }
    }
}