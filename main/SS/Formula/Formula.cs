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

namespace NPOI.SS.Formula
{
    using NPOI.SS.Formula;
    using NPOI.Util;

    using NPOI.SS.Util;
    using NPOI.SS.Formula.PTG;

    /// <summary>
    /// Encapsulates an encoded formula token array.
    /// </summary>
    /// @author Josh Micich
    public class Formula
    {

        private static readonly Formula EMPTY = new Formula(new byte[0], 0);

        /// <summary>
        /// immutable */
        /// </summary>
        private byte[] _byteEncoding;
        private int _encodedTokenLen;

        private Formula(byte[] byteEncoding, int encodedTokenLen)
        {
            _byteEncoding = (byte[]) byteEncoding.Clone();
            _encodedTokenLen = encodedTokenLen;
            //if (false) { // set to true to eagerly check Ptg decoding 
            //    LittleEndianByteArrayInputStream in1 = new LittleEndianByteArrayInputStream(byteEncoding);
            //    Ptg.ReadTokens(encodedTokenLen, in1);
            //    int nUnusedBytes = _byteEncoding.Length - in1.GetReadIndex();
            //    if (nUnusedBytes > 0) {
            //        // TODO - this seems to occur when IntersectionPtg is present
            //        // This example file "IntersectionPtg.xls"
            //        // used by test: TestIntersectionPtg.testReading()
            //        // has 10 bytes unused at the end of the formula
            //        // 10 extra bytes are just 0x01 and 0x00
            //        Console.WriteLine(nUnusedBytes + " unused bytes at end of formula");
            //    }
            //}
        }
        /// <summary>
        /// Convenience method for <see cref="read(int, LittleEndianInput, int)" />
        /// </summary>
        public static Formula Read(int encodedTokenLen, ILittleEndianInput in1)
        {
            return Read(encodedTokenLen, in1, encodedTokenLen);
        }
        /// <summary>
        /// When there are no array constants present, <c>encodedTokenLen</c>==<c>totalEncodedLen</c>
        /// </summary>
        /// <param name="encodedTokenLen">number of bytes in the stream taken by the plain formula tokens</param>
        /// <param name="totalEncodedLen">the total number of bytes in the formula (includes trailing encoding
        /// for array constants, but does not include 2 bytes for initial <c>ushort encodedTokenLen</c> field.
        /// </param>
        /// <returns>A new formula object as read from the stream.  Possibly empty, never <c>null</c>.</returns>
        public static Formula Read(int encodedTokenLen, ILittleEndianInput in1, int totalEncodedLen)
        {
            byte[] byteEncoding = new byte[totalEncodedLen];
            in1.ReadFully(byteEncoding);
            return new Formula(byteEncoding, encodedTokenLen);
        }

        public Ptg[] Tokens
        {
            get
            {
                ILittleEndianInput in1 = new LittleEndianByteArrayInputStream(_byteEncoding);
                return Ptg.ReadTokens(_encodedTokenLen, in1);
            }
        }
        /// <summary>
        /// Writes  The formula encoding is includes:
        /// <list type="bullet">
        /// <item><description>ushort tokenDataLen</description></item>
        /// <item><description>tokenData</description></item>
        /// <item><description>arrayConstantData (if present)</description></item>
        /// </list>
        /// </summary>
        public void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(_encodedTokenLen);
            out1.Write(_byteEncoding);
        }

        public void SerializeTokens(ILittleEndianOutput out1)
        {
            out1.Write(_byteEncoding, 0, _encodedTokenLen);
        }
        public void SerializeArrayConstantData(ILittleEndianOutput out1)
        {
            int len = _byteEncoding.Length - _encodedTokenLen;
            out1.Write(_byteEncoding, _encodedTokenLen, len);
        }


        /// <summary>
        /// </summary>
        /// <returns>total formula encoding length.  The formula encoding includes:
        /// <list type="bullet">
        /// <item><description>ushort tokenDataLen</description></item>
        /// <item><description>tokenData</description></item>
        /// <item><description>arrayConstantData (optional)</description></item>
        /// </list>
        /// Note - this value is different to <c>tokenDataLength</c>
        /// </returns>
        public int EncodedSize
        {
            get
            {
                return 2 + _byteEncoding.Length;
            }
        }
        /// <summary>
        /// This method is often used when the formula length does not appear immediately before
        /// the encoded token data.
        /// </summary>
        /// <returns>the encoded length of the plain formula tokens.  This does <em>not</em> include
        /// the leading ushort field, nor any trailing array constant data.
        /// </returns>
        public int EncodedTokenSize
        {
            get
            {
                return _encodedTokenLen;
            }
        }

        /// <summary>
        /// Creates a <see cref="Formula"/> object from a supplied <see cref="Ptg"/> array.
        /// Handles <c>null</c>s OK.
        /// </summary>
        /// <param name="ptgs">may be <c>null</c></param>
        /// <returns>Never <c>null</c> (Possibly empty if the supplied <c>ptgs</c> is <c>null</c>)</returns>
        public static Formula Create(Ptg[] ptgs)
        {
            if(ptgs == null || ptgs.Length < 1)
            {
                return EMPTY;
            }
            int totalSize = Ptg.GetEncodedSize(ptgs);
            byte[] encodedData = new byte[totalSize];
            Ptg.SerializePtgs(ptgs, encodedData, 0);
            int encodedTokenLen = Ptg.GetEncodedSizeWithoutArrayData(ptgs);
            return new Formula(encodedData, encodedTokenLen);
        }
        /// <summary>
        /// Gets the <see cref="Ptg"/> array from the supplied <see cref="Formula"/>.
        /// Handles <c>null</c>s OK.
        /// </summary>
        /// <param name="formula">may be <c>null</c></param>
        /// <returns>possibly <c>null</c> (if the supplied <c>formula</c> is <c>null</c>)</returns>
        public static Ptg[] GetTokens(Formula formula)
        {
            if(formula == null)
            {
                return null;
            }
            return formula.Tokens;
        }

        public Formula Copy()
        {
            // OK to return this because immutable
            return this;
        }

        /// <summary>
        /// Gets the locator for the corresponding <see cref="SharedFormulaRecord"/>, <see cref="ArrayRecord"/> or
        /// <see cref="TableRecord"/> if this formula belongs to such a grouping.  The <see cref="CellReference"/>
        /// returned by this method will  match the top left corner of the range of that grouping.
        /// The return value is usually not the same as the location of the cell containing this formula.
        /// </summary>
        /// <returns>the firstRow &amp; firstColumn of an array formula or shared formula that this formula
        /// belongs to.  <c>null</c> if this formula is not part of an array or shared formula.
        /// </returns>
        public CellReference ExpReference
        {
            get
            {
                byte[] data = _byteEncoding;
                if(data.Length != 5)
                {
                    // tExp and tTbl are always 5 bytes long, and the only ptg in the formula
                    return null;
                }
                switch(data[0])
                {
                    case ExpPtg.sid:
                        break;
                    case TblPtg.sid:
                        break;
                    default:
                        return null;
                }
                int firstRow = LittleEndian.GetUShort(data, 1);
                int firstColumn = LittleEndian.GetUShort(data, 3);
                return new CellReference(firstRow, firstColumn);
            }
        }
        public bool IsSame(Formula other)
        {
            return Arrays.Equals(_byteEncoding, other._byteEncoding);
        }
    }
}