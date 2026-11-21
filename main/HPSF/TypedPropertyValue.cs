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

using System;
using NPOI.Util;

namespace NPOI.HPSF
{
    public class TypedPropertyValue
    {
        private static POILogger logger = POILogFactory
                .GetLogger(typeof(TypedPropertyValue));

        private int _type;

        private Object _value;

        public TypedPropertyValue(int type, Object value)
        {
            _type = type;
            _value = value;
        }

        public Object Value
        {
            get { return _value; }
        }

        internal void Read( LittleEndianByteArrayInputStream lei ) {
            _type = lei.ReadShort();
            short padding = lei.ReadShort();
            if ( padding != 0 ) {
                //LOG.log( POILogger.WARN, "TypedPropertyValue padding at offset "
                //        + lei.getReadIndex() + " MUST be 0, but it's value is " + padding );
            }
            ReadValue( lei );
        }

        internal void ReadValue(LittleEndianByteArrayInputStream lei )
        {
            switch (_type)
            {
                case Variant.VT_EMPTY:
                case Variant.VT_NULL:
                    _value = null;
                    break;

                case Variant.VT_I1:
                    _value = lei.ReadByte();
                    break;

                case Variant.VT_UI1:
                    _value =  lei.ReadUByte();
                    break;

                case Variant.VT_I2:
                    _value = lei.ReadShort();
                    break;

                case Variant.VT_UI2:
                    _value = lei.ReadUShort();
                    break;
            
                case Variant.VT_INT:
                case Variant.VT_I4:
                    _value = lei.ReadInt();
                    break;

                case Variant.VT_UINT:
                case Variant.VT_UI4:
                case Variant.VT_ERROR:
                    _value = lei.ReadUInt();
                    break;

                case Variant.VT_I8:
                    _value = lei.ReadLong();
                    break;

                case Variant.VT_UI8: {
                    byte[] biBytesLE = new byte[LittleEndianConsts.LONG_SIZE];
                    lei.ReadFully(biBytesLE);

                    // first byte needs to be 0 for unsigned BigInteger
                    byte[] biBytesBE = new byte[9];
                    int i = biBytesLE.Length;
                    foreach (byte b in biBytesLE)
                    {
                        if (i<=8) 
                        {
                            biBytesBE[i] = b;
                        }
                        i--;
                    }
                    _value = new BigInteger(biBytesBE);
                    break;
                }

            
                case Variant.VT_R4:
                    byte[] b4 = new byte[LittleEndianConsts.INT_SIZE];
                    lei.ReadFully(b4);
                    _value = BitConverter.ToSingle(b4, 0); //Float.IntBitsToFloat(lei.ReadInt());
                    break;
            
                case Variant.VT_R8:
                    _value = lei.ReadDouble();
                    break;

                case Variant.VT_CY:
                    Currency cur = new Currency();
                    cur.Read(lei);
                    _value = cur;
                    break;
            

                case Variant.VT_DATE:
                    Date date = new Date();
                    date.Read(lei);
                    _value = date;
                    break;

                case Variant.VT_BSTR:
                case Variant.VT_LPSTR:
                    CodePageString cps = new CodePageString();
                    cps.Read(lei);
                    _value = cps;
                    break;

                case Variant.VT_BOOL:
                    VariantBool vb = new VariantBool();
                    vb.Read(lei);
                    _value = vb;
                    break;

                case Variant.VT_DECIMAL:
                    Decimal dec = new Decimal();
                    dec.Read(lei);
                    _value = dec;
                    break;

                case Variant.VT_LPWSTR:
                    UnicodeString us = new UnicodeString();
                    us.Read(lei);
                    _value = us;
                    break;

                case Variant.VT_FILETIME:
                    Filetime ft = new Filetime();
                    ft.Read(lei);
                    _value = ft;
                    break;

                case Variant.VT_BLOB:
                case Variant.VT_BLOB_OBJECT:
                    Blob blob = new Blob();
                    blob.Read(lei);
                    _value = blob;
                    break;

                case Variant.VT_STREAM:
                case Variant.VT_STORAGE:
                case Variant.VT_STREAMED_OBJECT:
                case Variant.VT_STORED_OBJECT:
                    IndirectPropertyName ipn = new IndirectPropertyName();
                    ipn.Read(lei);
                    _value = ipn;
                    break;

                case Variant.VT_CF:
                    ClipboardData cd = new ClipboardData();
                    cd.Read(lei);
                    _value = cd;
                    break;

                case Variant.VT_CLSID:
                    GUID guid = new GUID();
                    guid.Read(lei);
                    _value = lei;
                    break;

                case Variant.VT_VERSIONED_STREAM:
                    VersionedStream vs = new VersionedStream();
                    vs.Read(lei);
                    _value = vs;
                    break;

                case Variant.VT_VECTOR | Variant.VT_I2:
                case Variant.VT_VECTOR | Variant.VT_I4:
                case Variant.VT_VECTOR | Variant.VT_R4:
                case Variant.VT_VECTOR | Variant.VT_R8:
                case Variant.VT_VECTOR | Variant.VT_CY:
                case Variant.VT_VECTOR | Variant.VT_DATE:
                case Variant.VT_VECTOR | Variant.VT_BSTR:
                case Variant.VT_VECTOR | Variant.VT_ERROR:
                case Variant.VT_VECTOR | Variant.VT_BOOL:
                case Variant.VT_VECTOR | Variant.VT_VARIANT:
                case Variant.VT_VECTOR | Variant.VT_I1:
                case Variant.VT_VECTOR | Variant.VT_UI1:
                case Variant.VT_VECTOR | Variant.VT_UI2:
                case Variant.VT_VECTOR | Variant.VT_UI4:
                case Variant.VT_VECTOR | Variant.VT_I8:
                case Variant.VT_VECTOR | Variant.VT_UI8:
                case Variant.VT_VECTOR | Variant.VT_LPSTR:
                case Variant.VT_VECTOR | Variant.VT_LPWSTR:
                case Variant.VT_VECTOR | Variant.VT_FILETIME:
                case Variant.VT_VECTOR | Variant.VT_CF:
                case Variant.VT_VECTOR | Variant.VT_CLSID:
                    Vector vec = new Vector( (short) ( _type & 0x0FFF ) );
                    vec.Read(lei);
                    _value = vec;
                    break;

                case Variant.VT_ARRAY | Variant.VT_I2:
                case Variant.VT_ARRAY | Variant.VT_I4:
                case Variant.VT_ARRAY | Variant.VT_R4:
                case Variant.VT_ARRAY | Variant.VT_R8:
                case Variant.VT_ARRAY | Variant.VT_CY:
                case Variant.VT_ARRAY | Variant.VT_DATE:
                case Variant.VT_ARRAY | Variant.VT_BSTR:
                case Variant.VT_ARRAY | Variant.VT_ERROR:
                case Variant.VT_ARRAY | Variant.VT_BOOL:
                case Variant.VT_ARRAY | Variant.VT_VARIANT:
                case Variant.VT_ARRAY | Variant.VT_DECIMAL:
                case Variant.VT_ARRAY | Variant.VT_I1:
                case Variant.VT_ARRAY | Variant.VT_UI1:
                case Variant.VT_ARRAY | Variant.VT_UI2:
                case Variant.VT_ARRAY | Variant.VT_UI4:
                case Variant.VT_ARRAY | Variant.VT_INT:
                case Variant.VT_ARRAY | Variant.VT_UINT:
                    Array arr = new Array();
                    arr.Read(lei);
                    _value = arr;
                    break;

                default:
                    throw new InvalidOperationException(
                            "Unknown (possibly, incorrect) TypedPropertyValue type: "
                                    + _type);
            }
        }

        internal static void SkipPadding( LittleEndianByteArrayInputStream lei )
        {
            int offset = lei.GetReadIndex();
            int skipBytes = (4 - (offset & 3)) & 3;
            for (int i=0; i<skipBytes; i++) {
                lei.Mark(1);
                int b = lei.Read();
                if (b == -1 || b != 0) {
                    lei.Reset();
                    break;
                }
            }
        }
    }

}