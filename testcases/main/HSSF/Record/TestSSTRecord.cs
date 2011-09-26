/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF.Record
{
    using System;
    using System.IO;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.UserModel;
    using NPOI.Util;

    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using TestCases.HSSF;

    /**
     * @author Marc Johnson (mjohnson at apache dot org)
     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestClass]
    public class TestSSTRecord
    {

        /**
         * test processContinueRecord
         */
        public void TestProcessContinueRecord()
        {
            //jmh        byte[] testdata = HexRead.ReadData( _test_file_path + File.separator + "BigSSTRecord" );
            //jmh        byte[] input = new byte[testdata.Length - 4];
            //jmh
            //jmh        System.arraycopy( testdata, 4, input, 0, input.Length );
            //jmh        SSTRecord record =
            //jmh                new SSTRecord( LittleEndian.GetShort( testdata, 0 ),
            //jmh                        LittleEndian.GetShort( testdata, 2 ), input );
            //jmh        byte[] continueRecord = HexRead.ReadData( _test_file_path + File.separator + "BigSSTRecordCR" );
            //jmh
            //jmh        input = new byte[continueRecord.Length - 4];
            //jmh        System.arraycopy( continueRecord, 4, input, 0, input.Length );
            //jmh        record.processContinueRecord( input );
            //jmh        Assert.AreEqual( 1464, record.NumStrings );
            //jmh        Assert.AreEqual( 688, record.NumUniqueStrings );
            //jmh        Assert.AreEqual( 688, record.CountStrings );
            //jmh        byte[] ser_output = record.Serialize;
            //jmh        int offset = 0;
            //jmh        short type = LittleEndian.GetShort( ser_output, offset );
            //jmh
            //jmh        offset += LittleEndianConsts.SHORT_SIZE;
            //jmh        short length = LittleEndian.GetShort( ser_output, offset );
            //jmh
            //jmh        offset += LittleEndianConsts.SHORT_SIZE;
            //jmh        byte[] recordData = new byte[length];
            //jmh
            //jmh        System.arraycopy( ser_output, offset, recordData, 0, length );
            //jmh        offset += length;
            //jmh        SSTRecord testRecord = new SSTRecord( type, length, recordData );
            //jmh
            //jmh        Assert.AreEqual( ContinueRecord.sid,
            //jmh                LittleEndian.GetShort( ser_output, offset ) );
            //jmh        offset += LittleEndianConsts.SHORT_SIZE;
            //jmh        length = LittleEndian.GetShort( ser_output, offset );
            //jmh        offset += LittleEndianConsts.SHORT_SIZE;
            //jmh        byte[] cr = new byte[length];
            //jmh
            //jmh        System.arraycopy( ser_output, offset, cr, 0, length );
            //jmh        offset += length;
            //jmh        Assert.AreEqual( offset, ser_output.Length );
            //jmh        testRecord.processContinueRecord( cr );
            //jmh        Assert.AreEqual( record, testRecord );
            //jmh
            //jmh        // testing based on new bug report
            //jmh        testdata = HexRead.ReadData( _test_file_path + File.separator + "BigSSTRecord2" );
            //jmh        input = new byte[testdata.Length - 4];
            //jmh        System.arraycopy( testdata, 4, input, 0, input.Length );
            //jmh        record = new SSTRecord( LittleEndian.GetShort( testdata, 0 ),
            //jmh                LittleEndian.GetShort( testdata, 2 ), input );
            //jmh        byte[] continueRecord1 = HexRead.ReadData( _test_file_path + File.separator + "BigSSTRecord2CR1" );
            //jmh
            //jmh        input = new byte[continueRecord1.Length - 4];
            //jmh        System.arraycopy( continueRecord1, 4, input, 0, input.Length );
            //jmh        record.processContinueRecord( input );
            //jmh        byte[] continueRecord2 = HexRead.ReadData( _test_file_path + File.separator + "BigSSTRecord2CR2" );
            //jmh
            //jmh        input = new byte[continueRecord2.Length - 4];
            //jmh        System.arraycopy( continueRecord2, 4, input, 0, input.Length );
            //jmh        record.processContinueRecord( input );
            //jmh        byte[] continueRecord3 = HexRead.ReadData( _test_file_path + File.separator + "BigSSTRecord2CR3" );
            //jmh
            //jmh        input = new byte[continueRecord3.Length - 4];
            //jmh        System.arraycopy( continueRecord3, 4, input, 0, input.Length );
            //jmh        record.processContinueRecord( input );
            //jmh        byte[] continueRecord4 = HexRead.ReadData( _test_file_path + File.separator + "BigSSTRecord2CR4" );
            //jmh
            //jmh        input = new byte[continueRecord4.Length - 4];
            //jmh        System.arraycopy( continueRecord4, 4, input, 0, input.Length );
            //jmh        record.processContinueRecord( input );
            //jmh        byte[] continueRecord5 = HexRead.ReadData( _test_file_path + File.separator + "BigSSTRecord2CR5" );
            //jmh
            //jmh        input = new byte[continueRecord5.Length - 4];
            //jmh        System.arraycopy( continueRecord5, 4, input, 0, input.Length );
            //jmh        record.processContinueRecord( input );
            //jmh        byte[] continueRecord6 = HexRead.ReadData( _test_file_path + File.separator + "BigSSTRecord2CR6" );
            //jmh
            //jmh        input = new byte[continueRecord6.Length - 4];
            //jmh        System.arraycopy( continueRecord6, 4, input, 0, input.Length );
            //jmh        record.processContinueRecord( input );
            //jmh        byte[] continueRecord7 = HexRead.ReadData( _test_file_path + File.separator + "BigSSTRecord2CR7" );
            //jmh
            //jmh        input = new byte[continueRecord7.Length - 4];
            //jmh        System.arraycopy( continueRecord7, 4, input, 0, input.Length );
            //jmh        record.processContinueRecord( input );
            //jmh        Assert.AreEqual( 158642, record.NumStrings );
            //jmh        Assert.AreEqual( 5249, record.NumUniqueStrings );
            //jmh        Assert.AreEqual( 5249, record.CountStrings );
            //jmh        ser_output = record.Serialize;
            //jmh        offset = 0;
            //jmh        type = LittleEndian.GetShort( ser_output, offset );
            //jmh        offset += LittleEndianConsts.SHORT_SIZE;
            //jmh        length = LittleEndian.GetShort( ser_output, offset );
            //jmh        offset += LittleEndianConsts.SHORT_SIZE;
            //jmh        recordData = new byte[length];
            //jmh        System.arraycopy( ser_output, offset, recordData, 0, length );
            //jmh        offset += length;
            //jmh        testRecord = new SSTRecord( type, length, recordData );
            //jmh        for ( int Count = 0; Count < 7; Count++ )
            //jmh        {
            //jmh            Assert.AreEqual( ContinueRecord.sid,
            //jmh                    LittleEndian.GetShort( ser_output, offset ) );
            //jmh            offset += LittleEndianConsts.SHORT_SIZE;
            //jmh            length = LittleEndian.GetShort( ser_output, offset );
            //jmh            offset += LittleEndianConsts.SHORT_SIZE;
            //jmh            cr = new byte[length];
            //jmh            System.arraycopy( ser_output, offset, cr, 0, length );
            //jmh            testRecord.processContinueRecord( cr );
            //jmh            offset += length;
            //jmh        }
            //jmh        Assert.AreEqual( offset, ser_output.Length );
            //jmh        Assert.AreEqual( record, testRecord );
            //jmh        Assert.AreEqual( record.CountStrings, testRecord.CountStrings );
        }

        /**
         * Test capability of handling mondo big strings
         *
         * @exception IOException
         */
        private string ConvertByteArrayToString(byte[] bstring)
        {
            char[] chararray = new char[bstring.Length];
            for (int i = 0; i < bstring.Length; i++)
            {
                chararray[i] = Convert.ToChar(bstring[i]);
            }
            return new string(chararray);
        }

        [TestMethod]
        public void TestHugeStrings()
        {
            SSTRecord record = new SSTRecord();
            

            byte[][] bstrings =
                    {
                        new byte[9000], new byte[7433], new byte[9002],
                        new byte[16998]
                    };
            UnicodeString[] strings = new UnicodeString[bstrings.Length];
            int total_length = 0;

            for ( int k = 0; k < bstrings.Length; k++ )
            {
                Arrays.Fill(bstrings[k], (byte) ( Convert.ToInt32('a') + k ) );

                
                strings[k] = new UnicodeString(ConvertByteArrayToString( bstrings[k]) );
                record.AddString( strings[k] );
                total_length += 3 + bstrings[k].Length;
            }

            // add overhead of SST record
            total_length += 8;

            // add overhead of broken strings
            total_length += 4;

            // add overhead of six records
            total_length += ( 6 * 4 );
            byte[] content = new byte[record.RecordSize];

            record.Serialize( 0, content );
            Assert.AreEqual( total_length, content.Length );

            //DeSerialize the record.
            RecordInputStream recStream = new RecordInputStream(new MemoryStream(content));
            recStream.NextRecord();
            record = new SSTRecord(recStream);

            Assert.AreEqual( strings.Length, record.NumStrings );
            Assert.AreEqual( strings.Length, record.NumUniqueStrings );
            Assert.AreEqual( strings.Length, record.CountStrings );
            for ( int k = 0; k < strings.Length; k++ )
            {
                Assert.AreEqual( strings[k], record.GetString( k ) );
            }
            record = new SSTRecord();
            bstrings[1] = new byte[bstrings[1].Length - 1];
            for ( int k = 0; k < bstrings.Length; k++ )
            {
                if ( ( bstrings[k].Length % 2 ) == 1 )
                {
                    Arrays.Fill( bstrings[k], (byte) ( 'a' + k ) );
                    strings[k] = new UnicodeString( ConvertByteArrayToString(bstrings[k]) );
                }
                else
                {
                    char[] data = new char[bstrings[k].Length / 2];

                    Arrays.Fill( data, (char) ( Convert.ToInt32('\u2122') + k)) ;
                    strings[k] = new UnicodeString(new String( data ));
                }
                record.AddString( strings[k] );
            }
            content = new byte[record.RecordSize];
            record.Serialize( 0, content );
            total_length--;
            Assert.AreEqual( total_length, content.Length );

            recStream = new RecordInputStream(new MemoryStream(content));
            recStream.NextRecord();
            record = new SSTRecord(recStream);

            Assert.AreEqual( strings.Length, record.NumStrings );
            Assert.AreEqual( strings.Length, record.NumUniqueStrings );
            Assert.AreEqual( strings.Length, record.CountStrings );
            for ( int k = 0; k < strings.Length; k++ )
            {
                Assert.AreEqual( strings[k], record.GetString( k ) );
            }
        }

        /**
         * test SSTRecord boundary conditions
         */
        [TestMethod]
        public void TestSSTRecordBug()
        {
            // create an SSTRecord and Write a certain pattern of strings
            // to it ... then Serialize it and verify the content
            SSTRecord record = new SSTRecord();

            // the record will start with two integers, then this string
            // ... that will eat up 16 of the 8224 bytes that the record
            // can hold
            record.AddString( new UnicodeString("Hello") );

            // now we have an additional 8208 bytes, which is1 an exact
            // multiple of 16 bytes
            long testvalue = 1000000000000L;

            for ( int k = 0; k < 2000; k++ )
            {
                record.AddString( new UnicodeString((testvalue++).ToString()) );
            }
            byte[] content = new byte[record.RecordSize];

            record.Serialize( 0, content );
            Assert.AreEqual(8224, LittleEndian.GetShort(content, 2));
            Assert.AreEqual(ContinueRecord.sid, LittleEndian.GetShort(content, 8228));
            Assert.AreEqual(8224, LittleEndian.GetShort(content, 8228+2));
            Assert.AreEqual( (byte) 13, content[4 + 8228] );
            Assert.AreEqual(ContinueRecord.sid, LittleEndian.GetShort(content, 2*8228));
            Assert.AreEqual(8224, LittleEndian.GetShort(content, 8228*2+2));
            Assert.AreEqual( (byte) 13, content[4 + 8228 * 2] );
            Assert.AreEqual(ContinueRecord.sid, LittleEndian.GetShort(content, 3*8228));
            Assert.AreEqual( (byte) 13, content[4 + 8228 * 3] );
        }

        /**
         * test simple addString
         */
        [TestMethod]
        public void TestSimpleAddString()
        {
            SSTRecord record = new SSTRecord();
            UnicodeString s1 = new UnicodeString("Hello world");

            // \u2122 is1 the encoding of the trademark symbol ...
            UnicodeString s2 = new UnicodeString("Hello world\u2122");

            Assert.AreEqual( 0, record.AddString( s1 ) );
            Assert.AreEqual( s1, record.GetString( 0 ) );
            Assert.AreEqual( 1, record.CountStrings );
            Assert.AreEqual( 1, record.NumStrings );
            Assert.AreEqual( 1, record.NumUniqueStrings );
            Assert.AreEqual( 0, record.AddString( s1 ) );
            Assert.AreEqual( s1, record.GetString( 0 ) );
            Assert.AreEqual( 1, record.CountStrings );
            Assert.AreEqual( 2, record.NumStrings );
            Assert.AreEqual( 1, record.NumUniqueStrings );
            Assert.AreEqual( 1, record.AddString( s2 ) );
            Assert.AreEqual( s2, record.GetString( 1 ) );
            Assert.AreEqual( 2, record.CountStrings );
            Assert.AreEqual( 3, record.NumStrings );
            Assert.AreEqual( 2, record.NumUniqueStrings );
            System.Collections.IEnumerator iter = record.GetStrings();

            while ( iter.MoveNext() )
            {
                UnicodeString ucs = (UnicodeString) iter.Current;

                if ( ucs.Equals( s1 ) )
                {
                    Assert.AreEqual( (byte) 0, ucs.OptionFlags );
                }
                else if ( ucs.Equals( s2 ) )
                {
                    Assert.AreEqual( (byte) 1, ucs.OptionFlags );
                }
                else
                {
                    Assert.Fail( "cannot match string: " + ucs.String );
                }
            }
        }

        /**
         * test simple constructor
         */
        [TestMethod]
        public void TestSimpleConstructor()
        {
            SSTRecord record = new SSTRecord();

            Assert.AreEqual(0, record.NumStrings);
            Assert.AreEqual(0, record.NumUniqueStrings);
            Assert.AreEqual(0, record.CountStrings);
            byte[] output = record.Serialize();
            byte[] expected =
                {
                    (byte) record.Sid, (byte) ( record.Sid >> 8 ),
                    (byte) 8, (byte) 0, (byte) 0, (byte) 0, (byte) 0,
                    (byte) 0, (byte) 0, (byte) 0, (byte) 0, (byte) 0
                };

            Assert.AreEqual(expected.Length, output.Length);
            for (int k = 0; k < expected.Length; k++)
            {
                Assert.AreEqual(expected[k], output[k], k.ToString());
            }
        }

        /**
         * Tests that workbooks with rich text that duplicates a non rich text cell can be Read and written.
         */
        [TestMethod]
        public void TestReadWriteDuplicatedRichText1()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("duprich1.xls");
            NPOI.SS.UserModel.ISheet sheet = wb.GetSheetAt(1);
            Assert.AreEqual("01/05 (Wed)", sheet.GetRow(0).GetCell((short)8).StringCellValue);
            Assert.AreEqual("01/05 (Wed)", sheet.GetRow(1).GetCell((short)8).StringCellValue);

            MemoryStream baos = new MemoryStream();
            wb.Write(baos);

            // test the second file.
            wb = HSSFTestDataSamples.OpenSampleWorkbook("duprich2.xls");
            sheet = wb.GetSheetAt(0);
            int row = 0;
            Assert.AreEqual("Testing", sheet.GetRow(row++).GetCell((short)0).StringCellValue);
            Assert.AreEqual("rich", sheet.GetRow(row++).GetCell((short)0).StringCellValue);
            Assert.AreEqual("text", sheet.GetRow(row++).GetCell((short)0).StringCellValue);
            Assert.AreEqual("strings", sheet.GetRow(row++).GetCell((short)0).StringCellValue);
            Assert.AreEqual("Testing", sheet.GetRow(row++).GetCell((short)0).StringCellValue);
            Assert.AreEqual("Testing", sheet.GetRow(row++).GetCell((short)0).StringCellValue);

            wb.Write(baos);
        }
    }
}
