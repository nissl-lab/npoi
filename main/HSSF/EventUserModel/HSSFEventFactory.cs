
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

namespace NPOI.HSSF.EventUserModel
{
    using System.IO;
    using NPOI.POIFS.FileSystem;
    using NPOI.HSSF.Record;


    /// <summary>
    /// Low level event based HSSF Reader.  Pass either a DocumentInputStream to
    /// Process events along with a request object or pass a POIFS POIFSFileSystem to
    /// ProcessWorkbookEvents along with a request.
    /// This will cause your file to be Processed a record at a time.  Each record with
    /// a static id matching one that you have registed in your HSSFRequest will be passed
    /// to your associated HSSFListener.
    /// @author Andrew C. Oliver (acoliver at apache dot org)
    /// @author Carey Sublette  (careysub@earthling.net)
    /// </summary>
    public class HSSFEventFactory
    {
        /// <summary>
        /// Creates a new instance of HSSFEventFactory
        /// </summary>
        public HSSFEventFactory()
        {
        }

        /// <summary>
        /// Processes a file into essentially record events.
        /// </summary>
        /// <param name="req">an Instance of HSSFRequest which has your registered listeners</param>
        /// <param name="fs">a POIFS filesystem containing your workbook</param>
        public void ProcessWorkbookEvents(HSSFRequest req, POIFSFileSystem fs)
        {
            Stream in1 = fs.CreateDocumentInputStream("Workbook");

            ProcessEvents(req, in1);
        }

        /// <summary>
        /// Processes a file into essentially record events.
        /// </summary>
        /// <param name="req">an Instance of HSSFRequest which has your registered listeners</param>
        /// <param name="fs">a POIFS filesystem containing your workbook</param>
        /// <returns>numeric user-specified result code.</returns>
        public short AbortableProcessWorkbookEvents(HSSFRequest req, POIFSFileSystem fs)
        {
            Stream in1 = fs.CreateDocumentInputStream("Workbook");
            return AbortableProcessEvents(req, in1);
        }

        /// <summary>
        /// Processes a DocumentInputStream into essentially Record events.
        /// If an 
        /// <c>AbortableHSSFListener</c>
        ///  causes a halt to Processing during this call
        /// the method will return just as with 
        /// <c>abortableProcessEvents</c>
        /// , but no
        /// user code or 
        /// <c>HSSFUserException</c>
        ///  will be passed back.
        /// </summary>
        /// <param name="req">an Instance of HSSFRequest which has your registered listeners</param>
        /// <param name="in1">a DocumentInputStream obtained from POIFS's POIFSFileSystem object</param>
        public void ProcessEvents(HSSFRequest req, Stream in1)
        {
            try
            {
                GenericProcessEvents(req, new RecordInputStream(in1));
            }
            catch (HSSFUserException)
            {/*If an HSSFUserException user exception Is thrown, ignore it.*/ }
        }


        /// <summary>
        /// Processes a DocumentInputStream into essentially Record events.
        /// </summary>
        /// <param name="req">an Instance of HSSFRequest which has your registered listeners</param>
        /// <param name="in1">a DocumentInputStream obtained from POIFS's POIFSFileSystem object</param>
        /// <returns>numeric user-specified result code.</returns>
        public short AbortableProcessEvents(HSSFRequest req, Stream in1)
        {
            return GenericProcessEvents(req, new RecordInputStream(in1));
        }

        /// <summary>
        /// Processes a DocumentInputStream into essentially Record events.
        /// </summary>
        /// <param name="req">an Instance of HSSFRequest which has your registered listeners</param>
        /// <param name="in1">a DocumentInputStream obtained from POIFS's POIFSFileSystem object</param>
        /// <returns>numeric user-specified result code.</returns>
        protected short GenericProcessEvents(HSSFRequest req, RecordInputStream in1)
        {
            bool going = true;
            short userCode = 0;
            Record r = null;

            // Create a new RecordStream and use that
            HSSFRecordStream recordStream = new HSSFRecordStream(in1);

            // Process each record as they come in
            while (going)
            {
                r = recordStream.NextRecord();
                if (r != null)
                {
                    userCode = req.ProcessRecord(r);
                    if (userCode != 0) break;
                }
                else
                {
                    going = false;
                }
            }

            // All done, return our last code
            return userCode;
        }
    }
}