
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
    using System;
    using System.Collections;
    using NPOI.HSSF.Record;


    /// <summary>
    /// An HSSFRequest object should be constructed registering an instance or multiple
    /// instances of HSSFListener with each Record.sid you wish to listen for.
    /// @author  Andrew C. Oliver (acoliver at apache dot org)
    /// @author Carey Sublette (careysub@earthling.net)
    /// </summary>
    public class HSSFRequest
    {
        private Hashtable records;

        /// <summary>
        /// Creates a new instance of HSSFRequest
        /// </summary>
        public HSSFRequest()
        {
            records =
                new Hashtable(50);   // most folks won't listen for too many of these
        }

        /// <summary>
        /// Add an event listener for a particular record type.  The trick Is you have to know
        /// what the records are for or just start with our examples and build on them.  Alternatively,
        /// you CAN call AddListenerForAllRecords and you'll recieve ALL record events in one listener,
        /// but if you like to squeeze every last byte of efficiency out of life you my not like this.
        /// (its sure as heck what I plan to do)
        /// </summary>
        /// <param name="lsnr">for the event</param>
        /// <param name="sid">identifier for the record type this Is the .sid static member on the individual records</param>
        public void AddListener(IHSSFListener lsnr, short sid)
        {
            IList list = null;
            Object obj = records[sid];

            if (obj != null)
            {
                list = (IList)obj;
            }
            else
            {
                list = new ArrayList(
                    1);   // probably most people will use one listener
                list.Add(lsnr);
                records[sid]=list;
            }
        }

        /// <summary>
        /// This Is the equivilent of calling AddListener(myListener, sid) for EVERY
        /// record in the org.apache.poi.hssf.record package. This Is for lazy
        /// people like me. You can call this more than once with more than one listener, but
        /// that seems like a bad thing to do from a practice-perspective Unless you have a
        /// compelling reason to do so (like maybe you send the event two places or log it or
        /// something?).
        /// </summary>
        /// <param name="lsnr">a single listener to associate with ALL records</param>
        public void AddListenerForAllRecords(IHSSFListener lsnr)
        {
            short[] rectypes = RecordFactory.GetAllKnownRecordSIDs();

            for (int k = 0; k < rectypes.Length; k++)
            {
                AddListener(lsnr, rectypes[k]);
            }
        }

        /// <summary>
        /// Called by HSSFEventFactory, passes the Record to each listener associated with
        /// a record.sid.
        /// Exception and return value Added 2002-04-19 by Carey Sublette
        /// </summary>
        /// <param name="rec">The record.</param>
        /// <returns>numeric user-specified result code. If zero continue Processing.</returns>
        public short ProcessRecord(Record rec)
        {
            Object obj = records[rec.Sid];
            short userCode = 0;

            if (obj != null)
            {
                IList listeners = (IList)obj;

                for (int k = 0; k < listeners.Count; k++)
                {
                    Object listenObj = listeners[k];
                    if (listenObj is AbortableHSSFListener)
                    {
                        AbortableHSSFListener listener = (AbortableHSSFListener)listenObj;
                        userCode = listener.AbortableProcessRecord(rec);
                        if (userCode != 0) break;
                    }
                    else
                    {
                        IHSSFListener listener = (IHSSFListener)listenObj;
                        listener.ProcessRecord(rec);
                    }
                }
            }
            return userCode;
        }
    }
}