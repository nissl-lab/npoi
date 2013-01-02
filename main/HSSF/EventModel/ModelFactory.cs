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

namespace NPOI.HSSF.EventModel
{

    using System;
    using System.IO;
    using System.Collections;
    using System.Text;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;


    /**
     * ModelFactory Creates workbook and sheet models based upon 
     * events thrown by them there events from the EventRecordFactory.
     * 
     * @see org.apache.poi.hssf.eventmodel.EventRecordFactory
     * @author Andrew C. Oliver acoliver@apache.org
     */
    public sealed class ModelFactory : ERFListener
    {

        Model currentmodel;
        bool lastEOF;
        public IList listeners;

        /**
         * Constructor for ModelFactory.  Does practically nothing.
         */
        public ModelFactory(): base()
        {
            currentmodel = null;
            listeners = new ArrayList(1);
        }

        /**
         * register a ModelFactoryListener so that it can receive 
         * Models as they are created.
         */
        public void RegisterListener(ModelFactoryListener listener)
        {
            listeners.Add(listener);
        }

        /**
         * Start Processing the Workbook stream into Model events.
         */
        public void Run(Stream stream)
        {
            EventRecordFactory factory = new EventRecordFactory(this,null);
            lastEOF = true;
            factory.ProcessRecords(stream);
        }

        //ERFListener
        public bool ProcessRecord(Record rec)
        {
            if (rec.Sid == BOFRecord.sid)
            {
                if (lastEOF != true)
                {
                    throw new Exception("Not yet handled embedded models");
                }
                else
                {
                    BOFRecord bof = (BOFRecord)rec;
                    switch (bof.Type)
                    {
                        case BOFRecord.TYPE_WORKBOOK:
                            currentmodel = new InternalWorkbook();
                            break;
                        case BOFRecord.TYPE_WORKSHEET:
                            currentmodel = InternalSheet.CreateSheet();
                            break;
                        default:
                            throw new Exception("Unsupported model type " + bof.GetType());
                    }

                }
            }

            if (rec.Sid == EOFRecord.sid)
            {
                lastEOF = true;
                ThrowEvent(currentmodel);
            }
            else
            {
                lastEOF = false;
            }


            return true;
        }

        
        /**
         * Throws the model as an event to the listeners
         * @param model to be thrown
         */
        private void ThrowEvent(Model model)
        {
            IEnumerator i = listeners.GetEnumerator();
            while (i.MoveNext())
            {
                ModelFactoryListener mfl = (ModelFactoryListener)i.Current;
                mfl.Process(model);
            }
        }


    }
}