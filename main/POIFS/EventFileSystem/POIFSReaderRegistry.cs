
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


using System;
using System.Collections;

using NPOI.POIFS.FileSystem;

namespace NPOI.POIFS.EventFileSystem
{
    /**
     * A registry for POIFSReaderListeners and the DocumentDescriptors of
     * the documents those listeners are interested in
     *
     * @author Marc Johnson (mjohnson at apache dot org)
     * @version %I%, %G%
     */

    public class POIFSReaderRegistry
    {

        // the POIFSReaderListeners who listen to all POIFSReaderEvents
        private ArrayList omnivorousListeners;

        // Each mapping in this Map has a key consisting of a
        // POIFSReaderListener and a value cosisting of a Set of
        // DocumentDescriptors for the documents that POIFSReaderListener
        // is interested in; used to efficiently manage the registry
        private Hashtable selectiveListeners;

        // Each mapping in this Map has a key consisting of a
        // DocumentDescriptor and a value consisting of a Set of
        // POIFSReaderListeners for the document matching that
        // DocumentDescriptor; used when a document is found, to quickly
        // Get the listeners interested in that document
        private Hashtable chosenDocumentDescriptors;

        /**
         * Construct the registry
         */

        public POIFSReaderRegistry()
        {
            omnivorousListeners = new ArrayList();
            selectiveListeners = new Hashtable();
            chosenDocumentDescriptors = new Hashtable();
        }

        /**
         * Register a POIFSReaderListener for a particular document
         *
         * @param listener the listener
         * @param path the path of the document of interest
         * @param documentName the name of the document of interest
         */

        public void RegisterListener(POIFSReaderListener listener,
                              POIFSDocumentPath path,
                              String documentName)
        {
            if (!omnivorousListeners.Contains(listener))
            {

                // not an omnivorous listener (if it was, this method is a
                // no-op)
                ArrayList descriptors = (ArrayList)selectiveListeners[listener];

                if (descriptors == null)
                {

                    // this listener has not Registered before
                    descriptors = new ArrayList();
                    selectiveListeners[listener] = descriptors;
                }
                DocumentDescriptor descriptor = new DocumentDescriptor(path,
                                                    documentName);

                if (descriptors.Add(descriptor) >= 0)
                {

                    // this listener wasn't alReady listening for this
                    // document -- Add the listener to the Set of
                    // listeners for this document
                    ArrayList listeners =
                        (ArrayList)chosenDocumentDescriptors[descriptor];

                    if (listeners == null)
                    {

                        // nobody was listening for this document before
                        listeners = new ArrayList();
                        chosenDocumentDescriptors[descriptor] = listeners;
                    }
                    listeners.Add(listener);
                }
            }
        }

        /**
         * Register for all documents
         *
         * @param listener the listener who wants to Get all documents
         */

        public void RegisterListener(POIFSReaderListener listener)
        {
            if (!omnivorousListeners.Contains(listener))
            {

                // wasn't alReady listening for everything, so drop
                // anything listener might have been listening for and
                // then Add the listener to the Set of omnivorous
                // listeners
                RemoveSelectiveListener(listener);
                omnivorousListeners.Add(listener);
            }
        }

        /**
         * Get am iterator of listeners for a particular document
         *
         * @param path the document path
         * @param name the name of the document
         *
         * @return an Iterator POIFSReaderListeners; may be empty
         */

        public IEnumerator GetListeners(POIFSDocumentPath path, String name)
        {
            ArrayList rval = new ArrayList(omnivorousListeners);
            ArrayList selectiveListeners =
                (ArrayList)chosenDocumentDescriptors[new DocumentDescriptor(path,
                    name)];

            if (selectiveListeners != null)
            {
                rval.AddRange(selectiveListeners);
            }
            return rval.GetEnumerator();
        }

        private void RemoveSelectiveListener(POIFSReaderListener listener)
        {
            ArrayList selectedDescriptors = (ArrayList)selectiveListeners[listener];

            if (selectedDescriptors != null)
            {
                selectiveListeners.Remove(listener);
                IEnumerator iter = selectedDescriptors.GetEnumerator();

                while (iter.MoveNext())
                {
                    DropDocument(listener, (DocumentDescriptor)iter.Current);
                }
            }
        }

        private void DropDocument(POIFSReaderListener listener,
                                  DocumentDescriptor descriptor)
        {
            ArrayList listeners = (ArrayList)chosenDocumentDescriptors[descriptor];

            listeners.Remove(listener);
            if (listeners.Count == 0)
            {
                chosenDocumentDescriptors.Remove(descriptor);
            }
        }
    }   // end package scope class POIFSReaderRegistry

}