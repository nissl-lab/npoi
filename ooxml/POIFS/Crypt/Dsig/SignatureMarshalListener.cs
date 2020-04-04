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

namespace NPOI.POIFS.Crypt.Dsig
{
    using System;
    using System.Threading;
    using System.Xml;


    /**
     * This listener class is used, to modify the to be digested xml document,
     * e.g. to register id attributes or Set prefixes for registered namespaces
     */
    public class SignatureMarshalListener : IEventListener, ISignatureConfigurable
    {
        ///ThreadLocal<EventTarget> target = new ThreadLocal<EventTarget>();
        SignatureConfig signatureConfig;
        ////public void SetEventTarget(EventTarget target)
        ////{
        ////    this.target.Set(target);
        ////}

        public void handleEvent(IEvent e)
        {
            ////if (!(e is MutationEvent)) return;
            ////MutationEvent mutEvt = (MutationEvent)e;
            ////EventTarget et = mutEvt.Target;
            ////if (!(et is XmlElement)) return;
            ////handleElement((XmlElement)et);
            throw new NotImplementedException();
        }

        public void handleElement(XmlElement el)
        {
            //EventTarget target = this.target.Get();
            //String packageId = signatureConfig.PackageSignatureId;
            //if (el.HasAttribute("Id"))
            //{
            //    el.IdAttribute = (/*setter*/"Id", true);
            //}

            //SetListener(target, this, false);
            //if (packageId.Equals(el.GetAttribute("Id")))
            //{
            //    el.AttributeNS = (/*setter*/XML_NS, "xmlns:mdssi", OO_DIGSIG_NS);
            //}
            //SetPrefix(el);
            //SetListener(target, this, true);
            throw new NotImplementedException();
        }

        // helper method to keep it in one place
        ////public static void SetListener(EventTarget target, EventListener listener, bool enabled)
        ////{
        ////    String type = "DOMSubtreeModified";
        ////    bool useCapture = false;
        ////    if (enabled)
        ////    {
        ////        target.AddEventListener(type, listener, useCapture);
        ////    }
        ////    else
        ////    {
        ////        target.RemoveEventListener(type, listener, useCapture);
        ////    }
        ////}

        protected void SetPrefix(XmlNode el)
        {
            String prefix = signatureConfig.GetNamespacePrefixes()[(el.NamespaceURI)];
            if (prefix != null && el.Prefix == null)
            {
                el.Prefix = (/*setter*/prefix);
            }

            XmlNodeList nl = el.ChildNodes;
            for (int i = 0; i < nl.Count; i++)
            {
                SetPrefix(nl.Item(i));
            }
        }

        public void SetSignatureConfig(SignatureConfig signatureConfig)
        {
            this.signatureConfig = signatureConfig;
        }
    }
}