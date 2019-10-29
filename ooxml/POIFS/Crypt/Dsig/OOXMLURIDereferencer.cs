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

/* ====================================================================
   This product Contains an ASLv2 licensed version of the OOXML signer
   package from the eID Applet project
   http://code.google.com/p/eid-applet/source/browse/tRunk/README.txt  
   Copyright (C) 2008-2014 FedICT.
   ================================================================= */

namespace NPOI.POIFS.Crypt.Dsig
{
    using NPOI.OpenXml4Net.Exceptions;
    using NPOI.OpenXml4Net.OPC;
    using System;
    using System.IO;
    using System.Security.Cryptography.Xml;

    /**
     * JSR105 URI dereferencer for Office Open XML documents.
     */
    public class OOXMLURIDereferencer : IURIDereferencer, ISignatureConfigurable
    {

        //private static POILogger LOG = POILogFactory.GetLogger(typeof(OOXMLURIDereferencer));

        private SignatureConfig signatureConfig;
        private IURIDereferencer baseUriDereferencer;

        public void SetSignatureConfig(SignatureConfig signatureConfig)
        {
            this.signatureConfig = signatureConfig;
        }

        public IData dereference(IURIReference uriReference, SignedXml context)
        {
            if (baseUriDereferencer == null)
            {
                //baseUriDereferencer = signatureConfig.GetSignatureFactory().URIDereferencer;
                throw new NotImplementedException();
            }

            if (null == uriReference)
            {
                throw new NullReferenceException("URIReference cannot be null");
            }
            if (null == context)
            {
                throw new NullReferenceException("XMLCrytoContext cannot be null");
            }

            Uri uri;
            try
            {
                uri = new Uri(uriReference.getURI());
            }
            catch (UriFormatException e)
            {
                throw new Exception("could not URL decode the uri: " + uriReference.getURI(), e);
            }

            PackagePart part = FindPart(uri);
            if (part == null)
            {
                //LOG.Log(POILogger.DEBUG, "cannot Resolve, delegating to base DOM URI dereferencer", uri);
                //return this.baseUriDereferencer.Dereference(uriReference, context);
                throw new NotImplementedException();
            }

            Stream dataStream;
            try
            {
                dataStream = part.GetInputStream();

                // workaround for office 2007 pretty-printed .rels files
                if (part.PartName.ToString().EndsWith(".rels"))
                {
                    // although xmlsec has an option to ignore line breaks, currently this
                    // only affects .rels files, so we only modify these
                    // http://stackoverflow.com/questions/4728300
                    MemoryStream bos = new MemoryStream();
                    //for (int ch; (ch = dataStream.Read()) != -1;)
                    //{
                    //    if (ch == 10 || ch == 13) continue;
                    //    bos.Write(ch);
                    //}
                    dataStream = new MemoryStream(bos.ToArray());
                    throw new NotImplementedException();
                }
            }
            catch (IOException e)
            {
                //throw new URIReferenceException("I/O error: " + e.Message, e);
                throw new NotImplementedException();
            }

            //return new OctetStreamData(dataStream, uri.ToString(), null);
            throw new NotImplementedException();
        }

        private PackagePart FindPart(Uri uri)
        {
            //Console.WriteLine(POILogger.DEBUG, "dereference", uri);

            String path = uri.AbsolutePath;
            if (path == null || "".Equals(path))
            {
                //Console.WriteLine(POILogger.DEBUG, "illegal part name (expected)", uri);
                return null;
            }

            PackagePartName ppn;
            try
            {
                ppn = PackagingUriHelper.CreatePartName(path);
            }
            catch (InvalidFormatException e)
            {
                //Console.WriteLine(POILogger.WARN, "illegal part name (not expected)", uri);
                return null;
            }

            return signatureConfig.GetOpcPackage().GetPart(ppn);
        }
    }

}