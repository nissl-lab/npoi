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

namespace NPOI.POIFS.Crypt.Dsig.Services
{
    using System;
    using System.Collections.Generic;

    /**
     * JSR105 implementation of the RelationshipTransform transformation.
     * 
     * <p>
     * Specs: http://openiso.org/Ecma/376/Part2/12.2.4#26
     * </p>
     */
    public class RelationshipTransformService : TransformService {

        public static String TRANSFORM_URI = "http://schemas.Openxmlformats.org/package/2006/RelationshipTransform";

        private List<String> sourceIds;

        //private static POILogger LOG = POILogFactory.GetLogger(typeof(RelationshipTransformService));

        /**
         * Relationship Transform parameter specification class.
         */
        public class RelationshipTransformParameterSpec : TransformParameterSpec {
            List<String> sourceIds = new List<String>();
            public void AddRelationshipReference(String relationshipId) {
                sourceIds.Add(relationshipId);
            }
            public bool HasSourceIds() {
                return !sourceIds.IsEmpty();
            }
        }


        public RelationshipTransformService() : base()
        {
            //LOG.Log(POILogger.DEBUG, "constructor");
            this.sourceIds = new List<String>();
        }

        /**
         * Register the provider for this TransformService
         * 
         * @see javax.xml.Crypto.dsig.TransformService
         */
        public static void registerDsigProvider() {
            // the xml signature classes will try to find a special TransformerService,
            // which is ofcourse unknown to JCE before ...
            String dsigProvider = "POIXmlDsigProvider";
            if (Security.GetProperty(dsigProvider) == null) {
                //Provider p = new Provider(dsigProvider, 1.0, dsigProvider){
                //    static long serialVersionUID = 1L;
                //};
                //p.Put("TransformService." + TRANSFORM_URI, RelationshipTransformService.class.Name);
                //p.Put("TransformService." + TRANSFORM_URI + " MechanismType", "DOM");
                //Security.AddProvider(p);
                throw new NotImplementedException();
            }
        }



        public void Init(TransformParameterSpec params1) {
            //LOG.Log(POILogger.DEBUG, "Init(params)");
            if (!(params1 is RelationshipTransformParameterSpec)) {
                throw new InvalidAlgorithmParameterException();
            }
            RelationshipTransformParameterSpec relParams = (RelationshipTransformParameterSpec)params1;
            foreach (String sourceId in relParams.sourceIds) {
                this.sourceIds.Add(sourceId);
            }
        }


        public void Init(XMLStructure parent, XMLCryptoContext context) {
            LOG.Log(POILogger.DEBUG, "Init(parent,context)");
            LOG.Log(POILogger.DEBUG, "parent java type: " + parent.Class.Name);
            DOMStructure domParent = (DOMStructure)parent;
            Node parentNode = domParent.Node;

            try {
                TransformDocument transDoc = TransformDocument.Factory.Parse(parentNode);
                XmlObject[] xoList = transDoc.Transform.SelectChildren(RelationshipReferenceDocument.type.DocumentElementName);
                if (xoList.Length == 0) {
                    //LOG.Log(POILogger.WARN, "no RelationshipReference/@SourceId parameters present");
                }
                foreach (XmlObject xo in xoList) {
                    String sourceId = ((CTRelationshipReference)xo).SourceId;
                    LOG.Log(POILogger.DEBUG, "sourceId: ", sourceId);
                    this.sourceIds.Add(sourceId);
                }
            } catch (XmlException e) {
                throw new InvalidAlgorithmParameterException(e);
            }
        }


        public void marshalParams(XMLStructure parent, XMLCryptoContext context) {
            //LOG.Log(POILogger.DEBUG, "marshallParams(parent,context)");
            DOMStructure domParent = (DOMStructure)parent;
            Element parentNode = (Element)domParent.Node;
            // parentNode.AttributeNS=(/*setter*/XML_NS, "xmlns:mdssi", XML_DIGSIG_NS);
            Document doc = parentNode.OwnerDocument;

            foreach (String sourceId in sourceIds) {
                RelationshipReferenceDocument relRef = RelationshipReferenceDocument.Factory.NewInstance();
                relRef.AddNewRelationshipReference().SourceId = (/*setter*/sourceId);
                Node n = relRef.RelationshipReference.DomNode;
                n = doc.ImportNode(n, true);
                parentNode.AppendChild(n);
            }
        }

        public AlgorithmParameterSpec GetParameterSpec() {
            LOG.Log(POILogger.DEBUG, "getParameterSpec");
            return null;
        }

        public Data transform(Data data, XMLCryptoContext context) {
            LOG.Log(POILogger.DEBUG, "transform(data,context)");
            LOG.Log(POILogger.DEBUG, "data java type: " + data.Class.Name);
            OctetStreamData octetStreamData = (OctetStreamData)data;
            LOG.Log(POILogger.DEBUG, "URI: " + octetStreamData.URI);
            InputStream octetStream = octetStreamData.OctetStream;

            RelationshipsDocument relDoc;
            try {
                relDoc = RelationshipsDocument.Factory.Parse(octetStream);
            } catch (Exception e) {
                throw new TransformException(e.Message, e);
            }
            LOG.Log(POILogger.DEBUG, "relationships document", relDoc);

            CTRelationships rels = relDoc.Relationships;
            List<CTRelationship> relList = rels.RelationshipList;
            Iterator<CTRelationship> relIter = rels.RelationshipList.Iterator();
            while (relIter.HasNext()) {
                CTRelationship rel = relIter.Next();
                /*
                 * See: ISO/IEC 29500-2:2008(E) - 13.2.4.24 Relationships Transform
                 * Algorithm.
                 */
                if (!this.sourceIds.Contains(rel.Id)) {
                    LOG.Log(POILogger.DEBUG, "removing element: " + rel.Id);
                    relIter.Remove();
                } else {
                    if (!rel.IsSetTargetMode()) {
                        rel.TargetMode = (/*setter*/STTargetMode.INTERNAL);
                    }
                }
            }

            // TODO: remove non element nodes ???
            LOG.Log(POILogger.DEBUG, "# Relationship elements", relList.Size());

            //XmlSort.Sort(rels, new Comparator<XmlCursor>(){
            //    public int Compare(XmlCursor c1, XmlCursor c2) {
            //        String id1 = ((CTRelationship)c1.Object).Id;
            //        String id2 = ((CTRelationship)c2.Object).Id;
            //        return id1.CompareTo(id2);
            //    }
            //});

            try {
                MemoryStream bos = new MemoryStream();
                XmlOptions xo = new XmlOptions();
                xo.SaveNoXmlDecl;
                relDoc.Save(bos, xo);
                return new OctetStreamData(new MemoryStream(bos.ToByteArray()));
            } catch (IOException e) {
                throw new TransformException(e.Message, e);
            }
        }

        public Data transform(Data data, XMLCryptoContext context, OutputStream os) {
            //LOG.Log(POILogger.DEBUG, "transform(data,context,os)");
            return null;
        }

        public bool IsFeatureSupported(String feature) {
            //LOG.Log(POILogger.DEBUG, "isFeatureSupported(feature)");
            return false;
        }
    }

}