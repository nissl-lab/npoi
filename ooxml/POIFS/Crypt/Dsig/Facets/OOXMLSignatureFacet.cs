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

namespace NPOI.POIFS.Crypt.Dsig.Facets
{
    using NPOI.OpenXml4Net.OPC;
    using System;
    using System.Collections.Generic;
    using System.Security.Cryptography.Xml;
    using System.Xml;


    /**
     * Office OpenXML Signature Facet implementation.
     * 
     * @author fcorneli
     * @see <a href="http://msdn.microsoft.com/en-us/library/cc313071.aspx">[MS-OFFCRYPTO]: Office Document Cryptography Structure</a>
     */
    public class OOXMLSignatureFacet : SignatureFacet
    {


        public override void preSign(
            XmlDocument document
            , List<Reference> references
            , List<XmlNode> objects)
        {
            //LOG.Log(POILogger.DEBUG, "pre sign");
            AddManifestObject(document, references, objects);
            AddSignatureInfo(document, references, objects);
        }

        protected void AddManifestObject(
            XmlDocument document
            , List<Reference> references
            , List<XmlNode> objects)
        {

            List<Reference> manifestReferences = new List<Reference>();
            AddManifestReferences(manifestReferences);
            //Manifest manifest = GetSignatureFactory().newManifest(manifestReferences);

            //String objectId = "idPackageObject"; // really has to be this value.
            //List<XMLStructure> objectContent = new List<XMLStructure>();
            //objectContent.Add(manifest);

            //AddSignatureTime(document, objectContent);

            //XMLObject xo = GetSignatureFactory().newXMLObject(objectContent, objectId, null, null);
            //objects.Add(xo);

            //Reference reference = newReference("#" + objectId, null, XML_DIGSIG_NS + "Object", null, null);
            //references.Add(reference);
            throw new NotImplementedException();
        }

        protected void AddManifestReferences(List<Reference> manifestReferences)
        {

            OPCPackage ooxml = signatureConfig.GetOpcPackage();
            List<PackagePart> relsEntryNames = ooxml.GetPartsByContentType(ContentTypes.RELATIONSHIPS_PART);

            HashSet<String> digestedPartNames = new HashSet<String>();
            //foreach (PackagePart pp in relsEntryNames)
            //{
            //    String baseUri = pp.PartName.Name.ReplaceFirst("(.*)/_rels/.*", "$1");

            //    PackageRelationshipCollection prc;
            //    try
            //    {
            //        prc = new PackageRelationshipCollection(ooxml);
            //        prc.ParseRelationshipsPart(pp);
            //    }
            //    catch (InvalidFormatException e)
            //    {
            //        throw new XMLSignatureException("Invalid relationship descriptor: " + pp.PartName.Name, e);
            //    }

            //    RelationshipTransformParameterSpec parameterSpec = new RelationshipTransformParameterSpec();
            //    foreach (PackageRelationship relationship in prc)
            //    {
            //        String relationshipType = relationship.RelationshipType;

            //        /*
            //         * ECMA-376 Part 2 - 3rd edition
            //         * 13.2.4.16 Manifest Element
            //         * "The producer shall not create a Manifest element that references any data outside of the package."
            //         */
            //        if (TargetMode.EXTERNAL == relationship.TargetMode)
            //        {
            //            continue;
            //        }

            //        if (!isSignedRelationship(relationshipType)) continue;

            //        parameterSpec.AddRelationshipReference(relationship.Id);

            //        // TODO: find a better way ...
            //        String partName = baseUri + relationship.TargetURI.ToString();
                    //if (!partName.startsWith(baseUri))
                    //{
                    //    partName = baseUri + partName;
                    //}
            //        try
            //        {
            //            partName = new URI(partName).normalize().Path.Replace('\\', '/');
            //            LOG.Log(POILogger.DEBUG, "part name: " + partName);
            //        }
            //        catch (URISyntaxException e)
            //        {
            //            throw new XMLSignatureException(e);
            //        }

            //        String contentType;
            //        try
            //        {
            //            PackagePartName relName = PackagingURIHelper.CreatePartName(partName);
            //            PackagePart pp2 = ooxml.GetPart(relName);
            //            contentType = pp2.ContentType;
            //        }
            //        catch (InvalidFormatException e)
            //        {
            //            throw new XMLSignatureException(e);
            //        }

            //        if (relationshipType.EndsWith("customXml")
            //            && !(contentType.Equals("inkml+xml") || contentType.Equals("text/xml")))
            //        {
            //            LOG.Log(POILogger.DEBUG, "skipping customXml with content type: " + contentType);
            //            continue;
            //        }

            //        if (!digestedPartNames.Contains(partName))
            //        {
            //            // We only digest a part once.
            //            String uri = partName + "?ContentType=" + contentType;
            //            Reference reference = newReference(uri, null, null, null, null);
            //            manifestReferences.Add(reference);
            //            digestedPartNames.Add(partName);
            //        }
            //    }

            //    if (parameterSpec.HasSourceIds())
            //    {
            //        List<Transform> transforms = new List<Transform>();
            //        transforms.Add(newTransform(RelationshipTransformService.TRANSFORM_URI, parameterSpec));
            //        transforms.Add(newTransform(CanonicalizationMethod.INCLUSIVE));
            //        String uri = pp.PartName.Name
            //            + "?ContentType=application/vnd.Openxmlformats-package.relationships+xml";
            //        Reference reference = newReference(uri, transforms, null, null, null);
            //        manifestReferences.Add(reference);
            //    }
            //}
            throw new NotImplementedException();
        }


        protected void AddSignatureTime(XmlDocument document, List<XmlNode> objectContent)
        {
            /*
             * SignatureTime
             */
            //DateFormat fmt = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss'Z'");
            //fmt.TimeZone = (/*setter*/TimeZone.GetTimeZone("UTC"));
            //String nowStr = fmt.Format(signatureConfig.ExecutionTime);
            //LOG.Log(POILogger.DEBUG, "now: " + nowStr);

            //SignatureTimeDocument sigTime = SignatureTimeDocument.Factory.NewInstance();
            //CTSignatureTime ctTime = sigTime.AddNewSignatureTime();
            //ctTime.Format = (/*setter*/"YYYY-MM-DDThh:mm:ssTZD");
            //ctTime.Value = (/*setter*/nowStr);

            //Element n = (Element)document.ImportNode(ctTime.DomNode, true);
            //List<XMLStructure> signatureTimeContent = new List<XMLStructure>();
            //signatureTimeContent.Add(new DOMStructure(n));
            //SignatureProperty signatureTimeSignatureProperty = GetSignatureFactory()
            //    .newSignatureProperty(signatureTimeContent, "#" + signatureConfig.PackageSignatureId,
            //    "idSignatureTime");
            //List<SignatureProperty> signaturePropertyContent = new List<SignatureProperty>();
            //signaturePropertyContent.Add(signatureTimeSignatureProperty);
            //SignatureProperties signatureProperties = GetSignatureFactory()
            //    .newSignatureProperties(signaturePropertyContent,
            //    "id-signature-time-" + signatureConfig.ExecutionTime);
            //objectContent.Add(signatureProperties);
            throw new NotImplementedException();
        }

        protected void AddSignatureInfo(XmlDocument document,
            List<Reference> references,
            List<XmlNode> objects)
        {
            //List<XMLStructure> objectContent = new List<XMLStructure>();

            //SignatureInfoV1Document sigV1 = SignatureInfoV1Document.Factory.NewInstance();
            //CTSignatureInfoV1 ctSigV1 = sigV1.AddNewSignatureInfoV1();
            //ctSigV1.ManifestHashAlgorithm = (/*setter*/signatureConfig.DigestMethodUri);
            //Element n = (Element)document.ImportNode(ctSigV1.DomNode, true);
            //n.AttributeNS = (/*setter*/XML_NS, XMLConstants.XMLNS_ATTRIBUTE, MS_DIGSIG_NS);

            //List<XMLStructure> signatureInfoContent = new List<XMLStructure>();
            //signatureInfoContent.Add(new DOMStructure(n));
            //SignatureProperty signatureInfoSignatureProperty = GetSignatureFactory()
            //    .newSignatureProperty(signatureInfoContent, "#" + signatureConfig.PackageSignatureId,
            //    "idOfficeV1Details");

            //List<SignatureProperty> signaturePropertyContent = new List<SignatureProperty>();
            //signaturePropertyContent.Add(signatureInfoSignatureProperty);
            //SignatureProperties signatureProperties = GetSignatureFactory()
            //    .newSignatureProperties(signaturePropertyContent, null);
            //objectContent.Add(signatureProperties);

            //String objectId = "idOfficeObject";
            //objects.Add(getSignatureFactory().newXMLObject(objectContent, objectId, null, null));

            //Reference reference = newReference("#" + objectId, null, XML_DIGSIG_NS + "Object", null, null);
            //references.Add(reference);
            throw new NotImplementedException();
        }

        protected static String GetRelationshipReferenceURI(String zipEntryName)
        {
            return "/"
                + zipEntryName
                + "?ContentType=application/vnd.Openxmlformats-package.relationships+xml";
        }

        protected static String GetResourceReferenceURI(String resourceName, String contentType)
        {
            return "/" + resourceName + "?ContentType=" + contentType;
        }

        protected static bool IsSignedRelationship(String relationshipType)
        {
            //LOG.Log(POILogger.DEBUG, "relationship type: " + relationshipType);
            foreach (String signedTypeExtension in signed)
            {
                if (relationshipType.EndsWith(signedTypeExtension))
                {
                    return true;
                }
            }
            if (relationshipType.EndsWith("customXml"))
            {
                //LOG.Log(POILogger.DEBUG, "customXml relationship type");
                return true;
            }
            return false;
        }

        public static String[] contentTypes = {
        /*
         * Word
         */
        "application/vnd.Openxmlformats-officedocument.wordProcessingml.document.main+xml",
        "application/vnd.Openxmlformats-officedocument.wordProcessingml.fontTable+xml",
        "application/vnd.Openxmlformats-officedocument.wordProcessingml.Settings+xml",
        "application/vnd.Openxmlformats-officedocument.wordProcessingml.styles+xml",
        "application/vnd.Openxmlformats-officedocument.theme+xml",
        "application/vnd.Openxmlformats-officedocument.wordProcessingml.webSettings+xml",
        "application/vnd.Openxmlformats-officedocument.wordProcessingml.numbering+xml",

        /*
         * Word 2010
         */
        "application/vnd.ms-word.stylesWithEffects+xml",

        /*
         * Excel
         */
        "application/vnd.Openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
        "application/vnd.Openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        "application/vnd.Openxmlformats-officedocument.spreadsheetml.styles+xml",
        "application/vnd.Openxmlformats-officedocument.spreadsheetml.sheet.main+xml",

        /*
         * Powerpoint
         */
        "application/vnd.Openxmlformats-officedocument.presentationml.presentation.main+xml",
        "application/vnd.Openxmlformats-officedocument.presentationml.slideLayout+xml",
        "application/vnd.Openxmlformats-officedocument.presentationml.slideMaster+xml",
        "application/vnd.Openxmlformats-officedocument.presentationml.slide+xml",
        "application/vnd.Openxmlformats-officedocument.presentationml.tableStyles+xml",

        /*
         * Powerpoint 2010
         */
        "application/vnd.Openxmlformats-officedocument.presentationml.viewProps+xml",
        "application/vnd.Openxmlformats-officedocument.presentationml.presProps+xml"
    };

        /**
         * Office 2010 list of signed types (extensions).
         */
        public static String[] signed = {
        "powerPivotData", //
        "activeXControlBinary", //
        "attachedToolbars", //
        "connectorXml", //
        "downRev", //
        "functionPrototypes", //
        "graphicFrameDoc", //
        "groupShapeXml", //
        "ink", //
        "keyMapCustomizations", //
        "legacyDiagramText", //
        "legacyDocTextInfo", //
        "officeDocument", //
        "pictureXml", //
        "shapeXml", //
        "smartTags", //
        "ui/altText", //
        "ui/buttonSize", //
        "ui/controlID", //
        "ui/description", //
        "ui/enabled", //
        "ui/extensibility", //
        "ui/helperText", //
        "ui/imageID", //
        "ui/imageMso", //
        "ui/keyTip", //
        "ui/label", //
        "ui/lcid", //
        "ui/loud", //
        "ui/pressed", //
        "ui/progID", //
        "ui/ribbonID", //
        "ui/ShowImage", //
        "ui/ShowLabel", //
        "ui/supertip", //
        "ui/target", //
        "ui/text", //
        "ui/title", //
        "ui/tooltip", //
        "ui/userCustomization", //
        "ui/visible", //
        "userXmlData", //
        "vbaProject", //
        "wordVbaData", //
        "wsSortMap", //
        "xlBinaryIndex", //
        "xlExternalLinkPath/xlAlternateStartup", //
        "xlExternalLinkPath/xlLibrary", //
        "xlExternalLinkPath/xlPathMissing", //
        "xlExternalLinkPath/xlStartup", //
        "xlIntlMacrosheet", //
        "xlMacrosheet", //
        "customData", //
        "diagramDrawing", //
        "hdphoto", //
        "inkXml", //
        "media", //
        "slicer", //
        "slicerCache", //
        "stylesWithEffects", //
        "ui/extensibility", //
        "chartColorStyle", //
        "chartLayout", //
        "chartStyle", //
        "dictionary", //
        "timeline", //
        "timelineCache", //
        "aFChunk", //
        "attachedTemplate", //
        "audio", //
        "calcChain", //
        "chart", //
        "chartsheet", //
        "chartUserShapes", //
        "commentAuthors", //
        "comments", //
        "connections", //
        "control", //
        "customProperty", //
        "customXml", //
        "diagramColors", //
        "diagramData", //
        "diagramLayout", //
        "diagramQuickStyle", //
        "dialogsheet", //
        "drawing", //
        "endnotes", //
        "externalLink", //
        "externalLinkPath", //
        "font", //
        "fontTable", //
        "footer", //
        "footnotes", //
        "glossaryDocument", //
        "handoutMaster", //
        "header", //
        "hyperlink", //
        "image", //
        "mailMergeHeaderSource", //
        "mailMergeRecipientData", //
        "mailMergeSource", //
        "notesMaster", //
        "notesSlide", //
        "numbering", //
        "officeDocument", //
        "oleObject", //
        "package", //
        "pivotCacheDefInition", //
        "pivotCacheRecords", //
        "pivotTable", //
        "presProps", //
        "printerSettings", //
        "queryTable", //
        "recipientData", //
        "settings", //
        "sharedStrings", //
        "sheetMetadata", //
        "slide", //
        "slideLayout", //
        "slideMaster", //
        "slideUpdateInfo", //
        "slideUpdateUrl", //
        "styles", //
        "table", //
        "tableSingleCells", //
        "tableStyles", //
        "tags", //
        "theme", //
        "themeOverride", //
        "transform", //
        "video", //
        "viewProps", //
        "volatileDependencies", //
        "webSettings", //
        "worksheet", //
        "xmlMaps", //
        "ctrlProp", //
        "customData", //
        "diagram", //
        "diagramColorsHeader", //
        "diagramLayoutHeader", //
        "diagramQuickStyleHeader", //
        "documentParts", //
        "slicer", //
        "slicerCache", //
        "vmlDrawing" //
    };
    }
}