using NPOI.OpenXml4Net.Util;
using NPOI.OpenXmlFormats.Shared;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
//using NPOI.OpenXmlFormats.Shared;

namespace NPOI.OpenXmlFormats.Wordprocessing
{


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot("settings", Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = false)]
    public class CT_Settings
    {
        public static CT_Settings Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Settings ctObj = new CT_Settings();
            ctObj.activeWritingStyle = new List<CT_WritingStyle>();
            ctObj.docVars = new List<CT_DocVar>();
            ctObj.attachedSchema = new List<CT_String>();
            ctObj.smartTagType = new List<CT_SmartTagType>();
            ctObj.schemaLibrary = new List<CT_Schema>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "writeProtection")
                    ctObj.writeProtection = CT_WriteProtection.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "view")
                    ctObj.view = CT_View.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "zoom")
                    ctObj.zoom = CT_Zoom.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "removePersonalInformation")
                    ctObj.removePersonalInformation = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "removeDateAndTime")
                    ctObj.removeDateAndTime = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotDisplayPageBoundaries")
                    ctObj.doNotDisplayPageBoundaries = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "displayBackgroundShape")
                    ctObj.displayBackgroundShape = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "printPostScriptOverText")
                    ctObj.printPostScriptOverText = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "printFractionalCharacterWidth")
                    ctObj.printFractionalCharacterWidth = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "printFormsData")
                    ctObj.printFormsData = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "embedTrueTypeFonts")
                    ctObj.embedTrueTypeFonts = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "embedSystemFonts")
                    ctObj.embedSystemFonts = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "saveSubsetFonts")
                    ctObj.saveSubsetFonts = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "saveFormsData")
                    ctObj.saveFormsData = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "mirrorMargins")
                    ctObj.mirrorMargins = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "alignBordersAndEdges")
                    ctObj.alignBordersAndEdges = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bordersDoNotSurroundHeader")
                    ctObj.bordersDoNotSurroundHeader = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bordersDoNotSurroundFooter")
                    ctObj.bordersDoNotSurroundFooter = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "gutterAtTop")
                    ctObj.gutterAtTop = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "hideSpellingErrors")
                    ctObj.hideSpellingErrors = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "hideGrammaticalErrors")
                    ctObj.hideGrammaticalErrors = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "proofState")
                    ctObj.proofState = CT_Proof.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "formsDesign")
                    ctObj.formsDesign = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "attachedTemplate")
                    ctObj.attachedTemplate = CT_Rel.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "linkStyles")
                    ctObj.linkStyles = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "stylePaneFormatFilter")
                    ctObj.stylePaneFormatFilter = CT_ShortHexNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "stylePaneSortMethod")
                    ctObj.stylePaneSortMethod = CT_ShortHexNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "documentType")
                    ctObj.documentType = CT_DocType.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "mailMerge")
                    ctObj.mailMerge = CT_MailMerge.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "revisionView")
                    ctObj.revisionView = CT_TrackChangesView.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "trackRevisions")
                    ctObj.trackRevisions = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotTrackMoves")
                    ctObj.doNotTrackMoves = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotTrackFormatting")
                    ctObj.doNotTrackFormatting = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "documentProtection")
                    ctObj.documentProtection = CT_DocProtect.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "autoFormatOverride")
                    ctObj.autoFormatOverride = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "styleLockTheme")
                    ctObj.styleLockTheme = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "styleLockQFSet")
                    ctObj.styleLockQFSet = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "defaultTabStop")
                    ctObj.defaultTabStop = CT_TwipsMeasure.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "autoHyphenation")
                    ctObj.autoHyphenation = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "consecutiveHyphenLimit")
                    ctObj.consecutiveHyphenLimit = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "hyphenationZone")
                    ctObj.hyphenationZone = CT_TwipsMeasure.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotHyphenateCaps")
                    ctObj.doNotHyphenateCaps = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "showEnvelope")
                    ctObj.showEnvelope = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "summaryLength")
                    ctObj.summaryLength = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "clickAndTypeStyle")
                    ctObj.clickAndTypeStyle = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "defaultTableStyle")
                    ctObj.defaultTableStyle = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "evenAndOddHeaders")
                    ctObj.evenAndOddHeaders = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bookFoldRevPrinting")
                    ctObj.bookFoldRevPrinting = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bookFoldPrinting")
                    ctObj.bookFoldPrinting = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bookFoldPrintingSheets")
                    ctObj.bookFoldPrintingSheets = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "drawingGridHorizontalSpacing")
                    ctObj.drawingGridHorizontalSpacing = CT_TwipsMeasure.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "drawingGridVerticalSpacing")
                    ctObj.drawingGridVerticalSpacing = CT_TwipsMeasure.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "displayHorizontalDrawingGridEvery")
                    ctObj.displayHorizontalDrawingGridEvery = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "displayVerticalDrawingGridEvery")
                    ctObj.displayVerticalDrawingGridEvery = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotUseMarginsForDrawingGridOrigin")
                    ctObj.doNotUseMarginsForDrawingGridOrigin = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "drawingGridHorizontalOrigin")
                    ctObj.drawingGridHorizontalOrigin = CT_TwipsMeasure.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "drawingGridVerticalOrigin")
                    ctObj.drawingGridVerticalOrigin = CT_TwipsMeasure.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotShadeFormData")
                    ctObj.doNotShadeFormData = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "noPunctuationKerning")
                    ctObj.noPunctuationKerning = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "characterSpacingControl")
                    ctObj.characterSpacingControl = CT_CharacterSpacing.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "printTwoOnOne")
                    ctObj.printTwoOnOne = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "strictFirstAndLastChars")
                    ctObj.strictFirstAndLastChars = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "noLineBreaksAfter")
                    ctObj.noLineBreaksAfter = CT_Kinsoku.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "noLineBreaksBefore")
                    ctObj.noLineBreaksBefore = CT_Kinsoku.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "savePreviewPicture")
                    ctObj.savePreviewPicture = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotValidateAgainstSchema")
                    ctObj.doNotValidateAgainstSchema = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "saveInvalidXml")
                    ctObj.saveInvalidXml = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "ignoreMixedContent")
                    ctObj.ignoreMixedContent = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "alwaysShowPlaceholderText")
                    ctObj.alwaysShowPlaceholderText = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotDemarcateInvalidXml")
                    ctObj.doNotDemarcateInvalidXml = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "saveXmlDataOnly")
                    ctObj.saveXmlDataOnly = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "useXSLTWhenSaving")
                    ctObj.useXSLTWhenSaving = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "saveThroughXslt")
                    ctObj.saveThroughXslt = CT_SaveThroughXslt.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "showXMLTags")
                    ctObj.showXMLTags = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "alwaysMergeEmptyNamespace")
                    ctObj.alwaysMergeEmptyNamespace = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "updateFields")
                    ctObj.updateFields = CT_OnOff.Parse(childNode, namespaceManager);
                //else if(childNode.LocalName == "hdrShapeDefaults")
                //    ctObj.hdrShapeDefaults = XmlElement[].Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "footnotePr")
                    ctObj.footnotePr = CT_FtnDocProps.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "endnotePr")
                    ctObj.endnotePr = CT_EdnDocProps.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "compat")
                    ctObj.compat = CT_Compat.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "rsids")
                    ctObj.rsids = CT_DocRsids.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "mathPr")
                    ctObj.mathPr = CT_MathPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "uiCompat97To2003")
                    ctObj.uiCompat97To2003 = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "themeFontLang")
                    ctObj.themeFontLang = CT_Language.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "clrSchemeMapping")
                    ctObj.clrSchemeMapping = CT_ColorSchemeMapping.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotIncludeSubdocsInStats")
                    ctObj.doNotIncludeSubdocsInStats = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotAutoCompressPictures")
                    ctObj.doNotAutoCompressPictures = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "forceUpgrade")
                    ctObj.forceUpgrade = new CT_Empty();
                else if (childNode.LocalName == "captions")
                    ctObj.captions = CT_Captions.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "readModeInkLockDown")
                    ctObj.readModeInkLockDown = CT_ReadingModeInkLockDown.Parse(childNode, namespaceManager);
                //else if(childNode.LocalName == "shapeDefaults")
                //    ctObj.shapeDefaults = XmlElement[].Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotEmbedSmartTags")
                    ctObj.doNotEmbedSmartTags = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "decimalSymbol")
                    ctObj.decimalSymbol = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "listSeparator")
                    ctObj.listSeparator = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "activeWritingStyle")
                    ctObj.activeWritingStyle.Add(CT_WritingStyle.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "docVars")
                    ctObj.docVars.Add(CT_DocVar.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "attachedSchema")
                    ctObj.attachedSchema.Add(CT_String.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "smartTagType")
                    ctObj.smartTagType.Add(CT_SmartTagType.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "schemaLibrary")
                    ctObj.schemaLibrary.Add(CT_Schema.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw)
        {
           sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write("<w:settings xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" ");
            sw.Write("xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:sl=\"http://schemas.openxmlformats.org/schemaLibrary/2006/main\">");
            if (this.writeProtection != null)
                this.writeProtection.Write(sw, "writeProtection");
            if (this.view != null)
                this.view.Write(sw, "view");
            if (this.zoom != null)
                this.zoom.Write(sw, "zoom");
            if (this.removePersonalInformation != null)
                this.removePersonalInformation.Write(sw, "removePersonalInformation");
            if (this.removeDateAndTime != null)
                this.removeDateAndTime.Write(sw, "removeDateAndTime");
            if (this.doNotDisplayPageBoundaries != null)
                this.doNotDisplayPageBoundaries.Write(sw, "doNotDisplayPageBoundaries");
            if (this.displayBackgroundShape != null)
                this.displayBackgroundShape.Write(sw, "displayBackgroundShape");
            if (this.printPostScriptOverText != null)
                this.printPostScriptOverText.Write(sw, "printPostScriptOverText");
            if (this.printFractionalCharacterWidth != null)
                this.printFractionalCharacterWidth.Write(sw, "printFractionalCharacterWidth");
            if (this.printFormsData != null)
                this.printFormsData.Write(sw, "printFormsData");
            if (this.embedTrueTypeFonts != null)
                this.embedTrueTypeFonts.Write(sw, "embedTrueTypeFonts");
            if (this.embedSystemFonts != null)
                this.embedSystemFonts.Write(sw, "embedSystemFonts");
            if (this.saveSubsetFonts != null)
                this.saveSubsetFonts.Write(sw, "saveSubsetFonts");
            if (this.saveFormsData != null)
                this.saveFormsData.Write(sw, "saveFormsData");
            if (this.mirrorMargins != null)
                this.mirrorMargins.Write(sw, "mirrorMargins");
            if (this.alignBordersAndEdges != null)
                this.alignBordersAndEdges.Write(sw, "alignBordersAndEdges");
            if (this.bordersDoNotSurroundHeader != null)
                this.bordersDoNotSurroundHeader.Write(sw, "bordersDoNotSurroundHeader");
            if (this.bordersDoNotSurroundFooter != null)
                this.bordersDoNotSurroundFooter.Write(sw, "bordersDoNotSurroundFooter");
            if (this.gutterAtTop != null)
                this.gutterAtTop.Write(sw, "gutterAtTop");
            if (this.hideSpellingErrors != null)
                this.hideSpellingErrors.Write(sw, "hideSpellingErrors");
            if (this.hideGrammaticalErrors != null)
                this.hideGrammaticalErrors.Write(sw, "hideGrammaticalErrors");
            if (this.proofState != null)
                this.proofState.Write(sw, "proofState");
            if (this.formsDesign != null)
                this.formsDesign.Write(sw, "formsDesign");
            if (this.attachedTemplate != null)
                this.attachedTemplate.Write(sw, "attachedTemplate");
            if (this.linkStyles != null)
                this.linkStyles.Write(sw, "linkStyles");
            if (this.stylePaneFormatFilter != null)
                this.stylePaneFormatFilter.Write(sw, "stylePaneFormatFilter");
            if (this.stylePaneSortMethod != null)
                this.stylePaneSortMethod.Write(sw, "stylePaneSortMethod");
            if (this.documentType != null)
                this.documentType.Write(sw, "documentType");
            if (this.mailMerge != null)
                this.mailMerge.Write(sw, "mailMerge");
            if (this.revisionView != null)
                this.revisionView.Write(sw, "revisionView");
            if (this.trackRevisions != null)
                this.trackRevisions.Write(sw, "trackRevisions");
            if (this.doNotTrackMoves != null)
                this.doNotTrackMoves.Write(sw, "doNotTrackMoves");
            if (this.doNotTrackFormatting != null)
                this.doNotTrackFormatting.Write(sw, "doNotTrackFormatting");
            if (this.documentProtection != null)
                this.documentProtection.Write(sw, "documentProtection");
            if (this.autoFormatOverride != null)
                this.autoFormatOverride.Write(sw, "autoFormatOverride");
            if (this.styleLockTheme != null)
                this.styleLockTheme.Write(sw, "styleLockTheme");
            if (this.styleLockQFSet != null)
                this.styleLockQFSet.Write(sw, "styleLockQFSet");
            if (this.defaultTabStop != null)
                this.defaultTabStop.Write(sw, "defaultTabStop");
            if (this.autoHyphenation != null)
                this.autoHyphenation.Write(sw, "autoHyphenation");
            if (this.consecutiveHyphenLimit != null)
                this.consecutiveHyphenLimit.Write(sw, "consecutiveHyphenLimit");
            if (this.hyphenationZone != null)
                this.hyphenationZone.Write(sw, "hyphenationZone");
            if (this.doNotHyphenateCaps != null)
                this.doNotHyphenateCaps.Write(sw, "doNotHyphenateCaps");
            if (this.showEnvelope != null)
                this.showEnvelope.Write(sw, "showEnvelope");
            if (this.summaryLength != null)
                this.summaryLength.Write(sw, "summaryLength");
            if (this.clickAndTypeStyle != null)
                this.clickAndTypeStyle.Write(sw, "clickAndTypeStyle");
            if (this.defaultTableStyle != null)
                this.defaultTableStyle.Write(sw, "defaultTableStyle");
            if (this.evenAndOddHeaders != null)
                this.evenAndOddHeaders.Write(sw, "evenAndOddHeaders");
            if (this.bookFoldRevPrinting != null)
                this.bookFoldRevPrinting.Write(sw, "bookFoldRevPrinting");
            if (this.bookFoldPrinting != null)
                this.bookFoldPrinting.Write(sw, "bookFoldPrinting");
            if (this.bookFoldPrintingSheets != null)
                this.bookFoldPrintingSheets.Write(sw, "bookFoldPrintingSheets");
            if (this.drawingGridHorizontalSpacing != null)
                this.drawingGridHorizontalSpacing.Write(sw, "drawingGridHorizontalSpacing");
            if (this.drawingGridVerticalSpacing != null)
                this.drawingGridVerticalSpacing.Write(sw, "drawingGridVerticalSpacing");
            if (this.displayHorizontalDrawingGridEvery != null)
                this.displayHorizontalDrawingGridEvery.Write(sw, "displayHorizontalDrawingGridEvery");
            if (this.displayVerticalDrawingGridEvery != null)
                this.displayVerticalDrawingGridEvery.Write(sw, "displayVerticalDrawingGridEvery");
            if (this.doNotUseMarginsForDrawingGridOrigin != null)
                this.doNotUseMarginsForDrawingGridOrigin.Write(sw, "doNotUseMarginsForDrawingGridOrigin");
            if (this.drawingGridHorizontalOrigin != null)
                this.drawingGridHorizontalOrigin.Write(sw, "drawingGridHorizontalOrigin");
            if (this.drawingGridVerticalOrigin != null)
                this.drawingGridVerticalOrigin.Write(sw, "drawingGridVerticalOrigin");
            if (this.doNotShadeFormData != null)
                this.doNotShadeFormData.Write(sw, "doNotShadeFormData");
            if (this.noPunctuationKerning != null)
                this.noPunctuationKerning.Write(sw, "noPunctuationKerning");
            if (this.characterSpacingControl != null)
                this.characterSpacingControl.Write(sw, "characterSpacingControl");
            if (this.printTwoOnOne != null)
                this.printTwoOnOne.Write(sw, "printTwoOnOne");
            if (this.strictFirstAndLastChars != null)
                this.strictFirstAndLastChars.Write(sw, "strictFirstAndLastChars");
            if (this.noLineBreaksAfter != null)
                this.noLineBreaksAfter.Write(sw, "noLineBreaksAfter");
            if (this.noLineBreaksBefore != null)
                this.noLineBreaksBefore.Write(sw, "noLineBreaksBefore");
            if (this.savePreviewPicture != null)
                this.savePreviewPicture.Write(sw, "savePreviewPicture");
            if (this.doNotValidateAgainstSchema != null)
                this.doNotValidateAgainstSchema.Write(sw, "doNotValidateAgainstSchema");
            if (this.saveInvalidXml != null)
                this.saveInvalidXml.Write(sw, "saveInvalidXml");
            if (this.ignoreMixedContent != null)
                this.ignoreMixedContent.Write(sw, "ignoreMixedContent");
            if (this.alwaysShowPlaceholderText != null)
                this.alwaysShowPlaceholderText.Write(sw, "alwaysShowPlaceholderText");
            if (this.doNotDemarcateInvalidXml != null)
                this.doNotDemarcateInvalidXml.Write(sw, "doNotDemarcateInvalidXml");
            if (this.saveXmlDataOnly != null)
                this.saveXmlDataOnly.Write(sw, "saveXmlDataOnly");
            if (this.useXSLTWhenSaving != null)
                this.useXSLTWhenSaving.Write(sw, "useXSLTWhenSaving");
            if (this.saveThroughXslt != null)
                this.saveThroughXslt.Write(sw, "saveThroughXslt");
            if (this.showXMLTags != null)
                this.showXMLTags.Write(sw, "showXMLTags");
            if (this.alwaysMergeEmptyNamespace != null)
                this.alwaysMergeEmptyNamespace.Write(sw, "alwaysMergeEmptyNamespace");
            if (this.updateFields != null)
                this.updateFields.Write(sw, "updateFields");
            //if (this.hdrShapeDefaults != null)
            //    this.hdrShapeDefaults.Write(sw, "hdrShapeDefaults");
            if (this.footnotePr != null)
                this.footnotePr.Write(sw, "footnotePr");
            if (this.endnotePr != null)
                this.endnotePr.Write(sw, "endnotePr");
            if (this.compat != null)
                this.compat.Write(sw, "compat");
            if (this.rsids != null)
                this.rsids.Write(sw, "rsids");
            if (this.mathPr != null)
                this.mathPr.Write(sw, "mathPr");
            if (this.uiCompat97To2003 != null)
                this.uiCompat97To2003.Write(sw, "uiCompat97To2003");
            if (this.themeFontLang != null)
                this.themeFontLang.Write(sw, "themeFontLang");
            if (this.clrSchemeMapping != null)
                this.clrSchemeMapping.Write(sw, "clrSchemeMapping");
            if (this.doNotIncludeSubdocsInStats != null)
                this.doNotIncludeSubdocsInStats.Write(sw, "doNotIncludeSubdocsInStats");
            if (this.doNotAutoCompressPictures != null)
                this.doNotAutoCompressPictures.Write(sw, "doNotAutoCompressPictures");
            if (this.forceUpgrade != null)
                sw.Write("<w:forceUpgrade/>");
            if (this.captions != null)
                this.captions.Write(sw, "captions");
            if (this.readModeInkLockDown != null)
                this.readModeInkLockDown.Write(sw, "readModeInkLockDown");
            //if (this.shapeDefaults != null)
            //    this.shapeDefaults.Write(sw, "shapeDefaults");
            if (this.doNotEmbedSmartTags != null)
                this.doNotEmbedSmartTags.Write(sw, "doNotEmbedSmartTags");
            if (this.decimalSymbol != null)
                this.decimalSymbol.Write(sw, "decimalSymbol");
            if (this.listSeparator != null)
                this.listSeparator.Write(sw, "listSeparator");
            if (this.activeWritingStyle != null)
            {
                foreach (CT_WritingStyle x in this.activeWritingStyle)
                {
                    x.Write(sw, "activeWritingStyle");
                }
            }
            if (this.docVars != null)
            {
                foreach (CT_DocVar x in this.docVars)
                {
                    x.Write(sw, "docVars");
                }
            }
            if (this.attachedSchema != null)
            {
                foreach (CT_String x in this.attachedSchema)
                {
                    x.Write(sw, "attachedSchema");
                }
            }
            if (this.smartTagType != null)
            {
                foreach (CT_SmartTagType x in this.smartTagType)
                {
                    x.Write(sw, "smartTagType");
                }
            }
            if (this.schemaLibrary != null)
            {
                foreach (CT_Schema x in this.schemaLibrary)
                {
                    x.Write(sw, "schemaLibrary");
                }
            }
            sw.Write("</w:settings>");
        }

        private CT_WriteProtection writeProtectionField;

        private CT_View viewField;

        private CT_Zoom zoomField;

        private CT_OnOff removePersonalInformationField;

        private CT_OnOff removeDateAndTimeField;

        private CT_OnOff doNotDisplayPageBoundariesField;

        private CT_OnOff displayBackgroundShapeField;

        private CT_OnOff printPostScriptOverTextField;

        private CT_OnOff printFractionalCharacterWidthField;

        private CT_OnOff printFormsDataField;

        private CT_OnOff embedTrueTypeFontsField;

        private CT_OnOff embedSystemFontsField;

        private CT_OnOff saveSubsetFontsField;

        private CT_OnOff saveFormsDataField;

        private CT_OnOff mirrorMarginsField;

        private CT_OnOff alignBordersAndEdgesField;

        private CT_OnOff bordersDoNotSurroundHeaderField;

        private CT_OnOff bordersDoNotSurroundFooterField;

        private CT_OnOff gutterAtTopField;

        private CT_OnOff hideSpellingErrorsField;

        private CT_OnOff hideGrammaticalErrorsField;

        private List<CT_WritingStyle> activeWritingStyleField;

        private CT_Proof proofStateField;

        private CT_OnOff formsDesignField;

        private CT_Rel attachedTemplateField;

        private CT_OnOff linkStylesField;

        private CT_ShortHexNumber stylePaneFormatFilterField;

        private CT_ShortHexNumber stylePaneSortMethodField;

        private CT_DocType documentTypeField;

        private CT_MailMerge mailMergeField;

        private CT_TrackChangesView revisionViewField;

        private CT_OnOff trackRevisionsField;

        private CT_OnOff doNotTrackMovesField;

        private CT_OnOff doNotTrackFormattingField;

        private CT_DocProtect documentProtectionField;

        private CT_OnOff autoFormatOverrideField;

        private CT_OnOff styleLockThemeField;

        private CT_OnOff styleLockQFSetField;

        private CT_TwipsMeasure defaultTabStopField;

        private CT_OnOff autoHyphenationField;

        private CT_DecimalNumber consecutiveHyphenLimitField;

        private CT_TwipsMeasure hyphenationZoneField;

        private CT_OnOff doNotHyphenateCapsField;

        private CT_OnOff showEnvelopeField;

        private CT_DecimalNumber summaryLengthField;

        private CT_String clickAndTypeStyleField;

        private CT_String defaultTableStyleField;

        private CT_OnOff evenAndOddHeadersField;

        private CT_OnOff bookFoldRevPrintingField;

        private CT_OnOff bookFoldPrintingField;

        private CT_DecimalNumber bookFoldPrintingSheetsField;

        private CT_TwipsMeasure drawingGridHorizontalSpacingField;

        private CT_TwipsMeasure drawingGridVerticalSpacingField;

        private CT_DecimalNumber displayHorizontalDrawingGridEveryField;

        private CT_DecimalNumber displayVerticalDrawingGridEveryField;

        private CT_OnOff doNotUseMarginsForDrawingGridOriginField;

        private CT_TwipsMeasure drawingGridHorizontalOriginField;

        private CT_TwipsMeasure drawingGridVerticalOriginField;

        private CT_OnOff doNotShadeFormDataField;

        private CT_OnOff noPunctuationKerningField;

        private CT_CharacterSpacing characterSpacingControlField;

        private CT_OnOff printTwoOnOneField;

        private CT_OnOff strictFirstAndLastCharsField;

        private CT_Kinsoku noLineBreaksAfterField;

        private CT_Kinsoku noLineBreaksBeforeField;

        private CT_OnOff savePreviewPictureField;

        private CT_OnOff doNotValidateAgainstSchemaField;

        private CT_OnOff saveInvalidXmlField;

        private CT_OnOff ignoreMixedContentField;

        private CT_OnOff alwaysShowPlaceholderTextField;

        private CT_OnOff doNotDemarcateInvalidXmlField;

        private CT_OnOff saveXmlDataOnlyField;

        private CT_OnOff useXSLTWhenSavingField;

        private CT_SaveThroughXslt saveThroughXsltField;

        private CT_OnOff showXMLTagsField;

        private CT_OnOff alwaysMergeEmptyNamespaceField;

        private CT_OnOff updateFieldsField;

        private System.Xml.XmlElement[] hdrShapeDefaultsField;

        private CT_FtnDocProps footnotePrField;

        private CT_EdnDocProps endnotePrField;

        private CT_Compat compatField;

        private List<CT_DocVar> docVarsField;

        private CT_DocRsids rsidsField;

        private NPOI.OpenXmlFormats.Shared.CT_MathPr mathPrField;

        private CT_OnOff uiCompat97To2003Field;

        private List<CT_String> attachedSchemaField;

        private CT_Language themeFontLangField;

        private CT_ColorSchemeMapping clrSchemeMappingField;

        private CT_OnOff doNotIncludeSubdocsInStatsField;

        private CT_OnOff doNotAutoCompressPicturesField;

        private CT_Empty forceUpgradeField;

        private CT_Captions captionsField;

        private CT_ReadingModeInkLockDown readModeInkLockDownField;

        private List<CT_SmartTagType> smartTagTypeField;

        private List<CT_Schema> schemaLibraryField;

        private System.Xml.XmlElement[] shapeDefaultsField;

        private CT_OnOff doNotEmbedSmartTagsField;

        private CT_String decimalSymbolField;

        private CT_String listSeparatorField;

        public CT_Settings()
        {
            this.listSeparatorField = new CT_String();
            this.listSeparator.val = ",";
            
            this.decimalSymbolField = new CT_String();
            this.decimalSymbol.val = ".";
            //this.doNotEmbedSmartTagsField = new CT_OnOff();
            this.shapeDefaultsField = new System.Xml.XmlElement[0];
            //this.schemaLibraryField = new List<CT_Schema>();
            //this.smartTagTypeField = new List<CT_SmartTagType>();
            //this.readModeInkLockDownField = new CT_ReadingModeInkLockDown();
            //this.captionsField = new CT_Captions();
            //this.forceUpgradeField = new CT_Empty();
            //this.doNotAutoCompressPicturesField = new CT_OnOff();
            //this.doNotIncludeSubdocsInStatsField = new CT_OnOff();
            this.clrSchemeMappingField = new CT_ColorSchemeMapping();
            this.clrSchemeMapping.bg1 = ST_ColorSchemeIndex.light1;
            this.clrSchemeMapping.t1 = ST_ColorSchemeIndex.dark1;
            this.clrSchemeMapping.bg2 = ST_ColorSchemeIndex.light2;
            this.clrSchemeMapping.t2 = ST_ColorSchemeIndex.dark2;
            this.clrSchemeMapping.accent1 = ST_ColorSchemeIndex.accent1;
            this.clrSchemeMapping.accent2 = ST_ColorSchemeIndex.accent2;
            this.clrSchemeMapping.accent3 = ST_ColorSchemeIndex.accent3;
            this.clrSchemeMapping.accent4 = ST_ColorSchemeIndex.accent4;
            this.clrSchemeMapping.accent5 = ST_ColorSchemeIndex.accent5;
            this.clrSchemeMapping.accent6 = ST_ColorSchemeIndex.accent6;
            this.clrSchemeMapping.hyperlink = ST_ColorSchemeIndex.hyperlink;
            this.clrSchemeMapping.followedHyperlink = ST_ColorSchemeIndex.followedHyperlink;
            this.themeFontLangField = new CT_Language();
            this.themeFontLang.val = "en-US";
            this.themeFontLang.eastAsia = "zh-CN";
            //this.attachedSchemaField = new List<CT_String>();
            //this.uiCompat97To2003Field = new CT_OnOff();
            this.mathPrField = new NPOI.OpenXmlFormats.Shared.CT_MathPr();
            this.rsidsField = new CT_DocRsids();
            //this.docVarsField = new List<CT_DocVar>();
            this.compatField = new CT_Compat();
            //this.endnotePrField = new CT_EdnDocProps();
            //this.footnotePrField = new CT_FtnDocProps();
            //this.hdrShapeDefaultsField = new System.Xml.XmlElement[0];
            //this.updateFieldsField = new CT_OnOff();
            //this.alwaysMergeEmptyNamespaceField = new CT_OnOff();
            //this.showXMLTagsField = new CT_OnOff();
            //this.saveThroughXsltField = new CT_SaveThroughXslt();
            //this.useXSLTWhenSavingField = new CT_OnOff();
            //this.saveXmlDataOnlyField = new CT_OnOff();
            //this.doNotDemarcateInvalidXmlField = new CT_OnOff();
            //this.alwaysShowPlaceholderTextField = new CT_OnOff();
            //this.ignoreMixedContentField = new CT_OnOff();
            //this.saveInvalidXmlField = new CT_OnOff();
            //this.doNotValidateAgainstSchemaField = new CT_OnOff();
            //this.savePreviewPictureField = new CT_OnOff();
            //this.noLineBreaksBeforeField = new CT_Kinsoku();
            //this.noLineBreaksAfterField = new CT_Kinsoku();
            //this.strictFirstAndLastCharsField = new CT_OnOff();
            //this.printTwoOnOneField = new CT_OnOff();
            this.characterSpacingControlField = new CT_CharacterSpacing();
            this.characterSpacingControl.val = ST_CharacterSpacing.compressPunctuation;
            //this.noPunctuationKerningField = new CT_OnOff();
            //this.doNotShadeFormDataField = new CT_OnOff();
            //this.drawingGridVerticalOriginField = new CT_TwipsMeasure();
            //this.drawingGridHorizontalOriginField = new CT_TwipsMeasure();
            //this.doNotUseMarginsForDrawingGridOriginField = new CT_OnOff();
            this.displayVerticalDrawingGridEveryField = new CT_DecimalNumber();
            this.displayVerticalDrawingGridEvery.val = "2";
            this.displayHorizontalDrawingGridEveryField = new CT_DecimalNumber();
            this.displayHorizontalDrawingGridEvery.val = "0";
            this.drawingGridVerticalSpacingField = new CT_TwipsMeasure();
            this.drawingGridVerticalSpacing.val = 156;
            //this.drawingGridHorizontalSpacingField = new CT_TwipsMeasure();
            //this.bookFoldPrintingSheetsField = new CT_DecimalNumber();
            //this.bookFoldPrintingField = new CT_OnOff();
            //this.bookFoldRevPrintingField = new CT_OnOff();
            //this.evenAndOddHeadersField = new CT_OnOff();
            //this.defaultTableStyleField = new CT_String();
            //this.clickAndTypeStyleField = new CT_String();
            //this.summaryLengthField = new CT_DecimalNumber();
            //this.showEnvelopeField = new CT_OnOff();
            //this.doNotHyphenateCapsField = new CT_OnOff();
            //this.hyphenationZoneField = new CT_TwipsMeasure();
            //this.consecutiveHyphenLimitField = new CT_DecimalNumber();
            //this.autoHyphenationField = new CT_OnOff();
            this.defaultTabStopField = new CT_TwipsMeasure();
            this.defaultTabStopField.val = 420;
            //this.styleLockQFSetField = new CT_OnOff();
            //this.styleLockThemeField = new CT_OnOff();
            //this.autoFormatOverrideField = new CT_OnOff();
            //this.documentProtectionField = new CT_DocProtect();
            //this.doNotTrackFormattingField = new CT_OnOff();
            //this.doNotTrackMovesField = new CT_OnOff();
            //this.trackRevisionsField = new CT_OnOff();
            //this.revisionViewField = new CT_TrackChangesView();
            //this.mailMergeField = new CT_MailMerge();
            //this.documentTypeField = new CT_DocType();
            //this.stylePaneSortMethodField = new CT_ShortHexNumber();
            //this.stylePaneFormatFilterField = new CT_ShortHexNumber();
            //this.linkStylesField = new CT_OnOff();
            //this.attachedTemplateField = new CT_Rel();
            //this.formsDesignField = new CT_OnOff();
            //this.proofStateField = new CT_Proof();
            //this.activeWritingStyleField = new List<CT_WritingStyle>();
            //this.hideGrammaticalErrorsField = new CT_OnOff();
            //this.hideSpellingErrorsField = new CT_OnOff();
            //this.gutterAtTopField = new CT_OnOff();
            this.bordersDoNotSurroundFooterField = new CT_OnOff();
            this.bordersDoNotSurroundHeaderField = new CT_OnOff();
            //this.alignBordersAndEdgesField = new CT_OnOff();
            //this.mirrorMarginsField = new CT_OnOff();
            //this.saveFormsDataField = new CT_OnOff();
            //this.saveSubsetFontsField = new CT_OnOff();
            //this.embedSystemFontsField = new CT_OnOff();
            //this.embedTrueTypeFontsField = new CT_OnOff();
            //this.printFormsDataField = new CT_OnOff();
            //this.printFractionalCharacterWidthField = new CT_OnOff();
            //this.printPostScriptOverTextField = new CT_OnOff();
            //this.displayBackgroundShapeField = new CT_OnOff();
            //this.doNotDisplayPageBoundariesField = new CT_OnOff();
            //this.removeDateAndTimeField = new CT_OnOff();
            //this.removePersonalInformationField = new CT_OnOff();
            this.zoomField = new CT_Zoom();
            //this.viewField = new CT_View();
            //this.writeProtectionField = new CT_WriteProtection();
        }

        [XmlElement(Order = 0)]
        public CT_WriteProtection writeProtection
        {
            get
            {
                return this.writeProtectionField;
            }
            set
            {
                this.writeProtectionField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_View view
        {
            get
            {
                return this.viewField;
            }
            set
            {
                this.viewField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_Zoom zoom
        {
            get
            {
                return this.zoomField;
            }
            set
            {
                this.zoomField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_OnOff removePersonalInformation
        {
            get
            {
                return this.removePersonalInformationField;
            }
            set
            {
                this.removePersonalInformationField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_OnOff removeDateAndTime
        {
            get
            {
                return this.removeDateAndTimeField;
            }
            set
            {
                this.removeDateAndTimeField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_OnOff doNotDisplayPageBoundaries
        {
            get
            {
                return this.doNotDisplayPageBoundariesField;
            }
            set
            {
                this.doNotDisplayPageBoundariesField = value;
            }
        }

        [XmlElement(Order = 6)]
        public CT_OnOff displayBackgroundShape
        {
            get
            {
                return this.displayBackgroundShapeField;
            }
            set
            {
                this.displayBackgroundShapeField = value;
            }
        }

        [XmlElement(Order = 7)]
        public CT_OnOff printPostScriptOverText
        {
            get
            {
                return this.printPostScriptOverTextField;
            }
            set
            {
                this.printPostScriptOverTextField = value;
            }
        }

        [XmlElement(Order = 8)]
        public CT_OnOff printFractionalCharacterWidth
        {
            get
            {
                return this.printFractionalCharacterWidthField;
            }
            set
            {
                this.printFractionalCharacterWidthField = value;
            }
        }

        [XmlElement(Order = 9)]
        public CT_OnOff printFormsData
        {
            get
            {
                return this.printFormsDataField;
            }
            set
            {
                this.printFormsDataField = value;
            }
        }

        [XmlElement(Order = 10)]
        public CT_OnOff embedTrueTypeFonts
        {
            get
            {
                return this.embedTrueTypeFontsField;
            }
            set
            {
                this.embedTrueTypeFontsField = value;
            }
        }

        [XmlElement(Order = 11)]
        public CT_OnOff embedSystemFonts
        {
            get
            {
                return this.embedSystemFontsField;
            }
            set
            {
                this.embedSystemFontsField = value;
            }
        }

        [XmlElement(Order = 12)]
        public CT_OnOff saveSubsetFonts
        {
            get
            {
                return this.saveSubsetFontsField;
            }
            set
            {
                this.saveSubsetFontsField = value;
            }
        }

        [XmlElement(Order = 13)]
        public CT_OnOff saveFormsData
        {
            get
            {
                return this.saveFormsDataField;
            }
            set
            {
                this.saveFormsDataField = value;
            }
        }

        [XmlElement(Order = 14)]
        public CT_OnOff mirrorMargins
        {
            get
            {
                return this.mirrorMarginsField;
            }
            set
            {
                this.mirrorMarginsField = value;
            }
        }

        [XmlElement(Order = 15)]
        public CT_OnOff alignBordersAndEdges
        {
            get
            {
                return this.alignBordersAndEdgesField;
            }
            set
            {
                this.alignBordersAndEdgesField = value;
            }
        }

        [XmlElement(Order = 16)]
        public CT_OnOff bordersDoNotSurroundHeader
        {
            get
            {
                return this.bordersDoNotSurroundHeaderField;
            }
            set
            {
                this.bordersDoNotSurroundHeaderField = value;
            }
        }

        [XmlElement(Order = 17)]
        public CT_OnOff bordersDoNotSurroundFooter
        {
            get
            {
                return this.bordersDoNotSurroundFooterField;
            }
            set
            {
                this.bordersDoNotSurroundFooterField = value;
            }
        }

        [XmlElement(Order = 18)]
        public CT_OnOff gutterAtTop
        {
            get
            {
                return this.gutterAtTopField;
            }
            set
            {
                this.gutterAtTopField = value;
            }
        }

        [XmlElement(Order = 19)]
        public CT_OnOff hideSpellingErrors
        {
            get
            {
                return this.hideSpellingErrorsField;
            }
            set
            {
                this.hideSpellingErrorsField = value;
            }
        }

        [XmlElement(Order = 20)]
        public CT_OnOff hideGrammaticalErrors
        {
            get
            {
                return this.hideGrammaticalErrorsField;
            }
            set
            {
                this.hideGrammaticalErrorsField = value;
            }
        }

        [XmlElement("activeWritingStyle", Order = 21)]
        public List<CT_WritingStyle> activeWritingStyle
        {
            get
            {
                return this.activeWritingStyleField;
            }
            set
            {
                this.activeWritingStyleField = value;
            }
        }

        [XmlElement(Order = 22)]
        public CT_Proof proofState
        {
            get
            {
                return this.proofStateField;
            }
            set
            {
                this.proofStateField = value;
            }
        }

        [XmlElement(Order = 23)]
        public CT_OnOff formsDesign
        {
            get
            {
                return this.formsDesignField;
            }
            set
            {
                this.formsDesignField = value;
            }
        }

        [XmlElement(Order = 24)]
        public CT_Rel attachedTemplate
        {
            get
            {
                return this.attachedTemplateField;
            }
            set
            {
                this.attachedTemplateField = value;
            }
        }

        [XmlElement(Order = 25)]
        public CT_OnOff linkStyles
        {
            get
            {
                return this.linkStylesField;
            }
            set
            {
                this.linkStylesField = value;
            }
        }

        [XmlElement(Order = 26)]
        public CT_ShortHexNumber stylePaneFormatFilter
        {
            get
            {
                return this.stylePaneFormatFilterField;
            }
            set
            {
                this.stylePaneFormatFilterField = value;
            }
        }

        [XmlElement(Order = 27)]
        public CT_ShortHexNumber stylePaneSortMethod
        {
            get
            {
                return this.stylePaneSortMethodField;
            }
            set
            {
                this.stylePaneSortMethodField = value;
            }
        }

        [XmlElement(Order = 28)]
        public CT_DocType documentType
        {
            get
            {
                return this.documentTypeField;
            }
            set
            {
                this.documentTypeField = value;
            }
        }

        [XmlElement(Order = 29)]
        public CT_MailMerge mailMerge
        {
            get
            {
                return this.mailMergeField;
            }
            set
            {
                this.mailMergeField = value;
            }
        }

        [XmlElement(Order = 30)]
        public CT_TrackChangesView revisionView
        {
            get
            {
                return this.revisionViewField;
            }
            set
            {
                this.revisionViewField = value;
            }
        }

        [XmlElement(Order = 31)]
        public CT_OnOff trackRevisions
        {
            get
            {
                return this.trackRevisionsField;
            }
            set
            {
                this.trackRevisionsField = value;
            }
        }

        [XmlElement(Order = 32)]
        public CT_OnOff doNotTrackMoves
        {
            get
            {
                return this.doNotTrackMovesField;
            }
            set
            {
                this.doNotTrackMovesField = value;
            }
        }

        [XmlElement(Order = 33)]
        public CT_OnOff doNotTrackFormatting
        {
            get
            {
                return this.doNotTrackFormattingField;
            }
            set
            {
                this.doNotTrackFormattingField = value;
            }
        }

        [XmlElement(Order = 34)]
        public CT_DocProtect documentProtection
        {
            get
            {
                return this.documentProtectionField;
            }
            set
            {
                this.documentProtectionField = value;
            }
        }

        [XmlElement(Order = 35)]
        public CT_OnOff autoFormatOverride
        {
            get
            {
                return this.autoFormatOverrideField;
            }
            set
            {
                this.autoFormatOverrideField = value;
            }
        }

        [XmlElement(Order = 36)]
        public CT_OnOff styleLockTheme
        {
            get
            {
                return this.styleLockThemeField;
            }
            set
            {
                this.styleLockThemeField = value;
            }
        }

        [XmlElement(Order = 37)]
        public CT_OnOff styleLockQFSet
        {
            get
            {
                return this.styleLockQFSetField;
            }
            set
            {
                this.styleLockQFSetField = value;
            }
        }

        [XmlElement(Order = 38)]
        public CT_TwipsMeasure defaultTabStop
        {
            get
            {
                return this.defaultTabStopField;
            }
            set
            {
                this.defaultTabStopField = value;
            }
        }

        [XmlElement(Order = 39)]
        public CT_OnOff autoHyphenation
        {
            get
            {
                return this.autoHyphenationField;
            }
            set
            {
                this.autoHyphenationField = value;
            }
        }

        [XmlElement(Order = 40)]
        public CT_DecimalNumber consecutiveHyphenLimit
        {
            get
            {
                return this.consecutiveHyphenLimitField;
            }
            set
            {
                this.consecutiveHyphenLimitField = value;
            }
        }

        [XmlElement(Order = 41)]
        public CT_TwipsMeasure hyphenationZone
        {
            get
            {
                return this.hyphenationZoneField;
            }
            set
            {
                this.hyphenationZoneField = value;
            }
        }

        [XmlElement(Order = 42)]
        public CT_OnOff doNotHyphenateCaps
        {
            get
            {
                return this.doNotHyphenateCapsField;
            }
            set
            {
                this.doNotHyphenateCapsField = value;
            }
        }

        [XmlElement(Order = 43)]
        public CT_OnOff showEnvelope
        {
            get
            {
                return this.showEnvelopeField;
            }
            set
            {
                this.showEnvelopeField = value;
            }
        }

        [XmlElement(Order = 44)]
        public CT_DecimalNumber summaryLength
        {
            get
            {
                return this.summaryLengthField;
            }
            set
            {
                this.summaryLengthField = value;
            }
        }

        [XmlElement(Order = 45)]
        public CT_String clickAndTypeStyle
        {
            get
            {
                return this.clickAndTypeStyleField;
            }
            set
            {
                this.clickAndTypeStyleField = value;
            }
        }

        [XmlElement(Order = 46)]
        public CT_String defaultTableStyle
        {
            get
            {
                return this.defaultTableStyleField;
            }
            set
            {
                this.defaultTableStyleField = value;
            }
        }

        [XmlElement(Order = 47)]
        public CT_OnOff evenAndOddHeaders
        {
            get
            {
                return this.evenAndOddHeadersField;
            }
            set
            {
                this.evenAndOddHeadersField = value;
            }
        }

        [XmlElement(Order = 48)]
        public CT_OnOff bookFoldRevPrinting
        {
            get
            {
                return this.bookFoldRevPrintingField;
            }
            set
            {
                this.bookFoldRevPrintingField = value;
            }
        }

        [XmlElement(Order = 49)]
        public CT_OnOff bookFoldPrinting
        {
            get
            {
                return this.bookFoldPrintingField;
            }
            set
            {
                this.bookFoldPrintingField = value;
            }
        }

        [XmlElement(Order = 50)]
        public CT_DecimalNumber bookFoldPrintingSheets
        {
            get
            {
                return this.bookFoldPrintingSheetsField;
            }
            set
            {
                this.bookFoldPrintingSheetsField = value;
            }
        }

        [XmlElement(Order = 51)]
        public CT_TwipsMeasure drawingGridHorizontalSpacing
        {
            get
            {
                return this.drawingGridHorizontalSpacingField;
            }
            set
            {
                this.drawingGridHorizontalSpacingField = value;
            }
        }

        [XmlElement(Order = 52)]
        public CT_TwipsMeasure drawingGridVerticalSpacing
        {
            get
            {
                return this.drawingGridVerticalSpacingField;
            }
            set
            {
                this.drawingGridVerticalSpacingField = value;
            }
        }

        [XmlElement(Order = 53)]
        public CT_DecimalNumber displayHorizontalDrawingGridEvery
        {
            get
            {
                return this.displayHorizontalDrawingGridEveryField;
            }
            set
            {
                this.displayHorizontalDrawingGridEveryField = value;
            }
        }

        [XmlElement(Order = 54)]
        public CT_DecimalNumber displayVerticalDrawingGridEvery
        {
            get
            {
                return this.displayVerticalDrawingGridEveryField;
            }
            set
            {
                this.displayVerticalDrawingGridEveryField = value;
            }
        }

        [XmlElement(Order = 55)]
        public CT_OnOff doNotUseMarginsForDrawingGridOrigin
        {
            get
            {
                return this.doNotUseMarginsForDrawingGridOriginField;
            }
            set
            {
                this.doNotUseMarginsForDrawingGridOriginField = value;
            }
        }

        [XmlElement(Order = 56)]
        public CT_TwipsMeasure drawingGridHorizontalOrigin
        {
            get
            {
                return this.drawingGridHorizontalOriginField;
            }
            set
            {
                this.drawingGridHorizontalOriginField = value;
            }
        }

        [XmlElement(Order = 57)]
        public CT_TwipsMeasure drawingGridVerticalOrigin
        {
            get
            {
                return this.drawingGridVerticalOriginField;
            }
            set
            {
                this.drawingGridVerticalOriginField = value;
            }
        }

        [XmlElement(Order = 58)]
        public CT_OnOff doNotShadeFormData
        {
            get
            {
                return this.doNotShadeFormDataField;
            }
            set
            {
                this.doNotShadeFormDataField = value;
            }
        }

        [XmlElement(Order = 59)]
        public CT_OnOff noPunctuationKerning
        {
            get
            {
                return this.noPunctuationKerningField;
            }
            set
            {
                this.noPunctuationKerningField = value;
            }
        }

        [XmlElement(Order = 60)]
        public CT_CharacterSpacing characterSpacingControl
        {
            get
            {
                return this.characterSpacingControlField;
            }
            set
            {
                this.characterSpacingControlField = value;
            }
        }

        [XmlElement(Order = 61)]
        public CT_OnOff printTwoOnOne
        {
            get
            {
                return this.printTwoOnOneField;
            }
            set
            {
                this.printTwoOnOneField = value;
            }
        }

        [XmlElement(Order = 62)]
        public CT_OnOff strictFirstAndLastChars
        {
            get
            {
                return this.strictFirstAndLastCharsField;
            }
            set
            {
                this.strictFirstAndLastCharsField = value;
            }
        }

        [XmlElement(Order = 63)]
        public CT_Kinsoku noLineBreaksAfter
        {
            get
            {
                return this.noLineBreaksAfterField;
            }
            set
            {
                this.noLineBreaksAfterField = value;
            }
        }

        [XmlElement(Order = 64)]
        public CT_Kinsoku noLineBreaksBefore
        {
            get
            {
                return this.noLineBreaksBeforeField;
            }
            set
            {
                this.noLineBreaksBeforeField = value;
            }
        }

        [XmlElement(Order = 65)]
        public CT_OnOff savePreviewPicture
        {
            get
            {
                return this.savePreviewPictureField;
            }
            set
            {
                this.savePreviewPictureField = value;
            }
        }

        [XmlElement(Order = 66)]
        public CT_OnOff doNotValidateAgainstSchema
        {
            get
            {
                return this.doNotValidateAgainstSchemaField;
            }
            set
            {
                this.doNotValidateAgainstSchemaField = value;
            }
        }

        [XmlElement(Order = 67)]
        public CT_OnOff saveInvalidXml
        {
            get
            {
                return this.saveInvalidXmlField;
            }
            set
            {
                this.saveInvalidXmlField = value;
            }
        }

        [XmlElement(Order = 68)]
        public CT_OnOff ignoreMixedContent
        {
            get
            {
                return this.ignoreMixedContentField;
            }
            set
            {
                this.ignoreMixedContentField = value;
            }
        }

        [XmlElement(Order = 69)]
        public CT_OnOff alwaysShowPlaceholderText
        {
            get
            {
                return this.alwaysShowPlaceholderTextField;
            }
            set
            {
                this.alwaysShowPlaceholderTextField = value;
            }
        }

        [XmlElement(Order = 70)]
        public CT_OnOff doNotDemarcateInvalidXml
        {
            get
            {
                return this.doNotDemarcateInvalidXmlField;
            }
            set
            {
                this.doNotDemarcateInvalidXmlField = value;
            }
        }

        [XmlElement(Order = 71)]
        public CT_OnOff saveXmlDataOnly
        {
            get
            {
                return this.saveXmlDataOnlyField;
            }
            set
            {
                this.saveXmlDataOnlyField = value;
            }
        }

        [XmlElement(Order = 72)]
        public CT_OnOff useXSLTWhenSaving
        {
            get
            {
                return this.useXSLTWhenSavingField;
            }
            set
            {
                this.useXSLTWhenSavingField = value;
            }
        }

        [XmlElement(Order = 73)]
        public CT_SaveThroughXslt saveThroughXslt
        {
            get
            {
                return this.saveThroughXsltField;
            }
            set
            {
                this.saveThroughXsltField = value;
            }
        }

        [XmlElement(Order = 74)]
        public CT_OnOff showXMLTags
        {
            get
            {
                return this.showXMLTagsField;
            }
            set
            {
                this.showXMLTagsField = value;
            }
        }

        [XmlElement(Order = 75)]
        public CT_OnOff alwaysMergeEmptyNamespace
        {
            get
            {
                return this.alwaysMergeEmptyNamespaceField;
            }
            set
            {
                this.alwaysMergeEmptyNamespaceField = value;
            }
        }

        [XmlElement(Order = 76)]
        public CT_OnOff updateFields
        {
            get
            {
                return this.updateFieldsField;
            }
            set
            {
                this.updateFieldsField = value;
            }
        }

        [XmlArray(Order = 77)]
        [XmlArrayItem("", Namespace = "urn:schemas-microsoft-com:office:office", IsNullable = false)]
        public System.Xml.XmlElement[] hdrShapeDefaults
        {
            get
            {
                return this.hdrShapeDefaultsField;
            }
            set
            {
                this.hdrShapeDefaultsField = value;
            }
        }

        [XmlElement(Order = 78)]
        public CT_FtnDocProps footnotePr
        {
            get
            {
                return this.footnotePrField;
            }
            set
            {
                this.footnotePrField = value;
            }
        }

        [XmlElement(Order = 79)]
        public CT_EdnDocProps endnotePr
        {
            get
            {
                return this.endnotePrField;
            }
            set
            {
                this.endnotePrField = value;
            }
        }

        [XmlElement(Order = 80)]
        public CT_Compat compat
        {
            get
            {
                return this.compatField;
            }
            set
            {
                this.compatField = value;
            }
        }

        [XmlArray(Order = 81)]
        [XmlArrayItem("docVar", IsNullable = false)]
        public List<CT_DocVar> docVars
        {
            get
            {
                return this.docVarsField;
            }
            set
            {
                this.docVarsField = value;
            }
        }

        [XmlElement(Order = 82)]
        public CT_DocRsids rsids
        {
            get
            {
                return this.rsidsField;
            }
            set
            {
                this.rsidsField = value;
            }
        }

        [XmlElement(Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math", Order = 83)]
        public NPOI.OpenXmlFormats.Shared.CT_MathPr mathPr
        {
            get
            {
                return this.mathPrField;
            }
            set
            {
                this.mathPrField = value;
            }
        }

        [XmlElement(Order = 84)]
        public CT_OnOff uiCompat97To2003
        {
            get
            {
                return this.uiCompat97To2003Field;
            }
            set
            {
                this.uiCompat97To2003Field = value;
            }
        }

        [XmlElement("attachedSchema", Order = 85)]
        public List<CT_String> attachedSchema
        {
            get
            {
                return this.attachedSchemaField;
            }
            set
            {
                this.attachedSchemaField = value;
            }
        }

        [XmlElement(Order = 86)]
        public CT_Language themeFontLang
        {
            get
            {
                return this.themeFontLangField;
            }
            set
            {
                this.themeFontLangField = value;
            }
        }

        [XmlElement(Order = 87)]
        public CT_ColorSchemeMapping clrSchemeMapping
        {
            get
            {
                return this.clrSchemeMappingField;
            }
            set
            {
                this.clrSchemeMappingField = value;
            }
        }

        [XmlElement(Order = 88)]
        public CT_OnOff doNotIncludeSubdocsInStats
        {
            get
            {
                return this.doNotIncludeSubdocsInStatsField;
            }
            set
            {
                this.doNotIncludeSubdocsInStatsField = value;
            }
        }

        [XmlElement(Order = 89)]
        public CT_OnOff doNotAutoCompressPictures
        {
            get
            {
                return this.doNotAutoCompressPicturesField;
            }
            set
            {
                this.doNotAutoCompressPicturesField = value;
            }
        }

        [XmlElement(Order = 90)]
        public CT_Empty forceUpgrade
        {
            get
            {
                return this.forceUpgradeField;
            }
            set
            {
                this.forceUpgradeField = value;
            }
        }

        [XmlElement(Order = 91)]
        public CT_Captions captions
        {
            get
            {
                return this.captionsField;
            }
            set
            {
                this.captionsField = value;
            }
        }

        [XmlElement(Order = 92)]
        public CT_ReadingModeInkLockDown readModeInkLockDown
        {
            get
            {
                return this.readModeInkLockDownField;
            }
            set
            {
                this.readModeInkLockDownField = value;
            }
        }

        [XmlElement("smartTagType", Order = 93)]
        public List<CT_SmartTagType> smartTagType
        {
            get
            {
                return this.smartTagTypeField;
            }
            set
            {
                this.smartTagTypeField = value;
            }
        }

        [XmlArray(Namespace = "http://schemas.openxmlformats.org/schemaLibrary/2006/main", Order = 94)]
        [XmlArrayItem("schema", IsNullable = false)]
        public List<CT_Schema> schemaLibrary
        {
            get
            {
                return this.schemaLibraryField;
            }
            set
            {
                this.schemaLibraryField = value;
            }
        }

        [XmlArray(Order = 95)]
        [XmlArrayItem("", Namespace = "urn:schemas-microsoft-com:office:office", IsNullable = false)]
        public System.Xml.XmlElement[] shapeDefaults
        {
            get
            {
                return this.shapeDefaultsField;
            }
            set
            {
                this.shapeDefaultsField = value;
            }
        }

        [XmlElement(Order = 96)]
        public CT_OnOff doNotEmbedSmartTags
        {
            get
            {
                return this.doNotEmbedSmartTagsField;
            }
            set
            {
                this.doNotEmbedSmartTagsField = value;
            }
        }

        [XmlElement(Order = 97)]
        public CT_String decimalSymbol
        {
            get
            {
                return this.decimalSymbolField;
            }
            set
            {
                this.decimalSymbolField = value;
            }
        }

        [XmlElement(Order = 98)]
        public CT_String listSeparator
        {
            get
            {
                return this.listSeparatorField;
            }
            set
            {
                this.listSeparatorField = value;
            }
        }

        public bool IsSetZoom()
        {
            return this.zoom != null;
        }

        public CT_Zoom AddNewZoom()
        {
            this.zoom = new CT_Zoom();
            return this.zoom;
        }

        public bool IsSetUpdateFields()
        {
            return this.updateFieldsField != null;
        }

        public bool IsSetTrackRevisions()
        {
            return this.trackRevisionsField != null && this.trackRevisionsField.val;
        }

        public CT_OnOff AddNewTrackRevisions()
        {
            this.trackRevisionsField = new CT_OnOff();
            this.trackRevisionsField.val = true;
            return this.trackRevisionsField;
        }

        public void UnsetTrackRevisions()
        {
            this.trackRevisionsField = null;
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_WriteProtection
    {

        private ST_OnOff recommendedField;

        private bool recommendedFieldSpecified;

        private ST_CryptProv cryptProviderTypeField;

        private bool cryptProviderTypeFieldSpecified;

        private ST_AlgClass cryptAlgorithmClassField;

        private bool cryptAlgorithmClassFieldSpecified;

        private ST_AlgType cryptAlgorithmTypeField;

        private bool cryptAlgorithmTypeFieldSpecified;

        private string cryptAlgorithmSidField;

        private string cryptSpinCountField;

        private string cryptProviderField;

        private byte[] algIdExtField;

        private string algIdExtSourceField;

        private byte[] cryptProviderTypeExtField;

        private string cryptProviderTypeExtSourceField;

        private byte[] hashField;

        private byte[] saltField;
        public static CT_WriteProtection Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_WriteProtection ctObj = new CT_WriteProtection();
            if (node.Attributes["w:recommended"] != null)
                ctObj.recommended = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:recommended"].Value);
            if (node.Attributes["w:cryptProviderType"] != null)
                ctObj.cryptProviderType = (ST_CryptProv)Enum.Parse(typeof(ST_CryptProv), node.Attributes["w:cryptProviderType"].Value);
            if (node.Attributes["w:cryptAlgorithmClass"] != null)
                ctObj.cryptAlgorithmClass = (ST_AlgClass)Enum.Parse(typeof(ST_AlgClass), node.Attributes["w:cryptAlgorithmClass"].Value);
            if (node.Attributes["w:cryptAlgorithmType"] != null)
                ctObj.cryptAlgorithmType = (ST_AlgType)Enum.Parse(typeof(ST_AlgType), node.Attributes["w:cryptAlgorithmType"].Value);
            ctObj.cryptAlgorithmSid = XmlHelper.ReadString(node.Attributes["w:cryptAlgorithmSid"]);
            ctObj.cryptSpinCount = XmlHelper.ReadString(node.Attributes["w:cryptSpinCount"]);
            ctObj.cryptProvider = XmlHelper.ReadString(node.Attributes["w:cryptProvider"]);
            ctObj.algIdExt = XmlHelper.ReadBytes(node.Attributes["w:algIdExt"]);
            ctObj.algIdExtSource = XmlHelper.ReadString(node.Attributes["w:algIdExtSource"]);
            ctObj.cryptProviderTypeExt = XmlHelper.ReadBytes(node.Attributes["w:cryptProviderTypeExt"]);
            ctObj.cryptProviderTypeExtSource = XmlHelper.ReadString(node.Attributes["w:cryptProviderTypeExtSource"]);
            ctObj.hash = XmlHelper.ReadBytes(node.Attributes["w:hash"]);
            ctObj.salt = XmlHelper.ReadBytes(node.Attributes["w:salt"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:recommended", this.recommended.ToString());
            XmlHelper.WriteAttribute(sw, "w:cryptProviderType", this.cryptProviderType.ToString());
            XmlHelper.WriteAttribute(sw, "w:cryptAlgorithmClass", this.cryptAlgorithmClass.ToString());
            XmlHelper.WriteAttribute(sw, "w:cryptAlgorithmType", this.cryptAlgorithmType.ToString());
            XmlHelper.WriteAttribute(sw, "w:cryptAlgorithmSid", this.cryptAlgorithmSid);
            XmlHelper.WriteAttribute(sw, "w:cryptSpinCount", this.cryptSpinCount);
            XmlHelper.WriteAttribute(sw, "w:cryptProvider", this.cryptProvider);
            XmlHelper.WriteAttribute(sw, "w:algIdExt", this.algIdExt);
            XmlHelper.WriteAttribute(sw, "w:algIdExtSource", this.algIdExtSource);
            XmlHelper.WriteAttribute(sw, "w:cryptProviderTypeExt", this.cryptProviderTypeExt);
            XmlHelper.WriteAttribute(sw, "w:cryptProviderTypeExtSource", this.cryptProviderTypeExtSource);
            XmlHelper.WriteAttribute(sw, "w:hash", this.hash);
            XmlHelper.WriteAttribute(sw, "w:salt", this.salt);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff recommended
        {
            get
            {
                return this.recommendedField;
            }
            set
            {
                this.recommendedField = value;
            }
        }

        [XmlIgnore]
        public bool recommendedSpecified
        {
            get
            {
                return this.recommendedFieldSpecified;
            }
            set
            {
                this.recommendedFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_CryptProv cryptProviderType
        {
            get
            {
                return this.cryptProviderTypeField;
            }
            set
            {
                this.cryptProviderTypeField = value;
            }
        }

        [XmlIgnore]
        public bool cryptProviderTypeSpecified
        {
            get
            {
                return this.cryptProviderTypeFieldSpecified;
            }
            set
            {
                this.cryptProviderTypeFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_AlgClass cryptAlgorithmClass
        {
            get
            {
                return this.cryptAlgorithmClassField;
            }
            set
            {
                this.cryptAlgorithmClassField = value;
            }
        }

        [XmlIgnore]
        public bool cryptAlgorithmClassSpecified
        {
            get
            {
                return this.cryptAlgorithmClassFieldSpecified;
            }
            set
            {
                this.cryptAlgorithmClassFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_AlgType cryptAlgorithmType
        {
            get
            {
                return this.cryptAlgorithmTypeField;
            }
            set
            {
                this.cryptAlgorithmTypeField = value;
            }
        }

        [XmlIgnore]
        public bool cryptAlgorithmTypeSpecified
        {
            get
            {
                return this.cryptAlgorithmTypeFieldSpecified;
            }
            set
            {
                this.cryptAlgorithmTypeFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string cryptAlgorithmSid
        {
            get
            {
                return this.cryptAlgorithmSidField;
            }
            set
            {
                this.cryptAlgorithmSidField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string cryptSpinCount
        {
            get
            {
                return this.cryptSpinCountField;
            }
            set
            {
                this.cryptSpinCountField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string cryptProvider
        {
            get
            {
                return this.cryptProviderField;
            }
            set
            {
                this.cryptProviderField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] algIdExt
        {
            get
            {
                return this.algIdExtField;
            }
            set
            {
                this.algIdExtField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string algIdExtSource
        {
            get
            {
                return this.algIdExtSourceField;
            }
            set
            {
                this.algIdExtSourceField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] cryptProviderTypeExt
        {
            get
            {
                return this.cryptProviderTypeExtField;
            }
            set
            {
                this.cryptProviderTypeExtField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string cryptProviderTypeExtSource
        {
            get
            {
                return this.cryptProviderTypeExtSourceField;
            }
            set
            {
                this.cryptProviderTypeExtSourceField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "base64Binary")]
        public byte[] hash
        {
            get
            {
                return this.hashField;
            }
            set
            {
                this.hashField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "base64Binary")]
        public byte[] salt
        {
            get
            {
                return this.saltField;
            }
            set
            {
                this.saltField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_CryptProv
    {

    
        rsaAES,

    
        rsaFull,
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_AlgClass
    {

    
        hash,
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_AlgType
    {

    
        typeAny,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_View
    {

        private ST_View valField;
        public static CT_View Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_View ctObj = new CT_View();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_View)Enum.Parse(typeof(ST_View), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_View val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_View
    {

    
        none,

    
        print,

    
        outline,

    
        masterPages,

    
        normal,

    
        web,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Zoom
    {
        public CT_Zoom()
        {
            valField = ST_Zoom.none;
            percent = "100";
        }
        public static CT_Zoom Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Zoom ctObj = new CT_Zoom();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_Zoom)Enum.Parse(typeof(ST_Zoom), node.Attributes["w:val"].Value);
            ctObj.percent = XmlHelper.ReadString(node.Attributes["w:percent"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            if (this.val!= ST_Zoom.none)
                XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            XmlHelper.WriteAttribute(sw, "w:percent", this.percent);
            sw.Write("/>");
        }

        private ST_Zoom valField;

        private bool valFieldSpecified;

        private string percentField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_Zoom val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }

        [XmlIgnore]
        public bool valSpecified
        {
            get
            {
                return this.valFieldSpecified;
            }
            set
            {
                this.valFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string percent
        {
            get
            {
                return this.percentField;
            }
            set
            {
                this.percentField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_Zoom
    {

    
        none,

    
        fullPage,

    
        bestFit,

    
        textFit,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_WritingStyle
    {

        private string langField;

        private string vendorIDField;

        private string dllVersionField;

        private ST_OnOff nlCheckField;

        private bool nlCheckFieldSpecified;

        private ST_OnOff checkStyleField;

        private string appNameField;
        public static CT_WritingStyle Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_WritingStyle ctObj = new CT_WritingStyle();
            ctObj.lang = XmlHelper.ReadString(node.Attributes["w:lang"]);
            ctObj.vendorID = XmlHelper.ReadString(node.Attributes["w:vendorID"]);
            ctObj.dllVersion = XmlHelper.ReadString(node.Attributes["w:dllVersion"]);
            if (node.Attributes["w:nlCheck"] != null)
                ctObj.nlCheck = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:nlCheck"].Value);
            if (node.Attributes["w:checkStyle"] != null)
                ctObj.checkStyle = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:checkStyle"].Value);
            ctObj.appName = XmlHelper.ReadString(node.Attributes["w:appName"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:lang", this.lang);
            XmlHelper.WriteAttribute(sw, "w:vendorID", this.vendorID);
            XmlHelper.WriteAttribute(sw, "w:dllVersion", this.dllVersion);
            XmlHelper.WriteAttribute(sw, "w:nlCheck", this.nlCheck.ToString());
            XmlHelper.WriteAttribute(sw, "w:checkStyle", this.checkStyle.ToString());
            XmlHelper.WriteAttribute(sw, "w:appName", this.appName);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string lang
        {
            get
            {
                return this.langField;
            }
            set
            {
                this.langField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string vendorID
        {
            get
            {
                return this.vendorIDField;
            }
            set
            {
                this.vendorIDField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string dllVersion
        {
            get
            {
                return this.dllVersionField;
            }
            set
            {
                this.dllVersionField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff nlCheck
        {
            get
            {
                return this.nlCheckField;
            }
            set
            {
                this.nlCheckField = value;
            }
        }

        [XmlIgnore]
        public bool nlCheckSpecified
        {
            get
            {
                return this.nlCheckFieldSpecified;
            }
            set
            {
                this.nlCheckFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff checkStyle
        {
            get
            {
                return this.checkStyleField;
            }
            set
            {
                this.checkStyleField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string appName
        {
            get
            {
                return this.appNameField;
            }
            set
            {
                this.appNameField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Proof
    {

        private ST_Proof spellingField;

        private bool spellingFieldSpecified;

        private ST_Proof grammarField;

        private bool grammarFieldSpecified;
        public static CT_Proof Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Proof ctObj = new CT_Proof();
            if (node.Attributes["w:spelling"] != null)
                ctObj.spelling = (ST_Proof)Enum.Parse(typeof(ST_Proof), node.Attributes["w:spelling"].Value);
            if (node.Attributes["w:grammar"] != null)
                ctObj.grammar = (ST_Proof)Enum.Parse(typeof(ST_Proof), node.Attributes["w:grammar"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:spelling", this.spelling.ToString());
            XmlHelper.WriteAttribute(sw, "w:grammar", this.grammar.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_Proof spelling
        {
            get
            {
                return this.spellingField;
            }
            set
            {
                this.spellingField = value;
            }
        }

        [XmlIgnore]
        public bool spellingSpecified
        {
            get
            {
                return this.spellingFieldSpecified;
            }
            set
            {
                this.spellingFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_Proof grammar
        {
            get
            {
                return this.grammarField;
            }
            set
            {
                this.grammarField = value;
            }
        }

        [XmlIgnore]
        public bool grammarSpecified
        {
            get
            {
                return this.grammarFieldSpecified;
            }
            set
            {
                this.grammarFieldSpecified = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_Proof
    {

    
        clean,

    
        dirty,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocType
    {

        private ST_DocType valField;
        public static CT_DocType Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_DocType ctObj = new CT_DocType();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_DocType)Enum.Parse(typeof(ST_DocType), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_DocType val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_DocType
    {

    
        notSpecified,

    
        letter,

    
        eMail,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_MailMerge
    {

        private CT_MailMergeDocType mainDocumentTypeField;

        private CT_OnOff linkToQueryField;

        private CT_MailMergeDataType dataTypeField;

        private CT_String connectStringField;

        private CT_String queryField;

        private CT_Rel dataSourceField;

        private CT_Rel headerSourceField;

        private CT_OnOff doNotSuppressBlankLinesField;

        private CT_MailMergeDest destinationField;

        private CT_String addressFieldNameField;

        private CT_String mailSubjectField;

        private CT_OnOff mailAsAttachmentField;

        private CT_OnOff viewMergedDataField;

        private CT_DecimalNumber activeRecordField;

        private CT_DecimalNumber checkErrorsField;

        private CT_Odso odsoField;
        public static CT_MailMerge Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_MailMerge ctObj = new CT_MailMerge();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "mainDocumentType")
                    ctObj.mainDocumentType = CT_MailMergeDocType.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "linkToQuery")
                    ctObj.linkToQuery = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dataType")
                    ctObj.dataType = CT_MailMergeDataType.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "connectString")
                    ctObj.connectString = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "query")
                    ctObj.query = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dataSource")
                    ctObj.dataSource = CT_Rel.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "headerSource")
                    ctObj.headerSource = CT_Rel.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotSuppressBlankLines")
                    ctObj.doNotSuppressBlankLines = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "destination")
                    ctObj.destination = CT_MailMergeDest.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "addressFieldName")
                    ctObj.addressFieldName = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "mailSubject")
                    ctObj.mailSubject = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "mailAsAttachment")
                    ctObj.mailAsAttachment = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "viewMergedData")
                    ctObj.viewMergedData = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "activeRecord")
                    ctObj.activeRecord = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "checkErrors")
                    ctObj.checkErrors = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "odso")
                    ctObj.odso = CT_Odso.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.mainDocumentType != null)
                this.mainDocumentType.Write(sw, "mainDocumentType");
            if (this.linkToQuery != null)
                this.linkToQuery.Write(sw, "linkToQuery");
            if (this.dataType != null)
                this.dataType.Write(sw, "dataType");
            if (this.connectString != null)
                this.connectString.Write(sw, "connectString");
            if (this.query != null)
                this.query.Write(sw, "query");
            if (this.dataSource != null)
                this.dataSource.Write(sw, "dataSource");
            if (this.headerSource != null)
                this.headerSource.Write(sw, "headerSource");
            if (this.doNotSuppressBlankLines != null)
                this.doNotSuppressBlankLines.Write(sw, "doNotSuppressBlankLines");
            if (this.destination != null)
                this.destination.Write(sw, "destination");
            if (this.addressFieldName != null)
                this.addressFieldName.Write(sw, "addressFieldName");
            if (this.mailSubject != null)
                this.mailSubject.Write(sw, "mailSubject");
            if (this.mailAsAttachment != null)
                this.mailAsAttachment.Write(sw, "mailAsAttachment");
            if (this.viewMergedData != null)
                this.viewMergedData.Write(sw, "viewMergedData");
            if (this.activeRecord != null)
                this.activeRecord.Write(sw, "activeRecord");
            if (this.checkErrors != null)
                this.checkErrors.Write(sw, "checkErrors");
            if (this.odso != null)
                this.odso.Write(sw, "odso");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        public CT_MailMerge()
        {
            //this.odsoField = new CT_Odso();
            //this.checkErrorsField = new CT_DecimalNumber();
            //this.activeRecordField = new CT_DecimalNumber();
            //this.viewMergedDataField = new CT_OnOff();
            //this.mailAsAttachmentField = new CT_OnOff();
            //this.mailSubjectField = new CT_String();
            //this.addressFieldNameField = new CT_String();
            //this.destinationField = new CT_MailMergeDest();
            //this.doNotSuppressBlankLinesField = new CT_OnOff();
            //this.headerSourceField = new CT_Rel();
            //this.dataSourceField = new CT_Rel();
            //this.queryField = new CT_String();
            //this.connectStringField = new CT_String();
            //this.dataTypeField = new CT_MailMergeDataType();
            //this.linkToQueryField = new CT_OnOff();
            //this.mainDocumentTypeField = new CT_MailMergeDocType();
        }

        [XmlElement(Order = 0)]
        public CT_MailMergeDocType mainDocumentType
        {
            get
            {
                return this.mainDocumentTypeField;
            }
            set
            {
                this.mainDocumentTypeField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_OnOff linkToQuery
        {
            get
            {
                return this.linkToQueryField;
            }
            set
            {
                this.linkToQueryField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_MailMergeDataType dataType
        {
            get
            {
                return this.dataTypeField;
            }
            set
            {
                this.dataTypeField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_String connectString
        {
            get
            {
                return this.connectStringField;
            }
            set
            {
                this.connectStringField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_String query
        {
            get
            {
                return this.queryField;
            }
            set
            {
                this.queryField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_Rel dataSource
        {
            get
            {
                return this.dataSourceField;
            }
            set
            {
                this.dataSourceField = value;
            }
        }

        [XmlElement(Order = 6)]
        public CT_Rel headerSource
        {
            get
            {
                return this.headerSourceField;
            }
            set
            {
                this.headerSourceField = value;
            }
        }

        [XmlElement(Order = 7)]
        public CT_OnOff doNotSuppressBlankLines
        {
            get
            {
                return this.doNotSuppressBlankLinesField;
            }
            set
            {
                this.doNotSuppressBlankLinesField = value;
            }
        }

        [XmlElement(Order = 8)]
        public CT_MailMergeDest destination
        {
            get
            {
                return this.destinationField;
            }
            set
            {
                this.destinationField = value;
            }
        }

        [XmlElement(Order = 9)]
        public CT_String addressFieldName
        {
            get
            {
                return this.addressFieldNameField;
            }
            set
            {
                this.addressFieldNameField = value;
            }
        }

        [XmlElement(Order = 10)]
        public CT_String mailSubject
        {
            get
            {
                return this.mailSubjectField;
            }
            set
            {
                this.mailSubjectField = value;
            }
        }

        [XmlElement(Order = 11)]
        public CT_OnOff mailAsAttachment
        {
            get
            {
                return this.mailAsAttachmentField;
            }
            set
            {
                this.mailAsAttachmentField = value;
            }
        }

        [XmlElement(Order = 12)]
        public CT_OnOff viewMergedData
        {
            get
            {
                return this.viewMergedDataField;
            }
            set
            {
                this.viewMergedDataField = value;
            }
        }

        [XmlElement(Order = 13)]
        public CT_DecimalNumber activeRecord
        {
            get
            {
                return this.activeRecordField;
            }
            set
            {
                this.activeRecordField = value;
            }
        }

        [XmlElement(Order = 14)]
        public CT_DecimalNumber checkErrors
        {
            get
            {
                return this.checkErrorsField;
            }
            set
            {
                this.checkErrorsField = value;
            }
        }

        [XmlElement(Order = 15)]
        public CT_Odso odso
        {
            get
            {
                return this.odsoField;
            }
            set
            {
                this.odsoField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_MailMergeDocType
    {

        private ST_MailMergeDocType valField;
        public static CT_MailMergeDocType Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_MailMergeDocType ctObj = new CT_MailMergeDocType();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_MailMergeDocType)Enum.Parse(typeof(ST_MailMergeDocType), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_MailMergeDocType val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_MailMergeDocType
    {

    
        catalog,

    
        envelopes,

    
        mailingLabels,

    
        formLetters,

    
        email,

    
        fax,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_MailMergeDataType
    {

        private ST_MailMergeDataType valField;
        public static CT_MailMergeDataType Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_MailMergeDataType ctObj = new CT_MailMergeDataType();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_MailMergeDataType)Enum.Parse(typeof(ST_MailMergeDataType), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_MailMergeDataType val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_MailMergeDataType
    {

    
        textFile,

    
        database,

    
        spreadsheet,

    
        query,

    
        odbc,

    
        native,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_MailMergeDest
    {

        private ST_MailMergeDest valField;
        public static CT_MailMergeDest Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_MailMergeDest ctObj = new CT_MailMergeDest();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_MailMergeDest)Enum.Parse(typeof(ST_MailMergeDest), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_MailMergeDest val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_MailMergeDest
    {

    
        newDocument,

    
        printer,

    
        email,

    
        fax,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Odso
    {

        private CT_String udlField;

        private CT_String tableField;

        private CT_Rel srcField;

        private CT_DecimalNumber colDelimField;

        private CT_MailMergeSourceType typeField;

        private CT_OnOff fHdrField;

        private List<CT_OdsoFieldMapData> fieldMapDataField;

        private List<CT_Rel> recipientDataField;

        public CT_Odso()
        {
            //this.recipientDataField = new List<CT_Rel>();
            //this.fieldMapDataField = new List<CT_OdsoFieldMapData>();
            //this.fHdrField = new CT_OnOff();
            //this.typeField = new CT_MailMergeSourceType();
            //this.colDelimField = new CT_DecimalNumber();
            //this.srcField = new CT_Rel();
            //this.tableField = new CT_String();
            //this.udlField = new CT_String();
        }
        public static CT_Odso Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Odso ctObj = new CT_Odso();
            ctObj.fieldMapData = new List<CT_OdsoFieldMapData>();
            ctObj.recipientData = new List<CT_Rel>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "udl")
                    ctObj.udl = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "table")
                    ctObj.table = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "src")
                    ctObj.src = CT_Rel.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "colDelim")
                    ctObj.colDelim = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "type")
                    ctObj.type = CT_MailMergeSourceType.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "fHdr")
                    ctObj.fHdr = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "fieldMapData")
                    ctObj.fieldMapData.Add(CT_OdsoFieldMapData.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "recipientData")
                    ctObj.recipientData.Add(CT_Rel.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.udl != null)
                this.udl.Write(sw, "udl");
            if (this.table != null)
                this.table.Write(sw, "table");
            if (this.src != null)
                this.src.Write(sw, "src");
            if (this.colDelim != null)
                this.colDelim.Write(sw, "colDelim");
            if (this.type != null)
                this.type.Write(sw, "type");
            if (this.fHdr != null)
                this.fHdr.Write(sw, "fHdr");
            if (this.fieldMapData != null)
            {
                foreach (CT_OdsoFieldMapData x in this.fieldMapData)
                {
                    x.Write(sw, "fieldMapData");
                }
            }
            if (this.recipientData != null)
            {
                foreach (CT_Rel x in this.recipientData)
                {
                    x.Write(sw, "recipientData");
                }
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }


        [XmlElement(Order = 0)]
        public CT_String udl
        {
            get
            {
                return this.udlField;
            }
            set
            {
                this.udlField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_String table
        {
            get
            {
                return this.tableField;
            }
            set
            {
                this.tableField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_Rel src
        {
            get
            {
                return this.srcField;
            }
            set
            {
                this.srcField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_DecimalNumber colDelim
        {
            get
            {
                return this.colDelimField;
            }
            set
            {
                this.colDelimField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_MailMergeSourceType type
        {
            get
            {
                return this.typeField;
            }
            set
            {
                this.typeField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_OnOff fHdr
        {
            get
            {
                return this.fHdrField;
            }
            set
            {
                this.fHdrField = value;
            }
        }

        [XmlElement("fieldMapData", Order = 6)]
        public List<CT_OdsoFieldMapData> fieldMapData
        {
            get
            {
                return this.fieldMapDataField;
            }
            set
            {
                this.fieldMapDataField = value;
            }
        }

        [XmlElement("recipientData", Order = 7)]
        public List<CT_Rel> recipientData
        {
            get
            {
                return this.recipientDataField;
            }
            set
            {
                this.recipientDataField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_MailMergeSourceType
    {

        private ST_MailMergeSourceType valField;
        public static CT_MailMergeSourceType Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_MailMergeSourceType ctObj = new CT_MailMergeSourceType();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_MailMergeSourceType)Enum.Parse(typeof(ST_MailMergeSourceType), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_MailMergeSourceType val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_MailMergeSourceType
    {

    
        database,

    
        addressBook,

    
        document1,

    
        document2,

    
        text,

    
        email,

    
        native,

    
        legacy,

    
        master,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_OdsoFieldMapData
    {

        private CT_MailMergeOdsoFMDFieldType typeField;

        private CT_String nameField;

        private CT_String mappedNameField;

        private CT_DecimalNumber columnField;

        private CT_Lang lidField;

        private CT_OnOff dynamicAddressField;
        public static CT_OdsoFieldMapData Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_OdsoFieldMapData ctObj = new CT_OdsoFieldMapData();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "type")
                    ctObj.type = CT_MailMergeOdsoFMDFieldType.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "name")
                    ctObj.name = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "mappedName")
                    ctObj.mappedName = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "column")
                    ctObj.column = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lid")
                    ctObj.lid = CT_Lang.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dynamicAddress")
                    ctObj.dynamicAddress = CT_OnOff.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.type != null)
                this.type.Write(sw, "type");
            if (this.name != null)
                this.name.Write(sw, "name");
            if (this.mappedName != null)
                this.mappedName.Write(sw, "mappedName");
            if (this.column != null)
                this.column.Write(sw, "column");
            if (this.lid != null)
                this.lid.Write(sw, "lid");
            if (this.dynamicAddress != null)
                this.dynamicAddress.Write(sw, "dynamicAddress");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        public CT_OdsoFieldMapData()
        {
            //this.dynamicAddressField = new CT_OnOff();
            //this.lidField = new CT_Lang();
            //this.columnField = new CT_DecimalNumber();
            //this.mappedNameField = new CT_String();
            //this.nameField = new CT_String();
            //this.typeField = new CT_MailMergeOdsoFMDFieldType();
        }

        [XmlElement(Order = 0)]
        public CT_MailMergeOdsoFMDFieldType type
        {
            get
            {
                return this.typeField;
            }
            set
            {
                this.typeField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_String name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_String mappedName
        {
            get
            {
                return this.mappedNameField;
            }
            set
            {
                this.mappedNameField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_DecimalNumber column
        {
            get
            {
                return this.columnField;
            }
            set
            {
                this.columnField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_Lang lid
        {
            get
            {
                return this.lidField;
            }
            set
            {
                this.lidField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_OnOff dynamicAddress
        {
            get
            {
                return this.dynamicAddressField;
            }
            set
            {
                this.dynamicAddressField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_MailMergeOdsoFMDFieldType
    {

        private ST_MailMergeOdsoFMDFieldType valField;
        public static CT_MailMergeOdsoFMDFieldType Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_MailMergeOdsoFMDFieldType ctObj = new CT_MailMergeOdsoFMDFieldType();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_MailMergeOdsoFMDFieldType)Enum.Parse(typeof(ST_MailMergeOdsoFMDFieldType), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_MailMergeOdsoFMDFieldType val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_MailMergeOdsoFMDFieldType
    {

    
        @null,

    
        dbColumn,
    }





    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TrackChangesView
    {

        private ST_OnOff markupField;

        private bool markupFieldSpecified;

        private ST_OnOff commentsField;

        private bool commentsFieldSpecified;

        private ST_OnOff insDelField;

        private bool insDelFieldSpecified;

        private ST_OnOff formattingField;

        private bool formattingFieldSpecified;

        private ST_OnOff inkAnnotationsField;

        private bool inkAnnotationsFieldSpecified;
        public static CT_TrackChangesView Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TrackChangesView ctObj = new CT_TrackChangesView();
            if (node.Attributes["w:markup"] != null)
                ctObj.markup = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:markup"].Value);
            if (node.Attributes["w:comments"] != null)
                ctObj.comments = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:comments"].Value);
            if (node.Attributes["w:insDel"] != null)
                ctObj.insDel = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:insDel"].Value);
            if (node.Attributes["w:formatting"] != null)
                ctObj.formatting = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:formatting"].Value);
            if (node.Attributes["w:inkAnnotations"] != null)
                ctObj.inkAnnotations = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:inkAnnotations"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:markup", this.markup.ToString());
            XmlHelper.WriteAttribute(sw, "w:comments", this.comments.ToString());
            XmlHelper.WriteAttribute(sw, "w:insDel", this.insDel.ToString());
            XmlHelper.WriteAttribute(sw, "w:formatting", this.formatting.ToString());
            XmlHelper.WriteAttribute(sw, "w:inkAnnotations", this.inkAnnotations.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff markup
        {
            get
            {
                return this.markupField;
            }
            set
            {
                this.markupField = value;
            }
        }

        [XmlIgnore]
        public bool markupSpecified
        {
            get
            {
                return this.markupFieldSpecified;
            }
            set
            {
                this.markupFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff comments
        {
            get
            {
                return this.commentsField;
            }
            set
            {
                this.commentsField = value;
            }
        }

        [XmlIgnore]
        public bool commentsSpecified
        {
            get
            {
                return this.commentsFieldSpecified;
            }
            set
            {
                this.commentsFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff insDel
        {
            get
            {
                return this.insDelField;
            }
            set
            {
                this.insDelField = value;
            }
        }

        [XmlIgnore]
        public bool insDelSpecified
        {
            get
            {
                return this.insDelFieldSpecified;
            }
            set
            {
                this.insDelFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff formatting
        {
            get
            {
                return this.formattingField;
            }
            set
            {
                this.formattingField = value;
            }
        }

        [XmlIgnore]
        public bool formattingSpecified
        {
            get
            {
                return this.formattingFieldSpecified;
            }
            set
            {
                this.formattingFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff inkAnnotations
        {
            get
            {
                return this.inkAnnotationsField;
            }
            set
            {
                this.inkAnnotationsField = value;
            }
        }

        [XmlIgnore]
        public bool inkAnnotationsSpecified
        {
            get
            {
                return this.inkAnnotationsFieldSpecified;
            }
            set
            {
                this.inkAnnotationsFieldSpecified = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocProtect
    {

        private ST_DocProtect editField;

        private bool editFieldSpecified;

        private ST_OnOff formattingField;

        private bool formattingFieldSpecified;

        private ST_OnOff enforcementField;

        private bool enforcementFieldSpecified;

        private ST_CryptProv cryptProviderTypeField;

        private bool cryptProviderTypeFieldSpecified;

        private ST_AlgClass cryptAlgorithmClassField;

        private bool cryptAlgorithmClassFieldSpecified;

        private ST_AlgType cryptAlgorithmTypeField;

        private bool cryptAlgorithmTypeFieldSpecified;

        private string cryptAlgorithmSidField;

        private string cryptSpinCountField;

        private string cryptProviderField;

        private byte[] algIdExtField;

        private string algIdExtSourceField;

        private byte[] cryptProviderTypeExtField;

        private string cryptProviderTypeExtSourceField;

        private string hashField;

        private string saltField;
        public static CT_DocProtect Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_DocProtect ctObj = new CT_DocProtect();
            if (node.Attributes["w:edit"] != null)
                ctObj.edit = (ST_DocProtect)Enum.Parse(typeof(ST_DocProtect), node.Attributes["w:edit"].Value);
            if (node.Attributes["w:formatting"] != null)
                ctObj.formatting = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:formatting"].Value);
            if (node.Attributes["w:enforcement"] != null)
                ctObj.enforcement = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:enforcement"].Value);
            if (node.Attributes["w:cryptProviderType"] != null)
                ctObj.cryptProviderType = (ST_CryptProv)Enum.Parse(typeof(ST_CryptProv), node.Attributes["w:cryptProviderType"].Value);
            if (node.Attributes["w:cryptAlgorithmClass"] != null)
                ctObj.cryptAlgorithmClass = (ST_AlgClass)Enum.Parse(typeof(ST_AlgClass), node.Attributes["w:cryptAlgorithmClass"].Value);
            if (node.Attributes["w:cryptAlgorithmType"] != null)
                ctObj.cryptAlgorithmType = (ST_AlgType)Enum.Parse(typeof(ST_AlgType), node.Attributes["w:cryptAlgorithmType"].Value);
            ctObj.cryptAlgorithmSid = XmlHelper.ReadString(node.Attributes["w:cryptAlgorithmSid"]);
            ctObj.cryptSpinCount = XmlHelper.ReadString(node.Attributes["w:cryptSpinCount"]);
            ctObj.cryptProvider = XmlHelper.ReadString(node.Attributes["w:cryptProvider"]);
            ctObj.algIdExt = XmlHelper.ReadBytes(node.Attributes["w:algIdExt"]);
            ctObj.algIdExtSource = XmlHelper.ReadString(node.Attributes["w:algIdExtSource"]);
            ctObj.cryptProviderTypeExt = XmlHelper.ReadBytes(node.Attributes["w:cryptProviderTypeExt"]);
            ctObj.cryptProviderTypeExtSource = XmlHelper.ReadString(node.Attributes["w:cryptProviderTypeExtSource"]);
            ctObj.hash = XmlHelper.ReadString(node.Attributes["w:hash"]);
            ctObj.salt = XmlHelper.ReadString(node.Attributes["w:salt"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:edit", this.edit.ToString());
            XmlHelper.WriteAttribute(sw, "w:formatting", this.formatting.ToString());
            XmlHelper.WriteAttribute(sw, "w:enforcement", this.enforcement.ToString());
            XmlHelper.WriteAttribute(sw, "w:cryptProviderType", this.cryptProviderType.ToString());
            XmlHelper.WriteAttribute(sw, "w:cryptAlgorithmClass", this.cryptAlgorithmClass.ToString());
            XmlHelper.WriteAttribute(sw, "w:cryptAlgorithmType", this.cryptAlgorithmType.ToString());
            XmlHelper.WriteAttribute(sw, "w:cryptAlgorithmSid", this.cryptAlgorithmSid);
            XmlHelper.WriteAttribute(sw, "w:cryptSpinCount", this.cryptSpinCount);
            XmlHelper.WriteAttribute(sw, "w:cryptProvider", this.cryptProvider);
            XmlHelper.WriteAttribute(sw, "w:algIdExt", this.algIdExt);
            XmlHelper.WriteAttribute(sw, "w:algIdExtSource", this.algIdExtSource);
            XmlHelper.WriteAttribute(sw, "w:cryptProviderTypeExt", this.cryptProviderTypeExt);
            XmlHelper.WriteAttribute(sw, "w:cryptProviderTypeExtSource", this.cryptProviderTypeExtSource);
            XmlHelper.WriteAttribute(sw, "w:hash", this.hash);
            XmlHelper.WriteAttribute(sw, "w:salt", this.salt);
            sw.Write("/>");
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_DocProtect edit
        {
            get
            {
                return this.editField;
            }
            set
            {
                this.editField = value;
            }
        }

        [XmlIgnore]
        public bool editSpecified
        {
            get
            {
                return this.editField != ST_DocProtect.none;
                //return this.editFieldSpecified;
            }
            set
            {
                this.editFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff formatting
        {
            get
            {
                return this.formattingField;
            }
            set
            {
                this.formattingField = value;
            }
        }

        [XmlIgnore]
        public bool formattingSpecified
        {
            get
            {
                return this.formattingFieldSpecified;
            }
            set
            {
                this.formattingFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff enforcement
        {
            get
            {
                return this.enforcementField;
            }
            set
            {
                this.enforcementField = value;
            }
        }

        [XmlIgnore]
        public bool enforcementSpecified
        {
            get
            {
                return this.editField != ST_DocProtect.none;
                //return this.enforcementFieldSpecified;
            }
            set
            {
                this.enforcementFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_CryptProv cryptProviderType
        {
            get
            {
                return this.cryptProviderTypeField;
            }
            set
            {
                this.cryptProviderTypeField = value;
            }
        }

        [XmlIgnore]
        public bool cryptProviderTypeSpecified
        {
            get
            {
                return this.cryptProviderTypeFieldSpecified;
            }
            set
            {
                this.cryptProviderTypeFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_AlgClass cryptAlgorithmClass
        {
            get
            {
                return this.cryptAlgorithmClassField;
            }
            set
            {
                this.cryptAlgorithmClassField = value;
            }
        }

        [XmlIgnore]
        public bool cryptAlgorithmClassSpecified
        {
            get
            {
                return this.cryptAlgorithmClassFieldSpecified;
            }
            set
            {
                this.cryptAlgorithmClassFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_AlgType cryptAlgorithmType
        {
            get
            {
                return this.cryptAlgorithmTypeField;
            }
            set
            {
                this.cryptAlgorithmTypeField = value;
            }
        }

        [XmlIgnore]
        public bool cryptAlgorithmTypeSpecified
        {
            get
            {
                return this.cryptAlgorithmTypeFieldSpecified;
            }
            set
            {
                this.cryptAlgorithmTypeFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string cryptAlgorithmSid
        {
            get
            {
                return this.cryptAlgorithmSidField;
            }
            set
            {
                this.cryptAlgorithmSidField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string cryptSpinCount
        {
            get
            {
                return this.cryptSpinCountField;
            }
            set
            {
                this.cryptSpinCountField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string cryptProvider
        {
            get
            {
                return this.cryptProviderField;
            }
            set
            {
                this.cryptProviderField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] algIdExt
        {
            get
            {
                return this.algIdExtField;
            }
            set
            {
                this.algIdExtField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string algIdExtSource
        {
            get
            {
                return this.algIdExtSourceField;
            }
            set
            {
                this.algIdExtSourceField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] cryptProviderTypeExt
        {
            get
            {
                return this.cryptProviderTypeExtField;
            }
            set
            {
                this.cryptProviderTypeExtField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string cryptProviderTypeExtSource
        {
            get
            {
                return this.cryptProviderTypeExtSourceField;
            }
            set
            {
                this.cryptProviderTypeExtSourceField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "base64Binary")]
        public string hash
        {
            get
            {
                return this.hashField;
            }
            set
            {
                this.hashField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "base64Binary")]
        public string salt
        {
            get
            {
                return this.saltField;
            }
            set
            {
                this.saltField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_DocProtect
    {

    
        none,

    
        readOnly,

    
        comments,

    
        trackedChanges,

    
        forms,
    }




    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_CharacterSpacing
    {

        private ST_CharacterSpacing valField;
        public static CT_CharacterSpacing Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CharacterSpacing ctObj = new CT_CharacterSpacing();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_CharacterSpacing)Enum.Parse(typeof(ST_CharacterSpacing), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write("/>");
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_CharacterSpacing val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_CharacterSpacing
    {

    
        doNotCompress,

    
        compressPunctuation,

    
        compressPunctuationAndJapaneseKana,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Kinsoku
    {

        private string langField;

        private string valField;
        public static CT_Kinsoku Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Kinsoku ctObj = new CT_Kinsoku();
            ctObj.lang = XmlHelper.ReadString(node.Attributes["w:lang"]);
            ctObj.val = XmlHelper.ReadString(node.Attributes["w:val"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:lang", this.lang);
            XmlHelper.WriteAttribute(sw, "w:val", this.val);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string lang
        {
            get
            {
                return this.langField;
            }
            set
            {
                this.langField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SaveThroughXslt
    {

        private string idField;

        private string solutionIDField;
        public static CT_SaveThroughXslt Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SaveThroughXslt ctObj = new CT_SaveThroughXslt();
            ctObj.id = XmlHelper.ReadString(node.Attributes["r:id"]);
            ctObj.solutionID = XmlHelper.ReadString(node.Attributes["w:solutionID"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "r:id", this.id);
            XmlHelper.WriteAttribute(sw, "w:solutionID", this.solutionID);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id
        {
            get
            {
                return this.idField;
            }
            set
            {
                this.idField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string solutionID
        {
            get
            {
                return this.solutionIDField;
            }
            set
            {
                this.solutionIDField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Compat
    {

        private CT_OnOff useSingleBorderforContiguousCellsField;

        private CT_OnOff wpJustificationField;

        private CT_OnOff noTabHangIndField;

        private CT_OnOff noLeadingField;

        private CT_OnOff spaceForULField;

        private CT_OnOff noColumnBalanceField;

        private CT_OnOff balanceSingleByteDoubleByteWidthField;

        private CT_OnOff noExtraLineSpacingField;

        private CT_OnOff doNotLeaveBackslashAloneField;

        private CT_OnOff ulTrailSpaceField;

        private CT_OnOff doNotExpandShiftReturnField;

        private CT_OnOff spacingInWholePointsField;

        private CT_OnOff lineWrapLikeWord6Field;

        private CT_OnOff printBodyTextBeforeHeaderField;

        private CT_OnOff printColBlackField;

        private CT_OnOff wpSpaceWidthField;

        private CT_OnOff showBreaksInFramesField;

        private CT_OnOff subFontBySizeField;

        private CT_OnOff suppressBottomSpacingField;

        private CT_OnOff suppressTopSpacingField;

        private CT_OnOff suppressSpacingAtTopOfPageField;

        private CT_OnOff suppressTopSpacingWPField;

        private CT_OnOff suppressSpBfAfterPgBrkField;

        private CT_OnOff swapBordersFacingPagesField;

        private CT_OnOff convMailMergeEscField;

        private CT_OnOff truncateFontHeightsLikeWP6Field;

        private CT_OnOff mwSmallCapsField;

        private CT_OnOff usePrinterMetricsField;

        private CT_OnOff doNotSuppressParagraphBordersField;

        private CT_OnOff wrapTrailSpacesField;

        private CT_OnOff footnoteLayoutLikeWW8Field;

        private CT_OnOff shapeLayoutLikeWW8Field;

        private CT_OnOff alignTablesRowByRowField;

        private CT_OnOff forgetLastTabAlignmentField;

        private CT_OnOff adjustLineHeightInTableField;

        private CT_OnOff autoSpaceLikeWord95Field;

        private CT_OnOff noSpaceRaiseLowerField;

        private CT_OnOff doNotUseHTMLParagraphAutoSpacingField;

        private CT_OnOff layoutRawTableWidthField;

        private CT_OnOff layoutTableRowsApartField;

        private CT_OnOff useWord97LineBreakRulesField;

        private CT_OnOff doNotBreakWrappedTablesField;

        private CT_OnOff doNotSnapToGridInCellField;

        private CT_OnOff selectFldWithFirstOrLastCharField;

        private CT_OnOff applyBreakingRulesField;

        private CT_OnOff doNotWrapTextWithPunctField;

        private CT_OnOff doNotUseEastAsianBreakRulesField;

        private CT_OnOff useWord2002TableStyleRulesField;

        private CT_OnOff growAutofitField;

        private CT_OnOff useFELayoutField;

        private CT_OnOff useNormalStyleForListField;

        private CT_OnOff doNotUseIndentAsNumberingTabStopField;

        private CT_OnOff useAltKinsokuLineBreakRulesField;

        private CT_OnOff allowSpaceOfSameStyleInTableField;

        private CT_OnOff doNotSuppressIndentationField;

        private CT_OnOff doNotAutofitConstrainedTablesField;

        private CT_OnOff autofitToFirstFixedWidthCellField;

        private CT_OnOff underlineTabInNumListField;

        private CT_OnOff displayHangulFixedWidthField;

        private CT_OnOff splitPgBreakAndParaMarkField;

        private CT_OnOff doNotVertAlignCellWithSpField;

        private CT_OnOff doNotBreakConstrainedForcedTableField;

        private CT_OnOff doNotVertAlignInTxbxField;

        private CT_OnOff useAnsiKerningPairsField;

        private CT_OnOff cachedColBalanceField;
        public static CT_Compat Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Compat ctObj = new CT_Compat();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "useSingleBorderforContiguousCells")
                    ctObj.useSingleBorderforContiguousCells = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "wpJustification")
                    ctObj.wpJustification = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "noTabHangInd")
                    ctObj.noTabHangInd = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "noLeading")
                    ctObj.noLeading = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "spaceForUL")
                    ctObj.spaceForUL = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "noColumnBalance")
                    ctObj.noColumnBalance = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "balanceSingleByteDoubleByteWidth")
                    ctObj.balanceSingleByteDoubleByteWidth = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "noExtraLineSpacing")
                    ctObj.noExtraLineSpacing = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotLeaveBackslashAlone")
                    ctObj.doNotLeaveBackslashAlone = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "ulTrailSpace")
                    ctObj.ulTrailSpace = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotExpandShiftReturn")
                    ctObj.doNotExpandShiftReturn = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "spacingInWholePoints")
                    ctObj.spacingInWholePoints = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "lineWrapLikeWord6")
                    ctObj.lineWrapLikeWord6 = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "printBodyTextBeforeHeader")
                    ctObj.printBodyTextBeforeHeader = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "printColBlack")
                    ctObj.printColBlack = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "wpSpaceWidth")
                    ctObj.wpSpaceWidth = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "showBreaksInFrames")
                    ctObj.showBreaksInFrames = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "subFontBySize")
                    ctObj.subFontBySize = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "suppressBottomSpacing")
                    ctObj.suppressBottomSpacing = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "suppressTopSpacing")
                    ctObj.suppressTopSpacing = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "suppressSpacingAtTopOfPage")
                    ctObj.suppressSpacingAtTopOfPage = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "suppressTopSpacingWP")
                    ctObj.suppressTopSpacingWP = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "suppressSpBfAfterPgBrk")
                    ctObj.suppressSpBfAfterPgBrk = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "swapBordersFacingPages")
                    ctObj.swapBordersFacingPages = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "convMailMergeEsc")
                    ctObj.convMailMergeEsc = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "truncateFontHeightsLikeWP6")
                    ctObj.truncateFontHeightsLikeWP6 = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "mwSmallCaps")
                    ctObj.mwSmallCaps = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "usePrinterMetrics")
                    ctObj.usePrinterMetrics = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotSuppressParagraphBorders")
                    ctObj.doNotSuppressParagraphBorders = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "wrapTrailSpaces")
                    ctObj.wrapTrailSpaces = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "footnoteLayoutLikeWW8")
                    ctObj.footnoteLayoutLikeWW8 = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "shapeLayoutLikeWW8")
                    ctObj.shapeLayoutLikeWW8 = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "alignTablesRowByRow")
                    ctObj.alignTablesRowByRow = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "forgetLastTabAlignment")
                    ctObj.forgetLastTabAlignment = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "adjustLineHeightInTable")
                    ctObj.adjustLineHeightInTable = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "autoSpaceLikeWord95")
                    ctObj.autoSpaceLikeWord95 = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "noSpaceRaiseLower")
                    ctObj.noSpaceRaiseLower = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotUseHTMLParagraphAutoSpacing")
                    ctObj.doNotUseHTMLParagraphAutoSpacing = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "layoutRawTableWidth")
                    ctObj.layoutRawTableWidth = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "layoutTableRowsApart")
                    ctObj.layoutTableRowsApart = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "useWord97LineBreakRules")
                    ctObj.useWord97LineBreakRules = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotBreakWrappedTables")
                    ctObj.doNotBreakWrappedTables = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotSnapToGridInCell")
                    ctObj.doNotSnapToGridInCell = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "selectFldWithFirstOrLastChar")
                    ctObj.selectFldWithFirstOrLastChar = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "applyBreakingRules")
                    ctObj.applyBreakingRules = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotWrapTextWithPunct")
                    ctObj.doNotWrapTextWithPunct = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotUseEastAsianBreakRules")
                    ctObj.doNotUseEastAsianBreakRules = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "useWord2002TableStyleRules")
                    ctObj.useWord2002TableStyleRules = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "growAutofit")
                    ctObj.growAutofit = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "useFELayout")
                    ctObj.useFELayout = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "useNormalStyleForList")
                    ctObj.useNormalStyleForList = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotUseIndentAsNumberingTabStop")
                    ctObj.doNotUseIndentAsNumberingTabStop = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "useAltKinsokuLineBreakRules")
                    ctObj.useAltKinsokuLineBreakRules = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "allowSpaceOfSameStyleInTable")
                    ctObj.allowSpaceOfSameStyleInTable = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotSuppressIndentation")
                    ctObj.doNotSuppressIndentation = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotAutofitConstrainedTables")
                    ctObj.doNotAutofitConstrainedTables = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "autofitToFirstFixedWidthCell")
                    ctObj.autofitToFirstFixedWidthCell = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "underlineTabInNumList")
                    ctObj.underlineTabInNumList = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "displayHangulFixedWidth")
                    ctObj.displayHangulFixedWidth = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "splitPgBreakAndParaMark")
                    ctObj.splitPgBreakAndParaMark = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotVertAlignCellWithSp")
                    ctObj.doNotVertAlignCellWithSp = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotBreakConstrainedForcedTable")
                    ctObj.doNotBreakConstrainedForcedTable = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "doNotVertAlignInTxbx")
                    ctObj.doNotVertAlignInTxbx = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "useAnsiKerningPairs")
                    ctObj.useAnsiKerningPairs = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cachedColBalance")
                    ctObj.cachedColBalance = CT_OnOff.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.useSingleBorderforContiguousCells != null)
                this.useSingleBorderforContiguousCells.Write(sw, "useSingleBorderforContiguousCells");
            if (this.wpJustification != null)
                this.wpJustification.Write(sw, "wpJustification");
            if (this.noTabHangInd != null)
                this.noTabHangInd.Write(sw, "noTabHangInd");
            if (this.noLeading != null)
                this.noLeading.Write(sw, "noLeading");
            if (this.spaceForUL != null)
                this.spaceForUL.Write(sw, "spaceForUL");
            if (this.noColumnBalance != null)
                this.noColumnBalance.Write(sw, "noColumnBalance");
            if (this.balanceSingleByteDoubleByteWidth != null)
                this.balanceSingleByteDoubleByteWidth.Write(sw, "balanceSingleByteDoubleByteWidth");
            if (this.noExtraLineSpacing != null)
                this.noExtraLineSpacing.Write(sw, "noExtraLineSpacing");
            if (this.doNotLeaveBackslashAlone != null)
                this.doNotLeaveBackslashAlone.Write(sw, "doNotLeaveBackslashAlone");
            if (this.ulTrailSpace != null)
                this.ulTrailSpace.Write(sw, "ulTrailSpace");
            if (this.doNotExpandShiftReturn != null)
                this.doNotExpandShiftReturn.Write(sw, "doNotExpandShiftReturn");
            if (this.spacingInWholePoints != null)
                this.spacingInWholePoints.Write(sw, "spacingInWholePoints");
            if (this.lineWrapLikeWord6 != null)
                this.lineWrapLikeWord6.Write(sw, "lineWrapLikeWord6");
            if (this.printBodyTextBeforeHeader != null)
                this.printBodyTextBeforeHeader.Write(sw, "printBodyTextBeforeHeader");
            if (this.printColBlack != null)
                this.printColBlack.Write(sw, "printColBlack");
            if (this.wpSpaceWidth != null)
                this.wpSpaceWidth.Write(sw, "wpSpaceWidth");
            if (this.showBreaksInFrames != null)
                this.showBreaksInFrames.Write(sw, "showBreaksInFrames");
            if (this.subFontBySize != null)
                this.subFontBySize.Write(sw, "subFontBySize");
            if (this.suppressBottomSpacing != null)
                this.suppressBottomSpacing.Write(sw, "suppressBottomSpacing");
            if (this.suppressTopSpacing != null)
                this.suppressTopSpacing.Write(sw, "suppressTopSpacing");
            if (this.suppressSpacingAtTopOfPage != null)
                this.suppressSpacingAtTopOfPage.Write(sw, "suppressSpacingAtTopOfPage");
            if (this.suppressTopSpacingWP != null)
                this.suppressTopSpacingWP.Write(sw, "suppressTopSpacingWP");
            if (this.suppressSpBfAfterPgBrk != null)
                this.suppressSpBfAfterPgBrk.Write(sw, "suppressSpBfAfterPgBrk");
            if (this.swapBordersFacingPages != null)
                this.swapBordersFacingPages.Write(sw, "swapBordersFacingPages");
            if (this.convMailMergeEsc != null)
                this.convMailMergeEsc.Write(sw, "convMailMergeEsc");
            if (this.truncateFontHeightsLikeWP6 != null)
                this.truncateFontHeightsLikeWP6.Write(sw, "truncateFontHeightsLikeWP6");
            if (this.mwSmallCaps != null)
                this.mwSmallCaps.Write(sw, "mwSmallCaps");
            if (this.usePrinterMetrics != null)
                this.usePrinterMetrics.Write(sw, "usePrinterMetrics");
            if (this.doNotSuppressParagraphBorders != null)
                this.doNotSuppressParagraphBorders.Write(sw, "doNotSuppressParagraphBorders");
            if (this.wrapTrailSpaces != null)
                this.wrapTrailSpaces.Write(sw, "wrapTrailSpaces");
            if (this.footnoteLayoutLikeWW8 != null)
                this.footnoteLayoutLikeWW8.Write(sw, "footnoteLayoutLikeWW8");
            if (this.shapeLayoutLikeWW8 != null)
                this.shapeLayoutLikeWW8.Write(sw, "shapeLayoutLikeWW8");
            if (this.alignTablesRowByRow != null)
                this.alignTablesRowByRow.Write(sw, "alignTablesRowByRow");
            if (this.forgetLastTabAlignment != null)
                this.forgetLastTabAlignment.Write(sw, "forgetLastTabAlignment");
            if (this.adjustLineHeightInTable != null)
                this.adjustLineHeightInTable.Write(sw, "adjustLineHeightInTable");
            if (this.autoSpaceLikeWord95 != null)
                this.autoSpaceLikeWord95.Write(sw, "autoSpaceLikeWord95");
            if (this.noSpaceRaiseLower != null)
                this.noSpaceRaiseLower.Write(sw, "noSpaceRaiseLower");
            if (this.doNotUseHTMLParagraphAutoSpacing != null)
                this.doNotUseHTMLParagraphAutoSpacing.Write(sw, "doNotUseHTMLParagraphAutoSpacing");
            if (this.layoutRawTableWidth != null)
                this.layoutRawTableWidth.Write(sw, "layoutRawTableWidth");
            if (this.layoutTableRowsApart != null)
                this.layoutTableRowsApart.Write(sw, "layoutTableRowsApart");
            if (this.useWord97LineBreakRules != null)
                this.useWord97LineBreakRules.Write(sw, "useWord97LineBreakRules");
            if (this.doNotBreakWrappedTables != null)
                this.doNotBreakWrappedTables.Write(sw, "doNotBreakWrappedTables");
            if (this.doNotSnapToGridInCell != null)
                this.doNotSnapToGridInCell.Write(sw, "doNotSnapToGridInCell");
            if (this.selectFldWithFirstOrLastChar != null)
                this.selectFldWithFirstOrLastChar.Write(sw, "selectFldWithFirstOrLastChar");
            if (this.applyBreakingRules != null)
                this.applyBreakingRules.Write(sw, "applyBreakingRules");
            if (this.doNotWrapTextWithPunct != null)
                this.doNotWrapTextWithPunct.Write(sw, "doNotWrapTextWithPunct");
            if (this.doNotUseEastAsianBreakRules != null)
                this.doNotUseEastAsianBreakRules.Write(sw, "doNotUseEastAsianBreakRules");
            if (this.useWord2002TableStyleRules != null)
                this.useWord2002TableStyleRules.Write(sw, "useWord2002TableStyleRules");
            if (this.growAutofit != null)
                this.growAutofit.Write(sw, "growAutofit");
            if (this.useFELayout != null)
                this.useFELayout.Write(sw, "useFELayout");
            if (this.useNormalStyleForList != null)
                this.useNormalStyleForList.Write(sw, "useNormalStyleForList");
            if (this.doNotUseIndentAsNumberingTabStop != null)
                this.doNotUseIndentAsNumberingTabStop.Write(sw, "doNotUseIndentAsNumberingTabStop");
            if (this.useAltKinsokuLineBreakRules != null)
                this.useAltKinsokuLineBreakRules.Write(sw, "useAltKinsokuLineBreakRules");
            if (this.allowSpaceOfSameStyleInTable != null)
                this.allowSpaceOfSameStyleInTable.Write(sw, "allowSpaceOfSameStyleInTable");
            if (this.doNotSuppressIndentation != null)
                this.doNotSuppressIndentation.Write(sw, "doNotSuppressIndentation");
            if (this.doNotAutofitConstrainedTables != null)
                this.doNotAutofitConstrainedTables.Write(sw, "doNotAutofitConstrainedTables");
            if (this.autofitToFirstFixedWidthCell != null)
                this.autofitToFirstFixedWidthCell.Write(sw, "autofitToFirstFixedWidthCell");
            if (this.underlineTabInNumList != null)
                this.underlineTabInNumList.Write(sw, "underlineTabInNumList");
            if (this.displayHangulFixedWidth != null)
                this.displayHangulFixedWidth.Write(sw, "displayHangulFixedWidth");
            if (this.splitPgBreakAndParaMark != null)
                this.splitPgBreakAndParaMark.Write(sw, "splitPgBreakAndParaMark");
            if (this.doNotVertAlignCellWithSp != null)
                this.doNotVertAlignCellWithSp.Write(sw, "doNotVertAlignCellWithSp");
            if (this.doNotBreakConstrainedForcedTable != null)
                this.doNotBreakConstrainedForcedTable.Write(sw, "doNotBreakConstrainedForcedTable");
            if (this.doNotVertAlignInTxbx != null)
                this.doNotVertAlignInTxbx.Write(sw, "doNotVertAlignInTxbx");
            if (this.useAnsiKerningPairs != null)
                this.useAnsiKerningPairs.Write(sw, "useAnsiKerningPairs");
            if (this.cachedColBalance != null)
                this.cachedColBalance.Write(sw, "cachedColBalance");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        public CT_Compat()
        {
            //this.cachedColBalanceField = new CT_OnOff();
            //this.useAnsiKerningPairsField = new CT_OnOff();
            //this.doNotVertAlignInTxbxField = new CT_OnOff();
            //this.doNotBreakConstrainedForcedTableField = new CT_OnOff();
            //this.doNotVertAlignCellWithSpField = new CT_OnOff();
            //this.splitPgBreakAndParaMarkField = new CT_OnOff();
            //this.displayHangulFixedWidthField = new CT_OnOff();
            //this.underlineTabInNumListField = new CT_OnOff();
            //this.autofitToFirstFixedWidthCellField = new CT_OnOff();
            //this.doNotAutofitConstrainedTablesField = new CT_OnOff();
            //this.doNotSuppressIndentationField = new CT_OnOff();
            //this.allowSpaceOfSameStyleInTableField = new CT_OnOff();
            //this.useAltKinsokuLineBreakRulesField = new CT_OnOff();
            //this.doNotUseIndentAsNumberingTabStopField = new CT_OnOff();
            //this.useNormalStyleForListField = new CT_OnOff();
            this.useFELayoutField = new CT_OnOff();
            //this.growAutofitField = new CT_OnOff();
            //this.useWord2002TableStyleRulesField = new CT_OnOff();
            //this.doNotUseEastAsianBreakRulesField = new CT_OnOff();
            //this.doNotWrapTextWithPunctField = new CT_OnOff();
            //this.applyBreakingRulesField = new CT_OnOff();
            //this.selectFldWithFirstOrLastCharField = new CT_OnOff();
            //this.doNotSnapToGridInCellField = new CT_OnOff();
            //this.doNotBreakWrappedTablesField = new CT_OnOff();
            //this.useWord97LineBreakRulesField = new CT_OnOff();
            //this.layoutTableRowsApartField = new CT_OnOff();
            //this.layoutRawTableWidthField = new CT_OnOff();
            //this.doNotUseHTMLParagraphAutoSpacingField = new CT_OnOff();
            //this.noSpaceRaiseLowerField = new CT_OnOff();
            //this.autoSpaceLikeWord95Field = new CT_OnOff();
            this.adjustLineHeightInTableField = new CT_OnOff();
            //this.forgetLastTabAlignmentField = new CT_OnOff();
            //this.alignTablesRowByRowField = new CT_OnOff();
            //this.shapeLayoutLikeWW8Field = new CT_OnOff();
            //this.footnoteLayoutLikeWW8Field = new CT_OnOff();
            //this.wrapTrailSpacesField = new CT_OnOff();
            //this.doNotSuppressParagraphBordersField = new CT_OnOff();
            //this.usePrinterMetricsField = new CT_OnOff();
            //this.mwSmallCapsField = new CT_OnOff();
            //this.truncateFontHeightsLikeWP6Field = new CT_OnOff();
            //this.convMailMergeEscField = new CT_OnOff();
            //this.swapBordersFacingPagesField = new CT_OnOff();
            //this.suppressSpBfAfterPgBrkField = new CT_OnOff();
            //this.suppressTopSpacingWPField = new CT_OnOff();
            //this.suppressSpacingAtTopOfPageField = new CT_OnOff();
            //this.suppressTopSpacingField = new CT_OnOff();
            //this.suppressBottomSpacingField = new CT_OnOff();
            //this.subFontBySizeField = new CT_OnOff();
            //this.showBreaksInFramesField = new CT_OnOff();
            //this.wpSpaceWidthField = new CT_OnOff();
            //this.printColBlackField = new CT_OnOff();
            //this.printBodyTextBeforeHeaderField = new CT_OnOff();
            //this.lineWrapLikeWord6Field = new CT_OnOff();
            //this.spacingInWholePointsField = new CT_OnOff();
            this.doNotExpandShiftReturnField = new CT_OnOff();
            this.ulTrailSpaceField = new CT_OnOff();
            this.doNotLeaveBackslashAloneField = new CT_OnOff();
            //this.noExtraLineSpacingField = new CT_OnOff();
            this.balanceSingleByteDoubleByteWidthField = new CT_OnOff();
            //this.noColumnBalanceField = new CT_OnOff();
            this.spaceForULField = new CT_OnOff();
            //this.noLeadingField = new CT_OnOff();
            //this.noTabHangIndField = new CT_OnOff();
            //this.wpJustificationField = new CT_OnOff();
            //this.useSingleBorderforContiguousCellsField = new CT_OnOff();
        }

        [XmlElement(Order = 0)]
        public CT_OnOff useSingleBorderforContiguousCells
        {
            get
            {
                return this.useSingleBorderforContiguousCellsField;
            }
            set
            {
                this.useSingleBorderforContiguousCellsField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_OnOff wpJustification
        {
            get
            {
                return this.wpJustificationField;
            }
            set
            {
                this.wpJustificationField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_OnOff noTabHangInd
        {
            get
            {
                return this.noTabHangIndField;
            }
            set
            {
                this.noTabHangIndField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_OnOff noLeading
        {
            get
            {
                return this.noLeadingField;
            }
            set
            {
                this.noLeadingField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_OnOff spaceForUL
        {
            get
            {
                return this.spaceForULField;
            }
            set
            {
                this.spaceForULField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_OnOff noColumnBalance
        {
            get
            {
                return this.noColumnBalanceField;
            }
            set
            {
                this.noColumnBalanceField = value;
            }
        }

        [XmlElement(Order = 6)]
        public CT_OnOff balanceSingleByteDoubleByteWidth
        {
            get
            {
                return this.balanceSingleByteDoubleByteWidthField;
            }
            set
            {
                this.balanceSingleByteDoubleByteWidthField = value;
            }
        }

        [XmlElement(Order = 7)]
        public CT_OnOff noExtraLineSpacing
        {
            get
            {
                return this.noExtraLineSpacingField;
            }
            set
            {
                this.noExtraLineSpacingField = value;
            }
        }

        [XmlElement(Order = 8)]
        public CT_OnOff doNotLeaveBackslashAlone
        {
            get
            {
                return this.doNotLeaveBackslashAloneField;
            }
            set
            {
                this.doNotLeaveBackslashAloneField = value;
            }
        }

        [XmlElement(Order = 9)]
        public CT_OnOff ulTrailSpace
        {
            get
            {
                return this.ulTrailSpaceField;
            }
            set
            {
                this.ulTrailSpaceField = value;
            }
        }

        [XmlElement(Order = 10)]
        public CT_OnOff doNotExpandShiftReturn
        {
            get
            {
                return this.doNotExpandShiftReturnField;
            }
            set
            {
                this.doNotExpandShiftReturnField = value;
            }
        }

        [XmlElement(Order = 11)]
        public CT_OnOff spacingInWholePoints
        {
            get
            {
                return this.spacingInWholePointsField;
            }
            set
            {
                this.spacingInWholePointsField = value;
            }
        }

        [XmlElement(Order = 12)]
        public CT_OnOff lineWrapLikeWord6
        {
            get
            {
                return this.lineWrapLikeWord6Field;
            }
            set
            {
                this.lineWrapLikeWord6Field = value;
            }
        }

        [XmlElement(Order = 13)]
        public CT_OnOff printBodyTextBeforeHeader
        {
            get
            {
                return this.printBodyTextBeforeHeaderField;
            }
            set
            {
                this.printBodyTextBeforeHeaderField = value;
            }
        }

        [XmlElement(Order = 14)]
        public CT_OnOff printColBlack
        {
            get
            {
                return this.printColBlackField;
            }
            set
            {
                this.printColBlackField = value;
            }
        }

        [XmlElement(Order = 15)]
        public CT_OnOff wpSpaceWidth
        {
            get
            {
                return this.wpSpaceWidthField;
            }
            set
            {
                this.wpSpaceWidthField = value;
            }
        }

        [XmlElement(Order = 16)]
        public CT_OnOff showBreaksInFrames
        {
            get
            {
                return this.showBreaksInFramesField;
            }
            set
            {
                this.showBreaksInFramesField = value;
            }
        }

        [XmlElement(Order = 17)]
        public CT_OnOff subFontBySize
        {
            get
            {
                return this.subFontBySizeField;
            }
            set
            {
                this.subFontBySizeField = value;
            }
        }

        [XmlElement(Order = 18)]
        public CT_OnOff suppressBottomSpacing
        {
            get
            {
                return this.suppressBottomSpacingField;
            }
            set
            {
                this.suppressBottomSpacingField = value;
            }
        }

        [XmlElement(Order = 19)]
        public CT_OnOff suppressTopSpacing
        {
            get
            {
                return this.suppressTopSpacingField;
            }
            set
            {
                this.suppressTopSpacingField = value;
            }
        }

        [XmlElement(Order = 20)]
        public CT_OnOff suppressSpacingAtTopOfPage
        {
            get
            {
                return this.suppressSpacingAtTopOfPageField;
            }
            set
            {
                this.suppressSpacingAtTopOfPageField = value;
            }
        }

        [XmlElement(Order = 21)]
        public CT_OnOff suppressTopSpacingWP
        {
            get
            {
                return this.suppressTopSpacingWPField;
            }
            set
            {
                this.suppressTopSpacingWPField = value;
            }
        }

        [XmlElement(Order = 22)]
        public CT_OnOff suppressSpBfAfterPgBrk
        {
            get
            {
                return this.suppressSpBfAfterPgBrkField;
            }
            set
            {
                this.suppressSpBfAfterPgBrkField = value;
            }
        }

        [XmlElement(Order = 23)]
        public CT_OnOff swapBordersFacingPages
        {
            get
            {
                return this.swapBordersFacingPagesField;
            }
            set
            {
                this.swapBordersFacingPagesField = value;
            }
        }

        [XmlElement(Order = 24)]
        public CT_OnOff convMailMergeEsc
        {
            get
            {
                return this.convMailMergeEscField;
            }
            set
            {
                this.convMailMergeEscField = value;
            }
        }

        [XmlElement(Order = 25)]
        public CT_OnOff truncateFontHeightsLikeWP6
        {
            get
            {
                return this.truncateFontHeightsLikeWP6Field;
            }
            set
            {
                this.truncateFontHeightsLikeWP6Field = value;
            }
        }

        [XmlElement(Order = 26)]
        public CT_OnOff mwSmallCaps
        {
            get
            {
                return this.mwSmallCapsField;
            }
            set
            {
                this.mwSmallCapsField = value;
            }
        }

        [XmlElement(Order = 27)]
        public CT_OnOff usePrinterMetrics
        {
            get
            {
                return this.usePrinterMetricsField;
            }
            set
            {
                this.usePrinterMetricsField = value;
            }
        }

        [XmlElement(Order = 28)]
        public CT_OnOff doNotSuppressParagraphBorders
        {
            get
            {
                return this.doNotSuppressParagraphBordersField;
            }
            set
            {
                this.doNotSuppressParagraphBordersField = value;
            }
        }

        [XmlElement(Order = 29)]
        public CT_OnOff wrapTrailSpaces
        {
            get
            {
                return this.wrapTrailSpacesField;
            }
            set
            {
                this.wrapTrailSpacesField = value;
            }
        }

        [XmlElement(Order = 30)]
        public CT_OnOff footnoteLayoutLikeWW8
        {
            get
            {
                return this.footnoteLayoutLikeWW8Field;
            }
            set
            {
                this.footnoteLayoutLikeWW8Field = value;
            }
        }

        [XmlElement(Order = 31)]
        public CT_OnOff shapeLayoutLikeWW8
        {
            get
            {
                return this.shapeLayoutLikeWW8Field;
            }
            set
            {
                this.shapeLayoutLikeWW8Field = value;
            }
        }

        [XmlElement(Order = 32)]
        public CT_OnOff alignTablesRowByRow
        {
            get
            {
                return this.alignTablesRowByRowField;
            }
            set
            {
                this.alignTablesRowByRowField = value;
            }
        }

        [XmlElement(Order = 33)]
        public CT_OnOff forgetLastTabAlignment
        {
            get
            {
                return this.forgetLastTabAlignmentField;
            }
            set
            {
                this.forgetLastTabAlignmentField = value;
            }
        }

        [XmlElement(Order = 34)]
        public CT_OnOff adjustLineHeightInTable
        {
            get
            {
                return this.adjustLineHeightInTableField;
            }
            set
            {
                this.adjustLineHeightInTableField = value;
            }
        }

        [XmlElement(Order = 35)]
        public CT_OnOff autoSpaceLikeWord95
        {
            get
            {
                return this.autoSpaceLikeWord95Field;
            }
            set
            {
                this.autoSpaceLikeWord95Field = value;
            }
        }

        [XmlElement(Order = 36)]
        public CT_OnOff noSpaceRaiseLower
        {
            get
            {
                return this.noSpaceRaiseLowerField;
            }
            set
            {
                this.noSpaceRaiseLowerField = value;
            }
        }

        [XmlElement(Order = 37)]
        public CT_OnOff doNotUseHTMLParagraphAutoSpacing
        {
            get
            {
                return this.doNotUseHTMLParagraphAutoSpacingField;
            }
            set
            {
                this.doNotUseHTMLParagraphAutoSpacingField = value;
            }
        }

        [XmlElement(Order = 38)]
        public CT_OnOff layoutRawTableWidth
        {
            get
            {
                return this.layoutRawTableWidthField;
            }
            set
            {
                this.layoutRawTableWidthField = value;
            }
        }

        [XmlElement(Order = 39)]
        public CT_OnOff layoutTableRowsApart
        {
            get
            {
                return this.layoutTableRowsApartField;
            }
            set
            {
                this.layoutTableRowsApartField = value;
            }
        }

        [XmlElement(Order = 40)]
        public CT_OnOff useWord97LineBreakRules
        {
            get
            {
                return this.useWord97LineBreakRulesField;
            }
            set
            {
                this.useWord97LineBreakRulesField = value;
            }
        }

        [XmlElement(Order = 41)]
        public CT_OnOff doNotBreakWrappedTables
        {
            get
            {
                return this.doNotBreakWrappedTablesField;
            }
            set
            {
                this.doNotBreakWrappedTablesField = value;
            }
        }

        [XmlElement(Order = 42)]
        public CT_OnOff doNotSnapToGridInCell
        {
            get
            {
                return this.doNotSnapToGridInCellField;
            }
            set
            {
                this.doNotSnapToGridInCellField = value;
            }
        }

        [XmlElement(Order = 43)]
        public CT_OnOff selectFldWithFirstOrLastChar
        {
            get
            {
                return this.selectFldWithFirstOrLastCharField;
            }
            set
            {
                this.selectFldWithFirstOrLastCharField = value;
            }
        }

        [XmlElement(Order = 44)]
        public CT_OnOff applyBreakingRules
        {
            get
            {
                return this.applyBreakingRulesField;
            }
            set
            {
                this.applyBreakingRulesField = value;
            }
        }

        [XmlElement(Order = 45)]
        public CT_OnOff doNotWrapTextWithPunct
        {
            get
            {
                return this.doNotWrapTextWithPunctField;
            }
            set
            {
                this.doNotWrapTextWithPunctField = value;
            }
        }

        [XmlElement(Order = 46)]
        public CT_OnOff doNotUseEastAsianBreakRules
        {
            get
            {
                return this.doNotUseEastAsianBreakRulesField;
            }
            set
            {
                this.doNotUseEastAsianBreakRulesField = value;
            }
        }

        [XmlElement(Order = 47)]
        public CT_OnOff useWord2002TableStyleRules
        {
            get
            {
                return this.useWord2002TableStyleRulesField;
            }
            set
            {
                this.useWord2002TableStyleRulesField = value;
            }
        }

        [XmlElement(Order = 48)]
        public CT_OnOff growAutofit
        {
            get
            {
                return this.growAutofitField;
            }
            set
            {
                this.growAutofitField = value;
            }
        }

        [XmlElement(Order = 49)]
        public CT_OnOff useFELayout
        {
            get
            {
                return this.useFELayoutField;
            }
            set
            {
                this.useFELayoutField = value;
            }
        }

        [XmlElement(Order = 50)]
        public CT_OnOff useNormalStyleForList
        {
            get
            {
                return this.useNormalStyleForListField;
            }
            set
            {
                this.useNormalStyleForListField = value;
            }
        }

        [XmlElement(Order = 51)]
        public CT_OnOff doNotUseIndentAsNumberingTabStop
        {
            get
            {
                return this.doNotUseIndentAsNumberingTabStopField;
            }
            set
            {
                this.doNotUseIndentAsNumberingTabStopField = value;
            }
        }

        [XmlElement(Order = 52)]
        public CT_OnOff useAltKinsokuLineBreakRules
        {
            get
            {
                return this.useAltKinsokuLineBreakRulesField;
            }
            set
            {
                this.useAltKinsokuLineBreakRulesField = value;
            }
        }

        [XmlElement(Order = 53)]
        public CT_OnOff allowSpaceOfSameStyleInTable
        {
            get
            {
                return this.allowSpaceOfSameStyleInTableField;
            }
            set
            {
                this.allowSpaceOfSameStyleInTableField = value;
            }
        }

        [XmlElement(Order = 54)]
        public CT_OnOff doNotSuppressIndentation
        {
            get
            {
                return this.doNotSuppressIndentationField;
            }
            set
            {
                this.doNotSuppressIndentationField = value;
            }
        }

        [XmlElement(Order = 55)]
        public CT_OnOff doNotAutofitConstrainedTables
        {
            get
            {
                return this.doNotAutofitConstrainedTablesField;
            }
            set
            {
                this.doNotAutofitConstrainedTablesField = value;
            }
        }

        [XmlElement(Order = 56)]
        public CT_OnOff autofitToFirstFixedWidthCell
        {
            get
            {
                return this.autofitToFirstFixedWidthCellField;
            }
            set
            {
                this.autofitToFirstFixedWidthCellField = value;
            }
        }

        [XmlElement(Order = 57)]
        public CT_OnOff underlineTabInNumList
        {
            get
            {
                return this.underlineTabInNumListField;
            }
            set
            {
                this.underlineTabInNumListField = value;
            }
        }

        [XmlElement(Order = 58)]
        public CT_OnOff displayHangulFixedWidth
        {
            get
            {
                return this.displayHangulFixedWidthField;
            }
            set
            {
                this.displayHangulFixedWidthField = value;
            }
        }

        [XmlElement(Order = 59)]
        public CT_OnOff splitPgBreakAndParaMark
        {
            get
            {
                return this.splitPgBreakAndParaMarkField;
            }
            set
            {
                this.splitPgBreakAndParaMarkField = value;
            }
        }

        [XmlElement(Order = 60)]
        public CT_OnOff doNotVertAlignCellWithSp
        {
            get
            {
                return this.doNotVertAlignCellWithSpField;
            }
            set
            {
                this.doNotVertAlignCellWithSpField = value;
            }
        }

        [XmlElement(Order = 61)]
        public CT_OnOff doNotBreakConstrainedForcedTable
        {
            get
            {
                return this.doNotBreakConstrainedForcedTableField;
            }
            set
            {
                this.doNotBreakConstrainedForcedTableField = value;
            }
        }

        [XmlElement(Order = 62)]
        public CT_OnOff doNotVertAlignInTxbx
        {
            get
            {
                return this.doNotVertAlignInTxbxField;
            }
            set
            {
                this.doNotVertAlignInTxbxField = value;
            }
        }

        [XmlElement(Order = 63)]
        public CT_OnOff useAnsiKerningPairs
        {
            get
            {
                return this.useAnsiKerningPairsField;
            }
            set
            {
                this.useAnsiKerningPairsField = value;
            }
        }

        [XmlElement(Order = 64)]
        public CT_OnOff cachedColBalance
        {
            get
            {
                return this.cachedColBalanceField;
            }
            set
            {
                this.cachedColBalanceField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocVar
    {

        private string nameField;

        private string valField;
        public static CT_DocVar Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_DocVar ctObj = new CT_DocVar();
            ctObj.name = XmlHelper.ReadString(node.Attributes["w:name"]);
            ctObj.val = XmlHelper.ReadString(node.Attributes["w:val"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:name", this.name);
            XmlHelper.WriteAttribute(sw, "w:val", this.val);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocRsids
    {

        private CT_LongHexNumber rsidRootField;

        private List<CT_LongHexNumber> rsidField;

        public CT_DocRsids()
        {
            //this.rsidField = new List<CT_LongHexNumber>();
            //this.rsidRootField = new CT_LongHexNumber();
        }
        public static CT_DocRsids Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_DocRsids ctObj = new CT_DocRsids();
            ctObj.rsid = new List<CT_LongHexNumber>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "rsidRoot")
                    ctObj.rsidRoot = CT_LongHexNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "rsid")
                    ctObj.rsid.Add(CT_LongHexNumber.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.rsidRoot != null)
                this.rsidRoot.Write(sw, "rsidRoot");
            if (this.rsid != null)
            {
                foreach (CT_LongHexNumber x in this.rsid)
                {
                    x.Write(sw, "rsid");
                }
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_LongHexNumber rsidRoot
        {
            get
            {
                return this.rsidRootField;
            }
            set
            {
                this.rsidRootField = value;
            }
        }

        [XmlElement("rsid", Order = 1)]
        public List<CT_LongHexNumber> rsid
        {
            get
            {
                return this.rsidField;
            }
            set
            {
                this.rsidField = value;
            }
        }
    }





    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_ColorSchemeMapping
    {

        private ST_ColorSchemeIndex bg1Field;

        private bool bg1FieldSpecified;

        private ST_ColorSchemeIndex t1Field;

        private bool t1FieldSpecified;

        private ST_ColorSchemeIndex bg2Field;

        private bool bg2FieldSpecified;

        private ST_ColorSchemeIndex t2Field;

        private bool t2FieldSpecified;

        private ST_ColorSchemeIndex accent1Field;

        private bool accent1FieldSpecified;

        private ST_ColorSchemeIndex accent2Field;

        private bool accent2FieldSpecified;

        private ST_ColorSchemeIndex accent3Field;

        private bool accent3FieldSpecified;

        private ST_ColorSchemeIndex accent4Field;

        private bool accent4FieldSpecified;

        private ST_ColorSchemeIndex accent5Field;

        private bool accent5FieldSpecified;

        private ST_ColorSchemeIndex accent6Field;

        private bool accent6FieldSpecified;

        private ST_ColorSchemeIndex hyperlinkField;

        private bool hyperlinkFieldSpecified;

        private ST_ColorSchemeIndex followedHyperlinkField;

        private bool followedHyperlinkFieldSpecified;
        public static CT_ColorSchemeMapping Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ColorSchemeMapping ctObj = new CT_ColorSchemeMapping();
            if (node.Attributes["w:bg1"] != null)
                ctObj.bg1 = (ST_ColorSchemeIndex)Enum.Parse(typeof(ST_ColorSchemeIndex), node.Attributes["w:bg1"].Value);
            if (node.Attributes["w:t1"] != null)
                ctObj.t1 = (ST_ColorSchemeIndex)Enum.Parse(typeof(ST_ColorSchemeIndex), node.Attributes["w:t1"].Value);
            if (node.Attributes["w:bg2"] != null)
                ctObj.bg2 = (ST_ColorSchemeIndex)Enum.Parse(typeof(ST_ColorSchemeIndex), node.Attributes["w:bg2"].Value);
            if (node.Attributes["w:t2"] != null)
                ctObj.t2 = (ST_ColorSchemeIndex)Enum.Parse(typeof(ST_ColorSchemeIndex), node.Attributes["w:t2"].Value);
            if (node.Attributes["w:accent1"] != null)
                ctObj.accent1 = (ST_ColorSchemeIndex)Enum.Parse(typeof(ST_ColorSchemeIndex), node.Attributes["w:accent1"].Value);
            if (node.Attributes["w:accent2"] != null)
                ctObj.accent2 = (ST_ColorSchemeIndex)Enum.Parse(typeof(ST_ColorSchemeIndex), node.Attributes["w:accent2"].Value);
            if (node.Attributes["w:accent3"] != null)
                ctObj.accent3 = (ST_ColorSchemeIndex)Enum.Parse(typeof(ST_ColorSchemeIndex), node.Attributes["w:accent3"].Value);
            if (node.Attributes["w:accent4"] != null)
                ctObj.accent4 = (ST_ColorSchemeIndex)Enum.Parse(typeof(ST_ColorSchemeIndex), node.Attributes["w:accent4"].Value);
            if (node.Attributes["w:accent5"] != null)
                ctObj.accent5 = (ST_ColorSchemeIndex)Enum.Parse(typeof(ST_ColorSchemeIndex), node.Attributes["w:accent5"].Value);
            if (node.Attributes["w:accent6"] != null)
                ctObj.accent6 = (ST_ColorSchemeIndex)Enum.Parse(typeof(ST_ColorSchemeIndex), node.Attributes["w:accent6"].Value);
            if (node.Attributes["w:hyperlink"] != null)
                ctObj.hyperlink = (ST_ColorSchemeIndex)Enum.Parse(typeof(ST_ColorSchemeIndex), node.Attributes["w:hyperlink"].Value);
            if (node.Attributes["w:followedHyperlink"] != null)
                ctObj.followedHyperlink = (ST_ColorSchemeIndex)Enum.Parse(typeof(ST_ColorSchemeIndex), node.Attributes["w:followedHyperlink"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:bg1", this.bg1.ToString());
            XmlHelper.WriteAttribute(sw, "w:t1", this.t1.ToString());
            XmlHelper.WriteAttribute(sw, "w:bg2", this.bg2.ToString());
            XmlHelper.WriteAttribute(sw, "w:t2", this.t2.ToString());
            XmlHelper.WriteAttribute(sw, "w:accent1", this.accent1.ToString());
            XmlHelper.WriteAttribute(sw, "w:accent2", this.accent2.ToString());
            XmlHelper.WriteAttribute(sw, "w:accent3", this.accent3.ToString());
            XmlHelper.WriteAttribute(sw, "w:accent4", this.accent4.ToString());
            XmlHelper.WriteAttribute(sw, "w:accent5", this.accent5.ToString());
            XmlHelper.WriteAttribute(sw, "w:accent6", this.accent6.ToString());
            XmlHelper.WriteAttribute(sw, "w:hyperlink", this.hyperlink.ToString());
            XmlHelper.WriteAttribute(sw, "w:followedHyperlink", this.followedHyperlink.ToString());
            sw.Write("/>");
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ColorSchemeIndex bg1
        {
            get
            {
                return this.bg1Field;
            }
            set
            {
                this.bg1Field = value;
            }
        }

        [XmlIgnore]
        public bool bg1Specified
        {
            get
            {
                return this.bg1FieldSpecified;
            }
            set
            {
                this.bg1FieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ColorSchemeIndex t1
        {
            get
            {
                return this.t1Field;
            }
            set
            {
                this.t1Field = value;
            }
        }

        [XmlIgnore]
        public bool t1Specified
        {
            get
            {
                return this.t1FieldSpecified;
            }
            set
            {
                this.t1FieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ColorSchemeIndex bg2
        {
            get
            {
                return this.bg2Field;
            }
            set
            {
                this.bg2Field = value;
            }
        }

        [XmlIgnore]
        public bool bg2Specified
        {
            get
            {
                return this.bg2FieldSpecified;
            }
            set
            {
                this.bg2FieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ColorSchemeIndex t2
        {
            get
            {
                return this.t2Field;
            }
            set
            {
                this.t2Field = value;
            }
        }

        [XmlIgnore]
        public bool t2Specified
        {
            get
            {
                return this.t2FieldSpecified;
            }
            set
            {
                this.t2FieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ColorSchemeIndex accent1
        {
            get
            {
                return this.accent1Field;
            }
            set
            {
                this.accent1Field = value;
            }
        }

        [XmlIgnore]
        public bool accent1Specified
        {
            get
            {
                return this.accent1FieldSpecified;
            }
            set
            {
                this.accent1FieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ColorSchemeIndex accent2
        {
            get
            {
                return this.accent2Field;
            }
            set
            {
                this.accent2Field = value;
            }
        }

        [XmlIgnore]
        public bool accent2Specified
        {
            get
            {
                return this.accent2FieldSpecified;
            }
            set
            {
                this.accent2FieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ColorSchemeIndex accent3
        {
            get
            {
                return this.accent3Field;
            }
            set
            {
                this.accent3Field = value;
            }
        }

        [XmlIgnore]
        public bool accent3Specified
        {
            get
            {
                return this.accent3FieldSpecified;
            }
            set
            {
                this.accent3FieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ColorSchemeIndex accent4
        {
            get
            {
                return this.accent4Field;
            }
            set
            {
                this.accent4Field = value;
            }
        }

        [XmlIgnore]
        public bool accent4Specified
        {
            get
            {
                return this.accent4FieldSpecified;
            }
            set
            {
                this.accent4FieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ColorSchemeIndex accent5
        {
            get
            {
                return this.accent5Field;
            }
            set
            {
                this.accent5Field = value;
            }
        }

        [XmlIgnore]
        public bool accent5Specified
        {
            get
            {
                return this.accent5FieldSpecified;
            }
            set
            {
                this.accent5FieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ColorSchemeIndex accent6
        {
            get
            {
                return this.accent6Field;
            }
            set
            {
                this.accent6Field = value;
            }
        }

        [XmlIgnore]
        public bool accent6Specified
        {
            get
            {
                return this.accent6FieldSpecified;
            }
            set
            {
                this.accent6FieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ColorSchemeIndex hyperlink
        {
            get
            {
                return this.hyperlinkField;
            }
            set
            {
                this.hyperlinkField = value;
            }
        }

        [XmlIgnore]
        public bool hyperlinkSpecified
        {
            get
            {
                return this.hyperlinkFieldSpecified;
            }
            set
            {
                this.hyperlinkFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ColorSchemeIndex followedHyperlink
        {
            get
            {
                return this.followedHyperlinkField;
            }
            set
            {
                this.followedHyperlinkField = value;
            }
        }

        [XmlIgnore]
        public bool followedHyperlinkSpecified
        {
            get
            {
                return this.followedHyperlinkFieldSpecified;
            }
            set
            {
                this.followedHyperlinkFieldSpecified = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_ColorSchemeIndex
    {

    
        dark1,

    
        light1,

    
        dark2,

    
        light2,

    
        accent1,

    
        accent2,

    
        accent3,

    
        accent4,

    
        accent5,

    
        accent6,

    
        hyperlink,

    
        followedHyperlink,
    }




    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Captions
    {

        private List<CT_Caption> captionField;

        private List<CT_AutoCaption> autoCaptionsField;

        public CT_Captions()
        {
            //this.autoCaptionsField = new List<CT_AutoCaption>();
            //this.captionField = new List<CT_Caption>();
        }
        public static CT_Captions Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Captions ctObj = new CT_Captions();
            ctObj.caption = new List<CT_Caption>();
            ctObj.autoCaptions = new List<CT_AutoCaption>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "caption")
                    ctObj.caption.Add(CT_Caption.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "autoCaptions")
                    ctObj.autoCaptions.Add(CT_AutoCaption.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.caption != null)
            {
                foreach (CT_Caption x in this.caption)
                {
                    x.Write(sw, "caption");
                }
            }
            if (this.autoCaptions != null)
            {
                foreach (CT_AutoCaption x in this.autoCaptions)
                {
                    x.Write(sw, "autoCaptions");
                }
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement("caption", Order = 0)]
        public List<CT_Caption> caption
        {
            get
            {
                return this.captionField;
            }
            set
            {
                this.captionField = value;
            }
        }

        [XmlArray(Order = 1)]
        [XmlArrayItem("autoCaption", IsNullable = false)]
        public List<CT_AutoCaption> autoCaptions
        {
            get
            {
                return this.autoCaptionsField;
            }
            set
            {
                this.autoCaptionsField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Caption
    {

        private string nameField;

        private ST_CaptionPos posField;

        private bool posFieldSpecified;

        private ST_OnOff chapNumField;

        private bool chapNumFieldSpecified;

        private string headingField;

        private ST_OnOff noLabelField;

        private bool noLabelFieldSpecified;

        private ST_NumberFormat numFmtField;

        private bool numFmtFieldSpecified;

        private ST_ChapterSep sepField;

        private bool sepFieldSpecified;
        public static CT_Caption Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Caption ctObj = new CT_Caption();
            ctObj.name = XmlHelper.ReadString(node.Attributes["w:name"]);
            if (node.Attributes["w:pos"] != null)
                ctObj.pos = (ST_CaptionPos)Enum.Parse(typeof(ST_CaptionPos), node.Attributes["w:pos"].Value);
            if (node.Attributes["w:chapNum"] != null)
                ctObj.chapNum = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:chapNum"].Value);
            ctObj.heading = XmlHelper.ReadString(node.Attributes["w:heading"]);
            if (node.Attributes["w:noLabel"] != null)
                ctObj.noLabel = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:noLabel"].Value);
            if (node.Attributes["w:numFmt"] != null)
                ctObj.numFmt = (ST_NumberFormat)Enum.Parse(typeof(ST_NumberFormat), node.Attributes["w:numFmt"].Value);
            if (node.Attributes["w:sep"] != null)
                ctObj.sep = (ST_ChapterSep)Enum.Parse(typeof(ST_ChapterSep), node.Attributes["w:sep"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:name", this.name);
            XmlHelper.WriteAttribute(sw, "w:pos", this.pos.ToString());
            XmlHelper.WriteAttribute(sw, "w:chapNum", this.chapNum.ToString());
            XmlHelper.WriteAttribute(sw, "w:heading", this.heading);
            XmlHelper.WriteAttribute(sw, "w:noLabel", this.noLabel.ToString());
            XmlHelper.WriteAttribute(sw, "w:numFmt", this.numFmt.ToString());
            XmlHelper.WriteAttribute(sw, "w:sep", this.sep.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_CaptionPos pos
        {
            get
            {
                return this.posField;
            }
            set
            {
                this.posField = value;
            }
        }

        [XmlIgnore]
        public bool posSpecified
        {
            get
            {
                return this.posFieldSpecified;
            }
            set
            {
                this.posFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff chapNum
        {
            get
            {
                return this.chapNumField;
            }
            set
            {
                this.chapNumField = value;
            }
        }

        [XmlIgnore]
        public bool chapNumSpecified
        {
            get
            {
                return this.chapNumFieldSpecified;
            }
            set
            {
                this.chapNumFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string heading
        {
            get
            {
                return this.headingField;
            }
            set
            {
                this.headingField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff noLabel
        {
            get
            {
                return this.noLabelField;
            }
            set
            {
                this.noLabelField = value;
            }
        }

        [XmlIgnore]
        public bool noLabelSpecified
        {
            get
            {
                return this.noLabelFieldSpecified;
            }
            set
            {
                this.noLabelFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_NumberFormat numFmt
        {
            get
            {
                return this.numFmtField;
            }
            set
            {
                this.numFmtField = value;
            }
        }

        [XmlIgnore]
        public bool numFmtSpecified
        {
            get
            {
                return this.numFmtFieldSpecified;
            }
            set
            {
                this.numFmtFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ChapterSep sep
        {
            get
            {
                return this.sepField;
            }
            set
            {
                this.sepField = value;
            }
        }

        [XmlIgnore]
        public bool sepSpecified
        {
            get
            {
                return this.sepFieldSpecified;
            }
            set
            {
                this.sepFieldSpecified = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_CaptionPos
    {

    
        above,

    
        below,

    
        left,

    
        right,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_AutoCaption
    {

        private string nameField;

        private string captionField;
        public static CT_AutoCaption Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_AutoCaption ctObj = new CT_AutoCaption();
            ctObj.name = XmlHelper.ReadString(node.Attributes["w:name"]);
            ctObj.caption = XmlHelper.ReadString(node.Attributes["w:caption"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:name", this.name);
            XmlHelper.WriteAttribute(sw, "w:caption", this.caption);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string caption
        {
            get
            {
                return this.captionField;
            }
            set
            {
                this.captionField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_ReadingModeInkLockDown
    {

        private ST_OnOff actualPgField;

        private ulong wField;

        private ulong hField;

        private string fontSzField;
        public static CT_ReadingModeInkLockDown Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ReadingModeInkLockDown ctObj = new CT_ReadingModeInkLockDown();
            if (node.Attributes["w:actualPg"] != null)
                ctObj.actualPg = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:actualPg"].Value);
            ctObj.w = XmlHelper.ReadULong(node.Attributes["w:w"]);
            ctObj.h = XmlHelper.ReadULong(node.Attributes["w:h"]);
            ctObj.fontSz = XmlHelper.ReadString(node.Attributes["w:fontSz"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:actualPg", this.actualPg.ToString());
            XmlHelper.WriteAttribute(sw, "w:w", this.w);
            XmlHelper.WriteAttribute(sw, "w:h", this.h);
            XmlHelper.WriteAttribute(sw, "w:fontSz", this.fontSz);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff actualPg
        {
            get
            {
                return this.actualPgField;
            }
            set
            {
                this.actualPgField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong w
        {
            get
            {
                return this.wField;
            }
            set
            {
                this.wField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong h
        {
            get
            {
                return this.hField;
            }
            set
            {
                this.hField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string fontSz
        {
            get
            {
                return this.fontSzField;
            }
            set
            {
                this.fontSzField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SmartTagType
    {

        private string namespaceuriField;

        private string nameField;

        private string urlField;
        public static CT_SmartTagType Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SmartTagType ctObj = new CT_SmartTagType();
            ctObj.namespaceuri = XmlHelper.ReadString(node.Attributes["w:namespaceuri"]);
            ctObj.name = XmlHelper.ReadString(node.Attributes["w:name"]);
            ctObj.url = XmlHelper.ReadString(node.Attributes["w:url"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:namespaceuri", this.namespaceuri);
            XmlHelper.WriteAttribute(sw, "w:name", this.name);
            XmlHelper.WriteAttribute(sw, "w:url", this.url);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string namespaceuri
        {
            get
            {
                return this.namespaceuriField;
            }
            set
            {
                this.namespaceuriField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string url
        {
            get
            {
                return this.urlField;
            }
            set
            {
                this.urlField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot("webSettings", Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = false)]
    public class CT_WebSettings
    {

        private CT_Frameset framesetField;

        private CT_Divs divsField;

        private CT_String encodingField;

        private CT_OnOff optimizeForBrowserField;

        private CT_OnOff relyOnVMLField;

        private CT_OnOff allowPNGField;

        private CT_OnOff doNotRelyOnCSSField;

        private CT_OnOff doNotSaveAsSingleFileField;

        private CT_OnOff doNotOrganizeInFolderField;

        private CT_OnOff doNotUseLongFileNamesField;

        private CT_DecimalNumber pixelsPerInchField;

        private CT_TargetScreenSz targetScreenSzField;

        private CT_OnOff saveSmartTagsAsXmlField;

        public CT_WebSettings()
        {
            //this.saveSmartTagsAsXmlField = new CT_OnOff();
            //this.targetScreenSzField = new CT_TargetScreenSz();
            //this.pixelsPerInchField = new CT_DecimalNumber();
            //this.doNotUseLongFileNamesField = new CT_OnOff();
            //this.doNotOrganizeInFolderField = new CT_OnOff();
            //this.doNotSaveAsSingleFileField = new CT_OnOff();
            //this.doNotRelyOnCSSField = new CT_OnOff();
            this.allowPNGField = new CT_OnOff();
            //this.relyOnVMLField = new CT_OnOff();
            this.optimizeForBrowserField = new CT_OnOff();
            //this.encodingField = new CT_String();
            //this.divsField = new CT_Divs();
            //this.framesetField = new CT_Frameset();
        }

        [XmlElement(Order = 0)]
        public CT_Frameset frameset
        {
            get
            {
                return this.framesetField;
            }
            set
            {
                this.framesetField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_Divs divs
        {
            get
            {
                return this.divsField;
            }
            set
            {
                this.divsField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_String encoding
        {
            get
            {
                return this.encodingField;
            }
            set
            {
                this.encodingField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_OnOff optimizeForBrowser
        {
            get
            {
                return this.optimizeForBrowserField;
            }
            set
            {
                this.optimizeForBrowserField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_OnOff relyOnVML
        {
            get
            {
                return this.relyOnVMLField;
            }
            set
            {
                this.relyOnVMLField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_OnOff allowPNG
        {
            get
            {
                return this.allowPNGField;
            }
            set
            {
                this.allowPNGField = value;
            }
        }

        [XmlElement(Order = 6)]
        public CT_OnOff doNotRelyOnCSS
        {
            get
            {
                return this.doNotRelyOnCSSField;
            }
            set
            {
                this.doNotRelyOnCSSField = value;
            }
        }

        [XmlElement(Order = 7)]
        public CT_OnOff doNotSaveAsSingleFile
        {
            get
            {
                return this.doNotSaveAsSingleFileField;
            }
            set
            {
                this.doNotSaveAsSingleFileField = value;
            }
        }

        [XmlElement(Order = 8)]
        public CT_OnOff doNotOrganizeInFolder
        {
            get
            {
                return this.doNotOrganizeInFolderField;
            }
            set
            {
                this.doNotOrganizeInFolderField = value;
            }
        }

        [XmlElement(Order = 9)]
        public CT_OnOff doNotUseLongFileNames
        {
            get
            {
                return this.doNotUseLongFileNamesField;
            }
            set
            {
                this.doNotUseLongFileNamesField = value;
            }
        }

        [XmlElement(Order = 10)]
        public CT_DecimalNumber pixelsPerInch
        {
            get
            {
                return this.pixelsPerInchField;
            }
            set
            {
                this.pixelsPerInchField = value;
            }
        }

        [XmlElement(Order = 11)]
        public CT_TargetScreenSz targetScreenSz
        {
            get
            {
                return this.targetScreenSzField;
            }
            set
            {
                this.targetScreenSzField = value;
            }
        }

        [XmlElement(Order = 12)]
        public CT_OnOff saveSmartTagsAsXml
        {
            get
            {
                return this.saveSmartTagsAsXmlField;
            }
            set
            {
                this.saveSmartTagsAsXmlField = value;
            }
        }
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_ShapeDefaults
    {

        private System.Xml.XmlElement[] itemsField;

        public CT_ShapeDefaults()
        {
            this.itemsField = new System.Xml.XmlElement[0];
        }

        [XmlAnyElement(Namespace = "urn:schemas-microsoft-com:office:office", Order = 0)]
        public System.Xml.XmlElement[] Items
        {
            get
            {
                return this.itemsField;
            }
            set
            {
                this.itemsField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_AutoCaptions
    {

        private List<CT_AutoCaption> autoCaptionField;

        public CT_AutoCaptions()
        {
            this.autoCaptionField = new List<CT_AutoCaption>();
        }

        [XmlElement("autoCaption", Order = 0)]
        public List<CT_AutoCaption> autoCaption
        {
            get
            {
                return this.autoCaptionField;
            }
            set
            {
                this.autoCaptionField = value;
            }
        }
    }

}
