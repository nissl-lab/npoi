using System;
using System.Collections.Generic;
using System.Xml.Serialization;
//using NPOI.OpenXmlFormats.Shared;

namespace NPOI.OpenXmlFormats.Wordprocessing
{


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot("settings", Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = false)]
    public class CT_Settings
    {

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

        public CT_MailMerge()
        {
            this.odsoField = new CT_Odso();
            this.checkErrorsField = new CT_DecimalNumber();
            this.activeRecordField = new CT_DecimalNumber();
            this.viewMergedDataField = new CT_OnOff();
            this.mailAsAttachmentField = new CT_OnOff();
            this.mailSubjectField = new CT_String();
            this.addressFieldNameField = new CT_String();
            this.destinationField = new CT_MailMergeDest();
            this.doNotSuppressBlankLinesField = new CT_OnOff();
            this.headerSourceField = new CT_Rel();
            this.dataSourceField = new CT_Rel();
            this.queryField = new CT_String();
            this.connectStringField = new CT_String();
            this.dataTypeField = new CT_MailMergeDataType();
            this.linkToQueryField = new CT_OnOff();
            this.mainDocumentTypeField = new CT_MailMergeDocType();
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
            this.recipientDataField = new List<CT_Rel>();
            this.fieldMapDataField = new List<CT_OdsoFieldMapData>();
            this.fHdrField = new CT_OnOff();
            this.typeField = new CT_MailMergeSourceType();
            this.colDelimField = new CT_DecimalNumber();
            this.srcField = new CT_Rel();
            this.tableField = new CT_String();
            this.udlField = new CT_String();
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

        public CT_OdsoFieldMapData()
        {
            this.dynamicAddressField = new CT_OnOff();
            this.lidField = new CT_Lang();
            this.columnField = new CT_DecimalNumber();
            this.mappedNameField = new CT_String();
            this.nameField = new CT_String();
            this.typeField = new CT_MailMergeOdsoFMDFieldType();
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

        private byte[] hashField;

        private byte[] saltField;

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
            this.rsidField = new List<CT_LongHexNumber>();
            this.rsidRootField = new CT_LongHexNumber();
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
            this.autoCaptionsField = new List<CT_AutoCaption>();
            this.captionField = new List<CT_Caption>();
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
