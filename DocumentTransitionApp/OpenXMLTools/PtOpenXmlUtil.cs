/***************************************************************************

Copyright (c) Microsoft Corporation 2012-2015.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Published at http://OpenXmlDeveloper.org
Resource Center and Documentation: http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx

Developer: Eric White
Blog: http://www.ericwhite.com
Twitter: @EricWhiteDev
Email: eric@ericwhite.com

Version: 3.1.10
 * Add PtOpenXml.ListItemRun

Version: 2.7.03
 * Enhancements to support RTL

Version: 2.6.00
 * Enhancements to support HtmlConverter.cs

***************************************************************************/

using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace OpenXMLTools
{
    public static class PtOpenXmlExtensions
    {
        public static void PutXDocument(this OpenXmlPart part)
        {
            XDocument partXDocument = part.GetXDocument();
            if (partXDocument != null)
            {
                using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
                using (XmlWriter partXmlWriter = XmlWriter.Create(partStream))
                    partXDocument.Save(partXmlWriter);
            }
        }

        public static void PutXDocument(this OpenXmlPart part, XDocument document)
        {
            using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
            using (XmlWriter partXmlWriter = XmlWriter.Create(partStream))
                document.Save(partXmlWriter);
            part.RemoveAnnotations<XDocument>();
            part.AddAnnotation(document);
        }

        public static XDocument GetXDocument(this OpenXmlPart part)
        {
            try
            {
                XDocument partXDocument = part.Annotation<XDocument>();
                if (partXDocument != null)
                    return partXDocument;
                using (Stream partStream = part.GetStream())
                {
                    if (partStream.Length == 0)
                    {
                        partXDocument = new XDocument();
                        partXDocument.Declaration = new XDeclaration("1.0", "UTF-8", "yes");
                    }
                    else
                        using (XmlReader partXmlReader = XmlReader.Create(partStream))
                            partXDocument = XDocument.Load(partXmlReader);
                }
                part.AddAnnotation(partXDocument);
                return partXDocument;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public static List<OpenXmlPart> GetAllParts(this WordprocessingDocument doc)
        {
            // use the following so that parts are processed only once
            HashSet<OpenXmlPart> partList = new HashSet<OpenXmlPart>();
            foreach (IdPartPair p in doc.Parts)
                AddPart(partList, p.OpenXmlPart);
            return partList.OrderBy(p => p.ContentType).ThenBy(p => p.Uri.ToString()).ToList();
        }

        public static List<OpenXmlPart> GetAllParts(this SpreadsheetDocument doc)
        {
            // use the following so that parts are processed only once
            HashSet<OpenXmlPart> partList = new HashSet<OpenXmlPart>();
            foreach (IdPartPair p in doc.Parts)
                AddPart(partList, p.OpenXmlPart);
            return partList.OrderBy(p => p.ContentType).ThenBy(p => p.Uri.ToString()).ToList();
        }

        public static List<OpenXmlPart> GetAllParts(this PresentationDocument doc)
        {
            // use the following so that parts are processed only once
            HashSet<OpenXmlPart> partList = new HashSet<OpenXmlPart>();
            foreach (IdPartPair p in doc.Parts)
                AddPart(partList, p.OpenXmlPart);
            return partList.OrderBy(p => p.ContentType).ThenBy(p => p.Uri.ToString()).ToList();
        }

        private static void AddPart(HashSet<OpenXmlPart> partList, OpenXmlPart part)
        {
            if (partList.Contains(part))
                return;
            partList.Add(part);
            foreach (IdPartPair p in part.Parts)
                AddPart(partList, p.OpenXmlPart);
        }

        private static XmlNamespaceManager GetManagerFromXDocument(XDocument xDocument)
        {
            XmlReader reader = xDocument.CreateReader();
            XDocument newXDoc = XDocument.Load(reader);
            XElement rootElement = xDocument.Elements().FirstOrDefault();
            rootElement.ReplaceWith(newXDoc.Root);
            XmlNameTable nameTable = reader.NameTable;
            XmlNamespaceManager namespaceManager = new XmlNamespaceManager(nameTable);
            return namespaceManager;
        }
    }

    public static class W
    {
        public static XNamespace w =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static XName abstractNum = w + "abstractNum";
        public static XName abstractNumId = w + "abstractNumId";
        public static XName accent1 = w + "accent1";
        public static XName accent2 = w + "accent2";
        public static XName accent3 = w + "accent3";
        public static XName accent4 = w + "accent4";
        public static XName accent5 = w + "accent5";
        public static XName accent6 = w + "accent6";
        public static XName activeRecord = w + "activeRecord";
        public static XName activeWritingStyle = w + "activeWritingStyle";
        public static XName actualPg = w + "actualPg";
        public static XName addressFieldName = w + "addressFieldName";
        public static XName adjustLineHeightInTable = w + "adjustLineHeightInTable";
        public static XName adjustRightInd = w + "adjustRightInd";
        public static XName after = w + "after";
        public static XName afterAutospacing = w + "afterAutospacing";
        public static XName afterLines = w + "afterLines";
        public static XName algIdExt = w + "algIdExt";
        public static XName algIdExtSource = w + "algIdExtSource";
        public static XName alias = w + "alias";
        public static XName aliases = w + "aliases";
        public static XName alignBordersAndEdges = w + "alignBordersAndEdges";
        public static XName alignment = w + "alignment";
        public static XName alignTablesRowByRow = w + "alignTablesRowByRow";
        public static XName allowPNG = w + "allowPNG";
        public static XName allowSpaceOfSameStyleInTable = w + "allowSpaceOfSameStyleInTable";
        public static XName altChunk = w + "altChunk";
        public static XName altChunkPr = w + "altChunkPr";
        public static XName altName = w + "altName";
        public static XName alwaysMergeEmptyNamespace = w + "alwaysMergeEmptyNamespace";
        public static XName alwaysShowPlaceholderText = w + "alwaysShowPlaceholderText";
        public static XName anchor = w + "anchor";
        public static XName anchorLock = w + "anchorLock";
        public static XName annotationRef = w + "annotationRef";
        public static XName applyBreakingRules = w + "applyBreakingRules";
        public static XName appName = w + "appName";
        public static XName ascii = w + "ascii";
        public static XName asciiTheme = w + "asciiTheme";
        public static XName attachedSchema = w + "attachedSchema";
        public static XName attachedTemplate = w + "attachedTemplate";
        public static XName attr = w + "attr";
        public static XName author = w + "author";
        public static XName autofitToFirstFixedWidthCell = w + "autofitToFirstFixedWidthCell";
        public static XName autoFormatOverride = w + "autoFormatOverride";
        public static XName autoHyphenation = w + "autoHyphenation";
        public static XName autoRedefine = w + "autoRedefine";
        public static XName autoSpaceDE = w + "autoSpaceDE";
        public static XName autoSpaceDN = w + "autoSpaceDN";
        public static XName autoSpaceLikeWord95 = w + "autoSpaceLikeWord95";
        public static XName b = w + "b";
        public static XName background = w + "background";
        public static XName balanceSingleByteDoubleByteWidth = w + "balanceSingleByteDoubleByteWidth";
        public static XName bar = w + "bar";
        public static XName basedOn = w + "basedOn";
        public static XName bCs = w + "bCs";
        public static XName bdr = w + "bdr";
        public static XName before = w + "before";
        public static XName beforeAutospacing = w + "beforeAutospacing";
        public static XName beforeLines = w + "beforeLines";
        public static XName behavior = w + "behavior";
        public static XName behaviors = w + "behaviors";
        public static XName between = w + "between";
        public static XName bg1 = w + "bg1";
        public static XName bg2 = w + "bg2";
        public static XName bibliography = w + "bibliography";
        public static XName bidi = w + "bidi";
        public static XName bidiVisual = w + "bidiVisual";
        public static XName blockQuote = w + "blockQuote";
        public static XName body = w + "body";
        public static XName bodyDiv = w + "bodyDiv";
        public static XName bookFoldPrinting = w + "bookFoldPrinting";
        public static XName bookFoldPrintingSheets = w + "bookFoldPrintingSheets";
        public static XName bookFoldRevPrinting = w + "bookFoldRevPrinting";
        public static XName bookmarkEnd = w + "bookmarkEnd";
        public static XName bookmarkStart = w + "bookmarkStart";
        public static XName bordersDoNotSurroundFooter = w + "bordersDoNotSurroundFooter";
        public static XName bordersDoNotSurroundHeader = w + "bordersDoNotSurroundHeader";
        public static XName bottom = w + "bottom";
        public static XName bottomFromText = w + "bottomFromText";
        public static XName br = w + "br";
        public static XName cachedColBalance = w + "cachedColBalance";
        public static XName calcOnExit = w + "calcOnExit";
        public static XName calendar = w + "calendar";
        public static XName cantSplit = w + "cantSplit";
        public static XName caps = w + "caps";
        public static XName category = w + "category";
        public static XName cellDel = w + "cellDel";
        public static XName cellIns = w + "cellIns";
        public static XName cellMerge = w + "cellMerge";
        public static XName chapSep = w + "chapSep";
        public static XName chapStyle = w + "chapStyle";
        public static XName _char = w + "char";
        public static XName characterSpacingControl = w + "characterSpacingControl";
        public static XName charset = w + "charset";
        public static XName charSpace = w + "charSpace";
        public static XName checkBox = w + "checkBox";
        public static XName _checked = w + "checked";
        public static XName checkErrors = w + "checkErrors";
        public static XName checkStyle = w + "checkStyle";
        public static XName citation = w + "citation";
        public static XName clear = w + "clear";
        public static XName clickAndTypeStyle = w + "clickAndTypeStyle";
        public static XName clrSchemeMapping = w + "clrSchemeMapping";
        public static XName cnfStyle = w + "cnfStyle";
        public static XName code = w + "code";
        public static XName col = w + "col";
        public static XName colDelim = w + "colDelim";
        public static XName colFirst = w + "colFirst";
        public static XName colLast = w + "colLast";
        public static XName color = w + "color";
        public static XName cols = w + "cols";
        public static XName column = w + "column";
        public static XName combine = w + "combine";
        public static XName combineBrackets = w + "combineBrackets";
        public static XName comboBox = w + "comboBox";
        public static XName comment = w + "comment";
        public static XName commentRangeEnd = w + "commentRangeEnd";
        public static XName commentRangeStart = w + "commentRangeStart";
        public static XName commentReference = w + "commentReference";
        public static XName comments = w + "comments";
        public static XName compat = w + "compat";
        public static XName compatSetting = w + "compatSetting";
        public static XName connectString = w + "connectString";
        public static XName consecutiveHyphenLimit = w + "consecutiveHyphenLimit";
        public static XName contextualSpacing = w + "contextualSpacing";
        public static XName continuationSeparator = w + "continuationSeparator";
        public static XName control = w + "control";
        public static XName convMailMergeEsc = w + "convMailMergeEsc";
        public static XName count = w + "count";
        public static XName countBy = w + "countBy";
        public static XName cr = w + "cr";
        public static XName cryptAlgorithmClass = w + "cryptAlgorithmClass";
        public static XName cryptAlgorithmSid = w + "cryptAlgorithmSid";
        public static XName cryptAlgorithmType = w + "cryptAlgorithmType";
        public static XName cryptProvider = w + "cryptProvider";
        public static XName cryptProviderType = w + "cryptProviderType";
        public static XName cryptProviderTypeExt = w + "cryptProviderTypeExt";
        public static XName cryptProviderTypeExtSource = w + "cryptProviderTypeExtSource";
        public static XName cryptSpinCount = w + "cryptSpinCount";
        public static XName cs = w + "cs";
        public static XName csb0 = w + "csb0";
        public static XName csb1 = w + "csb1";
        public static XName cstheme = w + "cstheme";
        public static XName customMarkFollows = w + "customMarkFollows";
        public static XName customStyle = w + "customStyle";
        public static XName customXml = w + "customXml";
        public static XName customXmlDelRangeEnd = w + "customXmlDelRangeEnd";
        public static XName customXmlDelRangeStart = w + "customXmlDelRangeStart";
        public static XName customXmlInsRangeEnd = w + "customXmlInsRangeEnd";
        public static XName customXmlInsRangeStart = w + "customXmlInsRangeStart";
        public static XName customXmlMoveFromRangeEnd = w + "customXmlMoveFromRangeEnd";
        public static XName customXmlMoveFromRangeStart = w + "customXmlMoveFromRangeStart";
        public static XName customXmlMoveToRangeEnd = w + "customXmlMoveToRangeEnd";
        public static XName customXmlMoveToRangeStart = w + "customXmlMoveToRangeStart";
        public static XName customXmlPr = w + "customXmlPr";
        public static XName dataBinding = w + "dataBinding";
        public static XName dataSource = w + "dataSource";
        public static XName dataType = w + "dataType";
        public static XName date = w + "date";
        public static XName dateFormat = w + "dateFormat";
        public static XName dayLong = w + "dayLong";
        public static XName dayShort = w + "dayShort";
        public static XName ddList = w + "ddList";
        public static XName decimalSymbol = w + "decimalSymbol";
        public static XName _default = w + "default";
        public static XName defaultTableStyle = w + "defaultTableStyle";
        public static XName defaultTabStop = w + "defaultTabStop";
        public static XName defLockedState = w + "defLockedState";
        public static XName defQFormat = w + "defQFormat";
        public static XName defSemiHidden = w + "defSemiHidden";
        public static XName defUIPriority = w + "defUIPriority";
        public static XName defUnhideWhenUsed = w + "defUnhideWhenUsed";
        public static XName del = w + "del";
        public static XName delInstrText = w + "delInstrText";
        public static XName delText = w + "delText";
        public static XName description = w + "description";
        public static XName destination = w + "destination";
        public static XName dir = w + "dir";
        public static XName dirty = w + "dirty";
        public static XName displacedByCustomXml = w + "displacedByCustomXml";
        public static XName display = w + "display";
        public static XName displayBackgroundShape = w + "displayBackgroundShape";
        public static XName displayHangulFixedWidth = w + "displayHangulFixedWidth";
        public static XName displayHorizontalDrawingGridEvery = w + "displayHorizontalDrawingGridEvery";
        public static XName displayText = w + "displayText";
        public static XName displayVerticalDrawingGridEvery = w + "displayVerticalDrawingGridEvery";
        public static XName distance = w + "distance";
        public static XName div = w + "div";
        public static XName divBdr = w + "divBdr";
        public static XName divId = w + "divId";
        public static XName divs = w + "divs";
        public static XName divsChild = w + "divsChild";
        public static XName dllVersion = w + "dllVersion";
        public static XName docDefaults = w + "docDefaults";
        public static XName docGrid = w + "docGrid";
        public static XName docLocation = w + "docLocation";
        public static XName docPart = w + "docPart";
        public static XName docPartBody = w + "docPartBody";
        public static XName docPartCategory = w + "docPartCategory";
        public static XName docPartGallery = w + "docPartGallery";
        public static XName docPartList = w + "docPartList";
        public static XName docPartObj = w + "docPartObj";
        public static XName docPartPr = w + "docPartPr";
        public static XName docParts = w + "docParts";
        public static XName docPartUnique = w + "docPartUnique";
        public static XName document = w + "document";
        public static XName documentProtection = w + "documentProtection";
        public static XName documentType = w + "documentType";
        public static XName docVar = w + "docVar";
        public static XName docVars = w + "docVars";
        public static XName doNotAutoCompressPictures = w + "doNotAutoCompressPictures";
        public static XName doNotAutofitConstrainedTables = w + "doNotAutofitConstrainedTables";
        public static XName doNotBreakConstrainedForcedTable = w + "doNotBreakConstrainedForcedTable";
        public static XName doNotBreakWrappedTables = w + "doNotBreakWrappedTables";
        public static XName doNotDemarcateInvalidXml = w + "doNotDemarcateInvalidXml";
        public static XName doNotDisplayPageBoundaries = w + "doNotDisplayPageBoundaries";
        public static XName doNotEmbedSmartTags = w + "doNotEmbedSmartTags";
        public static XName doNotExpandShiftReturn = w + "doNotExpandShiftReturn";
        public static XName doNotHyphenateCaps = w + "doNotHyphenateCaps";
        public static XName doNotIncludeSubdocsInStats = w + "doNotIncludeSubdocsInStats";
        public static XName doNotLeaveBackslashAlone = w + "doNotLeaveBackslashAlone";
        public static XName doNotOrganizeInFolder = w + "doNotOrganizeInFolder";
        public static XName doNotRelyOnCSS = w + "doNotRelyOnCSS";
        public static XName doNotSaveAsSingleFile = w + "doNotSaveAsSingleFile";
        public static XName doNotShadeFormData = w + "doNotShadeFormData";
        public static XName doNotSnapToGridInCell = w + "doNotSnapToGridInCell";
        public static XName doNotSuppressBlankLines = w + "doNotSuppressBlankLines";
        public static XName doNotSuppressIndentation = w + "doNotSuppressIndentation";
        public static XName doNotSuppressParagraphBorders = w + "doNotSuppressParagraphBorders";
        public static XName doNotTrackFormatting = w + "doNotTrackFormatting";
        public static XName doNotTrackMoves = w + "doNotTrackMoves";
        public static XName doNotUseEastAsianBreakRules = w + "doNotUseEastAsianBreakRules";
        public static XName doNotUseHTMLParagraphAutoSpacing = w + "doNotUseHTMLParagraphAutoSpacing";
        public static XName doNotUseIndentAsNumberingTabStop = w + "doNotUseIndentAsNumberingTabStop";
        public static XName doNotUseLongFileNames = w + "doNotUseLongFileNames";
        public static XName doNotUseMarginsForDrawingGridOrigin = w + "doNotUseMarginsForDrawingGridOrigin";
        public static XName doNotValidateAgainstSchema = w + "doNotValidateAgainstSchema";
        public static XName doNotVertAlignCellWithSp = w + "doNotVertAlignCellWithSp";
        public static XName doNotVertAlignInTxbx = w + "doNotVertAlignInTxbx";
        public static XName doNotWrapTextWithPunct = w + "doNotWrapTextWithPunct";
        public static XName drawing = w + "drawing";
        public static XName drawingGridHorizontalOrigin = w + "drawingGridHorizontalOrigin";
        public static XName drawingGridHorizontalSpacing = w + "drawingGridHorizontalSpacing";
        public static XName drawingGridVerticalOrigin = w + "drawingGridVerticalOrigin";
        public static XName drawingGridVerticalSpacing = w + "drawingGridVerticalSpacing";
        public static XName dropCap = w + "dropCap";
        public static XName dropDownList = w + "dropDownList";
        public static XName dstrike = w + "dstrike";
        public static XName dxaOrig = w + "dxaOrig";
        public static XName dyaOrig = w + "dyaOrig";
        public static XName dynamicAddress = w + "dynamicAddress";
        public static XName eastAsia = w + "eastAsia";
        public static XName eastAsianLayout = w + "eastAsianLayout";
        public static XName eastAsiaTheme = w + "eastAsiaTheme";
        public static XName ed = w + "ed";
        public static XName edGrp = w + "edGrp";
        public static XName edit = w + "edit";
        public static XName effect = w + "effect";
        public static XName element = w + "element";
        public static XName em = w + "em";
        public static XName embedBold = w + "embedBold";
        public static XName embedBoldItalic = w + "embedBoldItalic";
        public static XName embedItalic = w + "embedItalic";
        public static XName embedRegular = w + "embedRegular";
        public static XName embedSystemFonts = w + "embedSystemFonts";
        public static XName embedTrueTypeFonts = w + "embedTrueTypeFonts";
        public static XName emboss = w + "emboss";
        public static XName enabled = w + "enabled";
        public static XName encoding = w + "encoding";
        public static XName endnote = w + "endnote";
        public static XName endnotePr = w + "endnotePr";
        public static XName endnoteRef = w + "endnoteRef";
        public static XName endnoteReference = w + "endnoteReference";
        public static XName endnotes = w + "endnotes";
        public static XName enforcement = w + "enforcement";
        public static XName entryMacro = w + "entryMacro";
        public static XName equalWidth = w + "equalWidth";
        public static XName equation = w + "equation";
        public static XName evenAndOddHeaders = w + "evenAndOddHeaders";
        public static XName exitMacro = w + "exitMacro";
        public static XName family = w + "family";
        public static XName ffData = w + "ffData";
        public static XName fHdr = w + "fHdr";
        public static XName fieldMapData = w + "fieldMapData";
        public static XName fill = w + "fill";
        public static XName first = w + "first";
        public static XName firstColumn = w + "firstColumn";
        public static XName firstLine = w + "firstLine";
        public static XName firstLineChars = w + "firstLineChars";
        public static XName firstRow = w + "firstRow";
        public static XName fitText = w + "fitText";
        public static XName flatBorders = w + "flatBorders";
        public static XName fldChar = w + "fldChar";
        public static XName fldCharType = w + "fldCharType";
        public static XName fldData = w + "fldData";
        public static XName fldLock = w + "fldLock";
        public static XName fldSimple = w + "fldSimple";
        public static XName fmt = w + "fmt";
        public static XName followedHyperlink = w + "followedHyperlink";
        public static XName font = w + "font";
        public static XName fontKey = w + "fontKey";
        public static XName fonts = w + "fonts";
        public static XName fontSz = w + "fontSz";
        public static XName footer = w + "footer";
        public static XName footerReference = w + "footerReference";
        public static XName footnote = w + "footnote";
        public static XName footnoteLayoutLikeWW8 = w + "footnoteLayoutLikeWW8";
        public static XName footnotePr = w + "footnotePr";
        public static XName footnoteRef = w + "footnoteRef";
        public static XName footnoteReference = w + "footnoteReference";
        public static XName footnotes = w + "footnotes";
        public static XName forceUpgrade = w + "forceUpgrade";
        public static XName forgetLastTabAlignment = w + "forgetLastTabAlignment";
        public static XName format = w + "format";
        public static XName formatting = w + "formatting";
        public static XName formProt = w + "formProt";
        public static XName formsDesign = w + "formsDesign";
        public static XName frame = w + "frame";
        public static XName frameLayout = w + "frameLayout";
        public static XName framePr = w + "framePr";
        public static XName frameset = w + "frameset";
        public static XName framesetSplitbar = w + "framesetSplitbar";
        public static XName ftr = w + "ftr";
        public static XName fullDate = w + "fullDate";
        public static XName gallery = w + "gallery";
        public static XName glossaryDocument = w + "glossaryDocument";
        public static XName grammar = w + "grammar";
        public static XName gridAfter = w + "gridAfter";
        public static XName gridBefore = w + "gridBefore";
        public static XName gridCol = w + "gridCol";
        public static XName gridSpan = w + "gridSpan";
        public static XName group = w + "group";
        public static XName growAutofit = w + "growAutofit";
        public static XName guid = w + "guid";
        public static XName gutter = w + "gutter";
        public static XName gutterAtTop = w + "gutterAtTop";
        public static XName h = w + "h";
        public static XName hAnchor = w + "hAnchor";
        public static XName hanging = w + "hanging";
        public static XName hangingChars = w + "hangingChars";
        public static XName hAnsi = w + "hAnsi";
        public static XName hAnsiTheme = w + "hAnsiTheme";
        public static XName hash = w + "hash";
        public static XName hdr = w + "hdr";
        public static XName hdrShapeDefaults = w + "hdrShapeDefaults";
        public static XName header = w + "header";
        public static XName headerReference = w + "headerReference";
        public static XName headerSource = w + "headerSource";
        public static XName helpText = w + "helpText";
        public static XName hidden = w + "hidden";
        public static XName hideGrammaticalErrors = w + "hideGrammaticalErrors";
        public static XName hideMark = w + "hideMark";
        public static XName hideSpellingErrors = w + "hideSpellingErrors";
        public static XName highlight = w + "highlight";
        public static XName hint = w + "hint";
        public static XName history = w + "history";
        public static XName hMerge = w + "hMerge";
        public static XName horzAnchor = w + "horzAnchor";
        public static XName hps = w + "hps";
        public static XName hpsBaseText = w + "hpsBaseText";
        public static XName hpsRaise = w + "hpsRaise";
        public static XName hRule = w + "hRule";
        public static XName hSpace = w + "hSpace";
        public static XName hyperlink = w + "hyperlink";
        public static XName hyphenationZone = w + "hyphenationZone";
        public static XName i = w + "i";
        public static XName iCs = w + "iCs";
        public static XName id = w + "id";
        public static XName ignoreMixedContent = w + "ignoreMixedContent";
        public static XName ilvl = w + "ilvl";
        public static XName imprint = w + "imprint";
        public static XName ind = w + "ind";
        public static XName initials = w + "initials";
        public static XName inkAnnotations = w + "inkAnnotations";
        public static XName ins = w + "ins";
        public static XName insDel = w + "insDel";
        public static XName insideH = w + "insideH";
        public static XName insideV = w + "insideV";
        public static XName instr = w + "instr";
        public static XName instrText = w + "instrText";
        public static XName isLgl = w + "isLgl";
        public static XName jc = w + "jc";
        public static XName keepLines = w + "keepLines";
        public static XName keepNext = w + "keepNext";
        public static XName kern = w + "kern";
        public static XName kinsoku = w + "kinsoku";
        public static XName lang = w + "lang";
        public static XName lastColumn = w + "lastColumn";
        public static XName lastRenderedPageBreak = w + "lastRenderedPageBreak";
        public static XName lastValue = w + "lastValue";
        public static XName lastRow = w + "lastRow";
        public static XName latentStyles = w + "latentStyles";
        public static XName layoutRawTableWidth = w + "layoutRawTableWidth";
        public static XName layoutTableRowsApart = w + "layoutTableRowsApart";
        public static XName leader = w + "leader";
        public static XName left = w + "left";
        public static XName leftChars = w + "leftChars";
        public static XName leftFromText = w + "leftFromText";
        public static XName legacy = w + "legacy";
        public static XName legacyIndent = w + "legacyIndent";
        public static XName legacySpace = w + "legacySpace";
        public static XName lid = w + "lid";
        public static XName line = w + "line";
        public static XName linePitch = w + "linePitch";
        public static XName lineRule = w + "lineRule";
        public static XName lines = w + "lines";
        public static XName lineWrapLikeWord6 = w + "lineWrapLikeWord6";
        public static XName link = w + "link";
        public static XName linkedToFile = w + "linkedToFile";
        public static XName linkStyles = w + "linkStyles";
        public static XName linkToQuery = w + "linkToQuery";
        public static XName listEntry = w + "listEntry";
        public static XName listItem = w + "listItem";
        public static XName listSeparator = w + "listSeparator";
        public static XName lnNumType = w + "lnNumType";
        public static XName _lock = w + "lock";
        public static XName locked = w + "locked";
        public static XName lsdException = w + "lsdException";
        public static XName lvl = w + "lvl";
        public static XName lvlJc = w + "lvlJc";
        public static XName lvlOverride = w + "lvlOverride";
        public static XName lvlPicBulletId = w + "lvlPicBulletId";
        public static XName lvlRestart = w + "lvlRestart";
        public static XName lvlText = w + "lvlText";
        public static XName mailAsAttachment = w + "mailAsAttachment";
        public static XName mailMerge = w + "mailMerge";
        public static XName mailSubject = w + "mailSubject";
        public static XName mainDocumentType = w + "mainDocumentType";
        public static XName mappedName = w + "mappedName";
        public static XName marBottom = w + "marBottom";
        public static XName marH = w + "marH";
        public static XName markup = w + "markup";
        public static XName marLeft = w + "marLeft";
        public static XName marRight = w + "marRight";
        public static XName marTop = w + "marTop";
        public static XName marW = w + "marW";
        public static XName matchSrc = w + "matchSrc";
        public static XName maxLength = w + "maxLength";
        public static XName mirrorIndents = w + "mirrorIndents";
        public static XName mirrorMargins = w + "mirrorMargins";
        public static XName monthLong = w + "monthLong";
        public static XName monthShort = w + "monthShort";
        public static XName moveFrom = w + "moveFrom";
        public static XName moveFromRangeEnd = w + "moveFromRangeEnd";
        public static XName moveFromRangeStart = w + "moveFromRangeStart";
        public static XName moveTo = w + "moveTo";
        public static XName moveToRangeEnd = w + "moveToRangeEnd";
        public static XName moveToRangeStart = w + "moveToRangeStart";
        public static XName multiLevelType = w + "multiLevelType";
        public static XName multiLine = w + "multiLine";
        public static XName mwSmallCaps = w + "mwSmallCaps";
        public static XName name = w + "name";
        public static XName namespaceuri = w + "namespaceuri";
        public static XName next = w + "next";
        public static XName nlCheck = w + "nlCheck";
        public static XName noBorder = w + "noBorder";
        public static XName noBreakHyphen = w + "noBreakHyphen";
        public static XName noColumnBalance = w + "noColumnBalance";
        public static XName noEndnote = w + "noEndnote";
        public static XName noExtraLineSpacing = w + "noExtraLineSpacing";
        public static XName noHBand = w + "noHBand";
        public static XName noLeading = w + "noLeading";
        public static XName noLineBreaksAfter = w + "noLineBreaksAfter";
        public static XName noLineBreaksBefore = w + "noLineBreaksBefore";
        public static XName noProof = w + "noProof";
        public static XName noPunctuationKerning = w + "noPunctuationKerning";
        public static XName noResizeAllowed = w + "noResizeAllowed";
        public static XName noSpaceRaiseLower = w + "noSpaceRaiseLower";
        public static XName noTabHangInd = w + "noTabHangInd";
        public static XName notTrueType = w + "notTrueType";
        public static XName noVBand = w + "noVBand";
        public static XName noWrap = w + "noWrap";
        public static XName nsid = w + "nsid";
        public static XName _null = w + "null";
        public static XName num = w + "num";
        public static XName numbering = w + "numbering";
        public static XName numberingChange = w + "numberingChange";
        public static XName numFmt = w + "numFmt";
        public static XName numId = w + "numId";
        public static XName numIdMacAtCleanup = w + "numIdMacAtCleanup";
        public static XName numPicBullet = w + "numPicBullet";
        public static XName numPicBulletId = w + "numPicBulletId";
        public static XName numPr = w + "numPr";
        public static XName numRestart = w + "numRestart";
        public static XName numStart = w + "numStart";
        public static XName numStyleLink = w + "numStyleLink";
        public static XName _object = w + "object";
        public static XName odso = w + "odso";
        public static XName offsetFrom = w + "offsetFrom";
        public static XName oMath = w + "oMath";
        public static XName optimizeForBrowser = w + "optimizeForBrowser";
        public static XName orient = w + "orient";
        public static XName original = w + "original";
        public static XName other = w + "other";
        public static XName outline = w + "outline";
        public static XName outlineLvl = w + "outlineLvl";
        public static XName overflowPunct = w + "overflowPunct";
        public static XName p = w + "p";
        public static XName pageBreakBefore = w + "pageBreakBefore";
        public static XName panose1 = w + "panose1";
        public static XName paperSrc = w + "paperSrc";
        public static XName pBdr = w + "pBdr";
        public static XName percent = w + "percent";
        public static XName permEnd = w + "permEnd";
        public static XName permStart = w + "permStart";
        public static XName personal = w + "personal";
        public static XName personalCompose = w + "personalCompose";
        public static XName personalReply = w + "personalReply";
        public static XName pgBorders = w + "pgBorders";
        public static XName pgMar = w + "pgMar";
        public static XName pgNum = w + "pgNum";
        public static XName pgNumType = w + "pgNumType";
        public static XName pgSz = w + "pgSz";
        public static XName pict = w + "pict";
        public static XName picture = w + "picture";
        public static XName pitch = w + "pitch";
        public static XName pixelsPerInch = w + "pixelsPerInch";
        public static XName placeholder = w + "placeholder";
        public static XName pos = w + "pos";
        public static XName position = w + "position";
        public static XName pPr = w + "pPr";
        public static XName pPrChange = w + "pPrChange";
        public static XName pPrDefault = w + "pPrDefault";
        public static XName prefixMappings = w + "prefixMappings";
        public static XName printBodyTextBeforeHeader = w + "printBodyTextBeforeHeader";
        public static XName printColBlack = w + "printColBlack";
        public static XName printerSettings = w + "printerSettings";
        public static XName printFormsData = w + "printFormsData";
        public static XName printFractionalCharacterWidth = w + "printFractionalCharacterWidth";
        public static XName printPostScriptOverText = w + "printPostScriptOverText";
        public static XName printTwoOnOne = w + "printTwoOnOne";
        public static XName proofErr = w + "proofErr";
        public static XName proofState = w + "proofState";
        public static XName pStyle = w + "pStyle";
        public static XName ptab = w + "ptab";
        public static XName qFormat = w + "qFormat";
        public static XName query = w + "query";
        public static XName r = w + "r";
        public static XName readModeInkLockDown = w + "readModeInkLockDown";
        public static XName recipientData = w + "recipientData";
        public static XName recommended = w + "recommended";
        public static XName relativeTo = w + "relativeTo";
        public static XName relyOnVML = w + "relyOnVML";
        public static XName removeDateAndTime = w + "removeDateAndTime";
        public static XName removePersonalInformation = w + "removePersonalInformation";
        public static XName restart = w + "restart";
        public static XName result = w + "result";
        public static XName revisionView = w + "revisionView";
        public static XName rFonts = w + "rFonts";
        public static XName richText = w + "richText";
        public static XName right = w + "right";
        public static XName rightChars = w + "rightChars";
        public static XName rightFromText = w + "rightFromText";
        public static XName rPr = w + "rPr";
        public static XName rPrChange = w + "rPrChange";
        public static XName rPrDefault = w + "rPrDefault";
        public static XName rsid = w + "rsid";
        public static XName rsidDel = w + "rsidDel";
        public static XName rsidP = w + "rsidP";
        public static XName rsidR = w + "rsidR";
        public static XName rsidRDefault = w + "rsidRDefault";
        public static XName rsidRoot = w + "rsidRoot";
        public static XName rsidRPr = w + "rsidRPr";
        public static XName rsids = w + "rsids";
        public static XName rsidSect = w + "rsidSect";
        public static XName rsidTr = w + "rsidTr";
        public static XName rStyle = w + "rStyle";
        public static XName rt = w + "rt";
        public static XName rtl = w + "rtl";
        public static XName rtlGutter = w + "rtlGutter";
        public static XName ruby = w + "ruby";
        public static XName rubyAlign = w + "rubyAlign";
        public static XName rubyBase = w + "rubyBase";
        public static XName rubyPr = w + "rubyPr";
        public static XName salt = w + "salt";
        public static XName saveFormsData = w + "saveFormsData";
        public static XName saveInvalidXml = w + "saveInvalidXml";
        public static XName savePreviewPicture = w + "savePreviewPicture";
        public static XName saveSmartTagsAsXml = w + "saveSmartTagsAsXml";
        public static XName saveSubsetFonts = w + "saveSubsetFonts";
        public static XName saveThroughXslt = w + "saveThroughXslt";
        public static XName saveXmlDataOnly = w + "saveXmlDataOnly";
        public static XName scrollbar = w + "scrollbar";
        public static XName sdt = w + "sdt";
        public static XName sdtContent = w + "sdtContent";
        public static XName sdtEndPr = w + "sdtEndPr";
        public static XName sdtPr = w + "sdtPr";
        public static XName sectPr = w + "sectPr";
        public static XName sectPrChange = w + "sectPrChange";
        public static XName selectFldWithFirstOrLastChar = w + "selectFldWithFirstOrLastChar";
        public static XName semiHidden = w + "semiHidden";
        public static XName sep = w + "sep";
        public static XName separator = w + "separator";
        public static XName settings = w + "settings";
        public static XName shadow = w + "shadow";
        public static XName shapeDefaults = w + "shapeDefaults";
        public static XName shapeid = w + "shapeid";
        public static XName shapeLayoutLikeWW8 = w + "shapeLayoutLikeWW8";
        public static XName shd = w + "shd";
        public static XName showBreaksInFrames = w + "showBreaksInFrames";
        public static XName showEnvelope = w + "showEnvelope";
        public static XName showingPlcHdr = w + "showingPlcHdr";
        public static XName showXMLTags = w + "showXMLTags";
        public static XName sig = w + "sig";
        public static XName size = w + "size";
        public static XName sizeAuto = w + "sizeAuto";
        public static XName smallCaps = w + "smallCaps";
        public static XName smartTag = w + "smartTag";
        public static XName smartTagPr = w + "smartTagPr";
        public static XName smartTagType = w + "smartTagType";
        public static XName snapToGrid = w + "snapToGrid";
        public static XName softHyphen = w + "softHyphen";
        public static XName solutionID = w + "solutionID";
        public static XName sourceFileName = w + "sourceFileName";
        public static XName space = w + "space";
        public static XName spaceForUL = w + "spaceForUL";
        public static XName spacing = w + "spacing";
        public static XName spacingInWholePoints = w + "spacingInWholePoints";
        public static XName specVanish = w + "specVanish";
        public static XName spelling = w + "spelling";
        public static XName splitPgBreakAndParaMark = w + "splitPgBreakAndParaMark";
        public static XName src = w + "src";
        public static XName start = w + "start";
        public static XName startOverride = w + "startOverride";
        public static XName statusText = w + "statusText";
        public static XName storeItemID = w + "storeItemID";
        public static XName storeMappedDataAs = w + "storeMappedDataAs";
        public static XName strictFirstAndLastChars = w + "strictFirstAndLastChars";
        public static XName strike = w + "strike";
        public static XName style = w + "style";
        public static XName styleId = w + "styleId";
        public static XName styleLink = w + "styleLink";
        public static XName styleLockQFSet = w + "styleLockQFSet";
        public static XName styleLockTheme = w + "styleLockTheme";
        public static XName stylePaneFormatFilter = w + "stylePaneFormatFilter";
        public static XName stylePaneSortMethod = w + "stylePaneSortMethod";
        public static XName styles = w + "styles";
        public static XName subDoc = w + "subDoc";
        public static XName subFontBySize = w + "subFontBySize";
        public static XName subsetted = w + "subsetted";
        public static XName suff = w + "suff";
        public static XName summaryLength = w + "summaryLength";
        public static XName suppressAutoHyphens = w + "suppressAutoHyphens";
        public static XName suppressBottomSpacing = w + "suppressBottomSpacing";
        public static XName suppressLineNumbers = w + "suppressLineNumbers";
        public static XName suppressOverlap = w + "suppressOverlap";
        public static XName suppressSpacingAtTopOfPage = w + "suppressSpacingAtTopOfPage";
        public static XName suppressSpBfAfterPgBrk = w + "suppressSpBfAfterPgBrk";
        public static XName suppressTopSpacing = w + "suppressTopSpacing";
        public static XName suppressTopSpacingWP = w + "suppressTopSpacingWP";
        public static XName swapBordersFacingPages = w + "swapBordersFacingPages";
        public static XName sym = w + "sym";
        public static XName sz = w + "sz";
        public static XName szCs = w + "szCs";
        public static XName t = w + "t";
        public static XName t1 = w + "t1";
        public static XName t2 = w + "t2";
        public static XName tab = w + "tab";
        public static XName table = w + "table";
        public static XName tabs = w + "tabs";
        public static XName tag = w + "tag";
        public static XName targetScreenSz = w + "targetScreenSz";
        public static XName tbl = w + "tbl";
        public static XName tblBorders = w + "tblBorders";
        public static XName tblCellMar = w + "tblCellMar";
        public static XName tblCellSpacing = w + "tblCellSpacing";
        public static XName tblGrid = w + "tblGrid";
        public static XName tblGridChange = w + "tblGridChange";
        public static XName tblHeader = w + "tblHeader";
        public static XName tblInd = w + "tblInd";
        public static XName tblLayout = w + "tblLayout";
        public static XName tblLook = w + "tblLook";
        public static XName tblOverlap = w + "tblOverlap";
        public static XName tblpPr = w + "tblpPr";
        public static XName tblPr = w + "tblPr";
        public static XName tblPrChange = w + "tblPrChange";
        public static XName tblPrEx = w + "tblPrEx";
        public static XName tblPrExChange = w + "tblPrExChange";
        public static XName tblpX = w + "tblpX";
        public static XName tblpXSpec = w + "tblpXSpec";
        public static XName tblpY = w + "tblpY";
        public static XName tblpYSpec = w + "tblpYSpec";
        public static XName tblStyle = w + "tblStyle";
        public static XName tblStyleColBandSize = w + "tblStyleColBandSize";
        public static XName tblStylePr = w + "tblStylePr";
        public static XName tblStyleRowBandSize = w + "tblStyleRowBandSize";
        public static XName tblW = w + "tblW";
        public static XName tc = w + "tc";
        public static XName tcBorders = w + "tcBorders";
        public static XName tcFitText = w + "tcFitText";
        public static XName tcMar = w + "tcMar";
        public static XName tcPr = w + "tcPr";
        public static XName tcPrChange = w + "tcPrChange";
        public static XName tcW = w + "tcW";
        public static XName temporary = w + "temporary";
        public static XName tentative = w + "tentative";
        public static XName text = w + "text";
        public static XName textAlignment = w + "textAlignment";
        public static XName textboxTightWrap = w + "textboxTightWrap";
        public static XName textDirection = w + "textDirection";
        public static XName textInput = w + "textInput";
        public static XName tgtFrame = w + "tgtFrame";
        public static XName themeColor = w + "themeColor";
        public static XName themeFill = w + "themeFill";
        public static XName themeFillShade = w + "themeFillShade";
        public static XName themeFillTint = w + "themeFillTint";
        public static XName themeFontLang = w + "themeFontLang";
        public static XName themeShade = w + "themeShade";
        public static XName themeTint = w + "themeTint";
        public static XName titlePg = w + "titlePg";
        public static XName tl2br = w + "tl2br";
        public static XName tmpl = w + "tmpl";
        public static XName tooltip = w + "tooltip";
        public static XName top = w + "top";
        public static XName topFromText = w + "topFromText";
        public static XName topLinePunct = w + "topLinePunct";
        public static XName tplc = w + "tplc";
        public static XName tr = w + "tr";
        public static XName tr2bl = w + "tr2bl";
        public static XName trackRevisions = w + "trackRevisions";
        public static XName trHeight = w + "trHeight";
        public static XName trPr = w + "trPr";
        public static XName trPrChange = w + "trPrChange";
        public static XName truncateFontHeightsLikeWP6 = w + "truncateFontHeightsLikeWP6";
        public static XName txbxContent = w + "txbxContent";
        public static XName type = w + "type";
        public static XName types = w + "types";
        public static XName u = w + "u";
        public static XName udl = w + "udl";
        public static XName uiCompat97To2003 = w + "uiCompat97To2003";
        public static XName uiPriority = w + "uiPriority";
        public static XName ulTrailSpace = w + "ulTrailSpace";
        public static XName underlineTabInNumList = w + "underlineTabInNumList";
        public static XName unhideWhenUsed = w + "unhideWhenUsed";
        public static XName updateFields = w + "updateFields";
        public static XName uri = w + "uri";
        public static XName url = w + "url";
        public static XName usb0 = w + "usb0";
        public static XName usb1 = w + "usb1";
        public static XName usb2 = w + "usb2";
        public static XName usb3 = w + "usb3";
        public static XName useAltKinsokuLineBreakRules = w + "useAltKinsokuLineBreakRules";
        public static XName useAnsiKerningPairs = w + "useAnsiKerningPairs";
        public static XName useFELayout = w + "useFELayout";
        public static XName useNormalStyleForList = w + "useNormalStyleForList";
        public static XName usePrinterMetrics = w + "usePrinterMetrics";
        public static XName useSingleBorderforContiguousCells = w + "useSingleBorderforContiguousCells";
        public static XName useWord2002TableStyleRules = w + "useWord2002TableStyleRules";
        public static XName useWord97LineBreakRules = w + "useWord97LineBreakRules";
        public static XName useXSLTWhenSaving = w + "useXSLTWhenSaving";
        public static XName val = w + "val";
        public static XName vAlign = w + "vAlign";
        public static XName value = w + "value";
        public static XName vAnchor = w + "vAnchor";
        public static XName vanish = w + "vanish";
        public static XName vendorID = w + "vendorID";
        public static XName vert = w + "vert";
        public static XName vertAlign = w + "vertAlign";
        public static XName vertAnchor = w + "vertAnchor";
        public static XName vertCompress = w + "vertCompress";
        public static XName view = w + "view";
        public static XName viewMergedData = w + "viewMergedData";
        public static XName vMerge = w + "vMerge";
        public static XName vMergeOrig = w + "vMergeOrig";
        public static XName vSpace = w + "vSpace";
        public static XName _w = w + "w";
        public static XName wAfter = w + "wAfter";
        public static XName wBefore = w + "wBefore";
        public static XName webHidden = w + "webHidden";
        public static XName webSettings = w + "webSettings";
        public static XName widowControl = w + "widowControl";
        public static XName wordWrap = w + "wordWrap";
        public static XName wpJustification = w + "wpJustification";
        public static XName wpSpaceWidth = w + "wpSpaceWidth";
        public static XName wrap = w + "wrap";
        public static XName wrapTrailSpaces = w + "wrapTrailSpaces";
        public static XName writeProtection = w + "writeProtection";
        public static XName x = w + "x";
        public static XName xAlign = w + "xAlign";
        public static XName xpath = w + "xpath";
        public static XName y = w + "y";
        public static XName yAlign = w + "yAlign";
        public static XName yearLong = w + "yearLong";
        public static XName yearShort = w + "yearShort";
        public static XName zoom = w + "zoom";
        public static XName zOrder = w + "zOrder";
        public static XName tblCaption = w + "tblCaption";
        public static XName tblDescription = w + "tblDescription";
        public static XName startChars = w + "startChars";
        public static XName end = w + "end";
        public static XName endChars = w + "endChars";
        public static XName evenHBand = w + "evenHBand";
        public static XName evenVBand = w + "evenVBand";
        public static XName firstRowFirstColumn = w + "firstRowFirstColumn";
        public static XName firstRowLastColumn = w + "firstRowLastColumn";
        public static XName lastRowFirstColumn = w + "lastRowFirstColumn";
        public static XName lastRowLastColumn = w + "lastRowLastColumn";
        public static XName oddHBand = w + "oddHBand";
        public static XName oddVBand = w + "oddVBand";
        public static XName headers = w + "headers";

        public static XName[] BlockLevelContentContainers =
        {
            W.body,
            W.tc,
            W.txbxContent,
            W.hdr,
            W.ftr,
            W.endnote,
            W.footnote
        };

        public static XName[] SubRunLevelContent =
        {
            W.br,
            W.cr,
            W.dayLong,
            W.dayShort,
            W.drawing,
            W.drawing,
            W.monthLong,
            W.monthShort,
            W.noBreakHyphen,
            W.ptab,
            W.pgNum,
            W.pict,
            W.softHyphen,
            W.sym,
            W.t,
            W.tab,
            W.yearLong,
            W.yearShort,
            MC.AlternateContent,
        };
    }

    public static class R
    {
        public static XNamespace r =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        public static XName blip = r + "blip";
        public static XName cs = r + "cs";
        public static XName dm = r + "dm";
        public static XName embed = r + "embed";
        public static XName href = r + "href";
        public static XName id = r + "id";
        public static XName link = r + "link";
        public static XName lo = r + "lo";
        public static XName pict = r + "pict";
        public static XName qs = r + "qs";
        public static XName verticalDpi = r + "verticalDpi";
    }

    public static class MC
    {
        public static XNamespace mc =
            "http://schemas.openxmlformats.org/markup-compatibility/2006";
        public static XName AlternateContent = mc + "AlternateContent";
        public static XName Choice = mc + "Choice";
        public static XName Fallback = mc + "Fallback";
        public static XName Ignorable = mc + "Ignorable";
        public static XName PreserveAttributes = mc + "PreserveAttributes";
    }

    public static class PresentationMLUtil
    {
        public static void FixUpPresentationDocument(PresentationDocument pDoc)
        {
            foreach (var part in pDoc.GetAllParts())
            {
                if (part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.slide+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.theme+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.drawingml.chart+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml" ||
                    part.ContentType == "application/vnd.ms-office.drawingml.diagramDrawing+xml")
                {
                    XDocument xd = part.GetXDocument();
                    xd.Descendants().Attributes("smtClean").Remove();
                    xd.Descendants().Attributes("smtId").Remove();
                    part.PutXDocument();
                }
               
                if (part.ContentType == "application/vnd.openxmlformats-officedocument.vmlDrawing")
                {
                    string fixedContent = null;

                    using (var stream = part.GetStream(FileMode.Open, FileAccess.ReadWrite))
                    {
                        using (var sr = new StreamReader(stream))
                        {
                            var input = sr.ReadToEnd();
                            string pattern = @"<!\[(?<test>.*)\]>";
                            fixedContent = Regex.Replace(input, pattern, m =>
                            {
                                var g = m.Groups[1].Value;
                                if (g.StartsWith("CDATA["))
                                    return "<![" + g + "]>";
                                else
                                    return "<![CDATA[" + g + "]]>";
                            },
                            RegexOptions.Multiline);

                            pattern = @"o:relid=[""'](?<id1>.*)[""'] o:relid=[""'](?<id2>.*)[""']";
                            fixedContent = Regex.Replace(fixedContent, pattern, m =>
                            {
                                var g = m.Groups[1].Value;
                                return @"o:relid=""" + g + @"""";
                            },
                            RegexOptions.Multiline);

                            fixedContent = fixedContent.Replace("</xml>ml>", "</xml>");

                            stream.SetLength(fixedContent.Length);
                        }
                    }
                    using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(fixedContent)))
                        part.FeedData(ms);
                }
            }
        }
    }
}
