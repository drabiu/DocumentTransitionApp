﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTools.Interfaces;
using OpenXMLTools.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using D = DocumentFormat.OpenXml.Drawing;

namespace OpenXMLTools
{
    public class WordTools : IWordTools
    {
        #region Public methods

        public WordprocessingDocument MergeWordMedia(WordprocessingDocument target, WordprocessingDocument source)
        {
            var sourceImageParts = source.MainDocumentPart.ImageParts;
            var targetImageParts = target.MainDocumentPart.ImageParts;
            var sourceRelIds = sourceImageParts.Select(s => source.MainDocumentPart.GetIdOfPart(s));
            var targetRelIds = targetImageParts.Select(t => target.MainDocumentPart.GetIdOfPart(t));
            //need to check differences if RemoveUnusedMedia hasn`t been used
            var relIdToAdd = sourceRelIds.Except(targetRelIds);
            HashSet<WordPartRelId> insertedItemsRelId = new HashSet<WordPartRelId>();
            foreach (var relId in relIdToAdd)
            {
                var sourcePart = source.MainDocumentPart.GetPartById(relId);
                var targetPart = target.MainDocumentPart.AddPart(sourcePart);
                var insertedRelId = target.MainDocumentPart.GetIdOfPart(targetPart);
                insertedItemsRelId.Add(new WordPartRelId(relId, insertedRelId));
            }

            FixMediaIds(target, insertedItemsRelId);

            return target;
        }

        public WordprocessingDocument MergeWordEmbeddings(WordprocessingDocument target, WordprocessingDocument source)
        {
            var sourceEmbbedings = source.MainDocumentPart.EmbeddedPackageParts;
            var targetEmbbedings = target.MainDocumentPart.EmbeddedPackageParts;
            var sourceRelIds = sourceEmbbedings.Select(s => source.MainDocumentPart.GetIdOfPart(s));
            var targetRelIds = targetEmbbedings.Select(t => target.MainDocumentPart.GetIdOfPart(t));
            var relIdToAdd = sourceRelIds.Except(targetRelIds);
            HashSet<WordPartRelId> insertedItemsRelId = new HashSet<WordPartRelId>();
            foreach (var relId in relIdToAdd)
            {
                var sourcePart = source.MainDocumentPart.GetPartById(relId);
                var targetPart = target.MainDocumentPart.AddPart(sourcePart);
                var insertedRelId = target.MainDocumentPart.GetIdOfPart(targetPart);
                insertedItemsRelId.Add(new WordPartRelId(relId, insertedRelId));
            }

            FixEmbeedingsIds(target, insertedItemsRelId);

            return target;
        }

        public WordprocessingDocument MergeWordCharts(WordprocessingDocument target, WordprocessingDocument source)
        {
            var charts = source.MainDocumentPart.ChartParts;
            return target;
        }

        public WordprocessingDocument RemoveUnusedMedia(WordprocessingDocument target)
        {

            return target;
        }

        public WordprocessingDocument RemoveUnusedEmbeddings(WordprocessingDocument target)
        {
            return target;
        }

        #endregion

        #region static Public methods

        public static StringBuilder GetWordsFromTextElements(StringBuilder text, int nameLength)
        {
            StringBuilder result = new StringBuilder();
            var listWords = text.ToString().Split(default(char[]), StringSplitOptions.RemoveEmptyEntries);
            foreach (var word in listWords)
            {
                result.Append(string.Format("{0} ", word));
                if (result.Length > nameLength)
                    break;
            }

            if (result.Length > 0)
                result.Remove(result.Length - 1, 1);

            return result;
        }

        public static HashSet<OpenXmlElement> GetAllSiblingListElements(Paragraph paragraph, List<OpenXmlElement> elements, int numberingId)
        {
            IList<OpenXmlElement> result = new List<OpenXmlElement>();
            if (GetNumberingId(paragraph) == numberingId)
            {
                result.Add(paragraph);
                var index = elements.FindIndex(e => e is Paragraph && (e as Paragraph).ParagraphId == paragraph.ParagraphId);
                foreach (var element in elements.Skip(index + 1))
                {
                    if (element is Paragraph && GetNumberingId(element as Paragraph) == numberingId)
                        result.Add(element);
                    else
                        break;
                }
            }

            return new HashSet<OpenXmlElement>(result);
        }

        public static int GetNumberingId(Paragraph paragraph)
        {
            int result = 0;
            var numberingProperties = paragraph.ParagraphProperties?.NumberingProperties;
            if (numberingProperties != null)
            {
                result = numberingProperties.NumberingId.Val?.Value ?? 0;
            }

            return result;
        }

        public static bool IsListParagraph(Paragraph paragraph)
        {

            var numberingProperties = paragraph.ParagraphProperties?.NumberingProperties;
            bool result = numberingProperties != null;

            return result;
        }

        public static bool HasWebHiddenRunProperties(Run run)
        {
            bool result = false;
            var runProperties = run.Descendants<RunProperties>();
            foreach (var runProp in runProperties)
            {
                var webHidden = runProp.ChildElements.OfType<WebHidden>();
                if (webHidden != null && webHidden.Count() > 0)
                    result = true;
            }

            return result;
        }

        #endregion

        #region Private methods

        private void FixMediaIds(WordprocessingDocument document, HashSet<WordPartRelId> insertedItemsRelId)
        {
            Body body = document.MainDocumentPart.Document.Body;
            var embeddedObjects = body.Descendants<EmbeddedObject>();
            var drawings = body.Descendants<Drawing>();
            foreach (var embeddedObject in embeddedObjects)
            {
                var imageDataObjects = embeddedObject.Descendants<ImageData>();
                foreach (var imageData in imageDataObjects)
                {
                    var changedRelId = insertedItemsRelId.FirstOrDefault(it => it.OldId == imageData.RelationshipId);
                    if (changedRelId != null)
                        imageData.RelationshipId = changedRelId.NewId;
                }
            }

            foreach (var drawing in drawings)
            {
                var blipObjects = drawing.Descendants<D.Blip>();
                foreach (var blip in blipObjects)
                {
                    var changedRelId = insertedItemsRelId.FirstOrDefault(it => it.OldId == blip.Embed);
                    if (changedRelId != null)
                        blip.Embed = changedRelId.NewId;
                }
            }
        }

        private void FixEmbeedingsIds(WordprocessingDocument document, HashSet<WordPartRelId> insertedItemsRelId)
        {
            var embeddedObjects = document.MainDocumentPart.Document.Body.Descendants<EmbeddedObject>();
            foreach (var embeddedObject in embeddedObjects)
            {
                var oleObjects = embeddedObject.Elements<OleObject>();
                foreach (var oleObject in oleObjects)
                {
                    var changedRelId = insertedItemsRelId.FirstOrDefault(it => it.OldId == oleObject.Id);
                    if (changedRelId != null)
                        oleObject.Id = changedRelId.NewId;
                }
            }
        }

        #endregion
    }
}
