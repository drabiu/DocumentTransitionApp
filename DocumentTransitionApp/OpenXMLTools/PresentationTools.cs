using System;
using System.Collections.Generic;
using System.Linq;
using Drawing = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using OpenXMLTools.Interfaces;
using System.Text;
using D = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;

namespace OpenXMLTools
{
    public class PresentationTools : IPresentationTools
    {
        #region Public methods

        public PresentationDocument InsertSlidesFromTemplate(PresentationDocument target, PresentationDocument template, IList<string> slideRelationshipIdList)
        {
            if (target == null)
            {
                throw new ArgumentNullException("target");
            }

            if (template == null)
            {
                throw new ArgumentNullException("template");
            }

            uint maxSlideId = 256;

            var slideIdList = target.PresentationPart.Presentation.SlideIdList;
            var presentationPart = target.PresentationPart;
            //Find the highest slide ID in the current list.
            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.Id.Value > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }
            }

            presentationPart.Presentation.SlideMasterIdList.RemoveAllChildren();
            uint uniqueId = GetMaxIdFromChild(presentationPart.Presentation.SlideMasterIdList);
            //check what if relationshipids repeat?
            foreach (string slideRelationshipId in slideRelationshipIdList)
            {
                maxSlideId++;
                uniqueId++;

                //Create the slide part and copy the data from the first part           
                //SlidePart newSlidePart = presentationPart.AddNewPart<SlidePart>();              
                var templateSlide = (SlidePart)template.PresentationPart.GetPartById(slideRelationshipId);
                var newIdFromTemplateId = string.Format("source{0}", slideRelationshipId);
                SlidePart newSlidePart = presentationPart.AddPart(templateSlide, newIdFromTemplateId);
                //newSlidePart.FeedData(templateSlide.GetStream());
                //Use the same slide layout as that of the template slide.
                //if (templateSlide.SlideLayoutPart != null)
                //{
                //    newSlidePart.AddPart(templateSlide.SlideLayoutPart);
                //}

                if (newSlidePart.SlideLayoutPart != null)
                {
                    SlideMasterPart destMasterPart = newSlidePart.SlideLayoutPart.SlideMasterPart;
                    presentationPart.AddPart(destMasterPart);

                    SlideMasterId newSlideMasterId = new SlideMasterId();
                    newSlideMasterId.RelationshipId = presentationPart.GetIdOfPart(destMasterPart);
                    newSlideMasterId.Id = uniqueId;

                    presentationPart.Presentation.SlideMasterIdList.Append(newSlideMasterId);
                }
                                
                //Insert the new slide into the slide list.
                SlideId newSlideId = slideIdList.AppendChild(new SlideId());

                //Set the slide id and relationship id
                newSlideId.Id = maxSlideId;
                newSlideId.RelationshipId = presentationPart.GetIdOfPart(newSlidePart);                             
            }

            FixSlideLayoutIds(presentationPart, uniqueId);
            //after adding OPEN XML SDK 3.0 this fix no longer needed but removes validation errors
            //PresentationMLUtil.FixUpPresentationDocument(target);

            target.PresentationPart.Presentation.Save();

            return target;
        }

        public PresentationDocument InsertSlidesFromTemplate(PresentationDocument target, PresentationDocument template)
        {
            var slideRelationshipIdList = template.PresentationPart.Presentation.SlideIdList.Elements<SlideId>().Select(s => s.RelationshipId.Value).ToList();

            return InsertSlidesFromTemplate(target, template, slideRelationshipIdList);
        }

        public PresentationDocument InsertNewSlide(PresentationDocument presentationDocument, int position, string slideTitle)
        {
            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            if (slideTitle == null)
            {
                throw new ArgumentNullException("slideTitle");
            }

            PresentationPart presentationPart = presentationDocument.PresentationPart;
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;
            // Verify that the presentation is not empty.
            if (presentationPart == null)
            {
                throw new InvalidOperationException("The presentation document is empty.");
            }

            if (position > slideIdList.ChildElements.Count)
            {
                throw new InvalidOperationException("The position is greather than number of slides");
            }

            // Declare and instantiate a new slide.
            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            uint drawingObjectId = 1;

            // Construct the slide content.            
            // Specify the non-visual properties of the new slide.
            NonVisualGroupShapeProperties nonVisualProperties = slide.CommonSlideData.ShapeTree.AppendChild(new NonVisualGroupShapeProperties());
            nonVisualProperties.NonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = 1, Name = "" };
            nonVisualProperties.NonVisualGroupShapeDrawingProperties = new NonVisualGroupShapeDrawingProperties();
            nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

            // Specify the group shape properties of the new slide.
            slide.CommonSlideData.ShapeTree.AppendChild(new GroupShapeProperties());

            // Declare and instantiate the title shape of the new slide.
            Shape titleShape = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());

            drawingObjectId++;

            // Specify the required shape properties for the title shape. 
            titleShape.NonVisualShapeProperties = new NonVisualShapeProperties
                (new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Title" },
                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title }));
            titleShape.ShapeProperties = new ShapeProperties();

            // Specify the text of the title shape.
            titleShape.TextBody = new TextBody(new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(new Drawing.Run(new Drawing.Text() { Text = slideTitle })));

            // Declare and instantiate the body shape of the new slide.
            Shape bodyShape = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());
            drawingObjectId++;

            // Specify the required shape properties for the body shape.
            bodyShape.NonVisualShapeProperties = new NonVisualShapeProperties(new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Content Placeholder" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1 }));
            bodyShape.ShapeProperties = new ShapeProperties();

            // Specify the text of the body shape.
            bodyShape.TextBody = new TextBody(new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph());

            // Create the slide part for the new slide.
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

            // Save the new slide part.
            slide.Save(slidePart);

            // Modify the slide ID list in the presentation part.
            // The slide ID list should not be null.


            // Find the highest slide ID in the current list.
            uint maxSlideId = 1;
            SlideId prevSlideId = null;

            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }

                position--;
                if (position == 0)
                {
                    prevSlideId = slideId;
                }

            }

            maxSlideId++;

            // Get the ID of the previous slide.
            SlidePart lastSlidePart;

            if (prevSlideId != null)
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
            }
            else
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
            }

            // Use the same slide layout as that of the previous slide.
            if (null != lastSlidePart.SlideLayoutPart)
            {
                slidePart.AddPart(lastSlidePart.SlideLayoutPart);
            }

            // Insert the new slide into the slide list after the previous slide.
            SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

            // Save the modified presentation.
            presentationPart.Presentation.Save();

            return presentationDocument;
        }

        public PresentationDocument RemoveAllSlides(PresentationDocument presentationDocument)
        {
            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            Presentation presentation = presentationDocument.PresentationPart.Presentation;
            SlideIdList slideIdList = presentation.SlideIdList;
            int index = CountSlides(presentationDocument) - 1;
            while(index >= 0)
            {
                DeleteSlide(presentationDocument, index);
                index--;
            }

            presentation.Save();

            return presentationDocument;
        }

        public PresentationDocument DeleteSlide(PresentationDocument presentationDocument, int slideIndex)
        {
            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            // Use the CountSlides sample to get the number of slides in the presentation.
            int slidesCount = CountSlides(presentationDocument);

            if (slideIndex < 0 || slideIndex >= slidesCount)
            {
                throw new InvalidOperationException("slideIndex out of range");
            }

            // Get the presentation part from the presentation document. 
            PresentationPart presentationPart = presentationDocument.PresentationPart;
            // Get the presentation from the presentation part.
            Presentation presentation = presentationPart.Presentation;
            // Get the list of slide IDs in the presentation.
            SlideIdList slideIdList = presentation.SlideIdList;
            // Get the slide ID of the specified slide
            SlideId slideId = slideIdList.ChildElements[slideIndex] as SlideId;
            // Get the relationship ID of the slide.
            string slideRelId = slideId.RelationshipId;
            // Remove the slide from the slide list.
            slideIdList.RemoveChild(slideId);
            if (presentation.CustomShowList != null)
            {
                // Iterate through the list of custom shows.
                foreach (var customShow in presentation.CustomShowList.Elements<CustomShow>())
                {
                    if (customShow.SlideList != null)
                    {
                        // Declare a link list of slide list entries.
                        LinkedList<SlideListEntry> slideListEntries = new LinkedList<SlideListEntry>();
                        foreach (SlideListEntry slideListEntry in customShow.SlideList.Elements())
                        {
                            // Find the slide reference to remove from the custom show.
                            if (slideListEntry.Id != null && slideListEntry.Id == slideRelId)
                            {
                                slideListEntries.AddLast(slideListEntry);
                            }
                        }

                        // Remove all references to the slide from the custom show.
                        foreach (SlideListEntry slideListEntry in slideListEntries)
                        {
                            customShow.SlideList.RemoveChild(slideListEntry);
                        }
                    }
                }
            }

            presentation.Save();
            // Get the slide part for the specified slide.
            SlidePart slidePart = presentationPart.GetPartById(slideRelId) as SlidePart;

            // Remove the slide part.
            presentationPart.DeletePart(slidePart);

            return presentationDocument;
        }

        #endregion

        #region static Public methods

        public static string GetSlideTitle(SlidePart slidePart, int nameLength)
        {
            if (slidePart == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            string paragraphSeparator = null;
            if (slidePart.Slide != null)
            {
                var shapes = from shape in slidePart.Slide.Descendants<Shape>()
                             where IsTitleShape(shape)
                             select shape;

                StringBuilder paragraphText = new StringBuilder();
                foreach (var shape in shapes)
                {
                    foreach (var paragraph in shape.TextBody.Descendants<D.Paragraph>())
                    {
                        paragraphText.Append(paragraphSeparator);
                        paragraphText.Append("[Sld]: ");
                        foreach (var text in paragraph.Descendants<D.Text>())
                        {
                            paragraphText.Append(text.Text);
                        }

                        paragraphSeparator = "\n";
                    }
                }

                StringBuilder result = new StringBuilder();
                var listWords = paragraphText.ToString().Split(default(char[]), StringSplitOptions.RemoveEmptyEntries);
                foreach (var word in listWords)
                {
                    result.Append(string.Format("{0} ", word));
                    if (result.Length > nameLength)
                        break;
                }

                result.Remove(result.Length - 1, 1);

                return result.ToString();
            }

            return string.Empty;
        }

        public static SlidePart GetSlidePart(PresentationDocument document, int index)
        {
            List<SlideId> slidesList = document.PresentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();

            return document.PresentationPart.GetPartById(slidesList[index].RelationshipId) as SlidePart;
        }

        public static uint GetMaxIdFromChild(OpenXmlElement el)
        {
            uint max = 2147483648;
            //Get max id value from set of children
            foreach (OpenXmlElement child in el.ChildElements)
            {
                OpenXmlAttribute attribute = child.GetAttribute("id", "");
                uint id = uint.Parse(attribute.Value);
                if (id > max)
                    max = id;
            }

            return max;
        }

        #endregion

        #region private methods

        private static void FixSlideLayoutIds(PresentationPart presPart, uint uniqueId)
        {
            //Need to make sure all slide layouts have unique ids
            foreach (SlideMasterPart slideMasterPart in presPart.SlideMasterParts)
            {
                foreach (SlideLayoutId slideLayoutId in slideMasterPart.SlideMaster.SlideLayoutIdList)
                {
                    uniqueId++;
                    slideLayoutId.Id = uniqueId;
                }

                slideMasterPart.SlideMaster.Save();
            }
        }

        private static bool IsTitleShape(Shape shape)
        {
            var placeholderShape = shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild<PlaceholderShape>();
            if (placeholderShape != null && placeholderShape.Type != null && placeholderShape.Type.HasValue)
            {
                switch ((PlaceholderValues)placeholderShape.Type)
                {
                    case PlaceholderValues.Title:
                    case PlaceholderValues.CenteredTitle:
                        return true;
                    default:
                        return false;
                }
            }

            return false;
        }

        private int CountSlides(PresentationDocument presentationDocument)
        {
            // Check for a null document object.
            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            int slidesCount = 0;
            // Get the presentation part of document.
            PresentationPart presentationPart = presentationDocument.PresentationPart;

            // Get the slide count from the SlideParts.
            if (presentationPart != null)
            {
                slidesCount = presentationPart.SlideParts.Count();
            }

            // Return the slide count to the previous method.
            return slidesCount;
        }

        #endregion
    }
}
