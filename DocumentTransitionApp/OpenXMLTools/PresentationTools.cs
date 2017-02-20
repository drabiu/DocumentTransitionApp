using System;
using System.Collections.Generic;
using System.Linq;
using Drawing = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using OpenXMLTools.Interfaces;

namespace OpenXMLTools
{
    public class PresentationTools : IPresentationTools
    {
        public void InsertSlideFromTemplate(PresentationPart presentationPart, MemoryStream mem, string sourceRelationshipId)
        {
            uint maxSlideId = 256;

            var slideIdList = presentationPart.Presentation.SlideIdList;
            //Find the highest slide ID in the current list.

            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.Id.Value > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }
            }

            maxSlideId++;
            //Create the slide part and copy the data from the first part
            SlidePart newSlidePart;
            //SlidePart newSlidePart = presentationPart.AddNewPart<SlidePart>();
            //Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            using (PresentationDocument templateDocument = PresentationDocument.Open(mem, false))
            {
                newSlidePart = (SlidePart)templateDocument.PresentationPart.GetPartById(sourceRelationshipId);
                //var templateSlide = (SlidePart)templateDocument.PresentationPart.GetPartById(sourceRelationshipId);
                //newSlidePart.FeedData(templateSlide.GetStream());
                //Use the same slide layout as that of the template slide.
                //if (templateSlide.SlideLayoutPart != null)
                //{
                // newSlidePart.AddPart(templateSlide.SlideLayoutPart);
                //}


                //slide.Save(newSlidePart);

                //Insert the new slide into the slide list.
                SlideId newSlideId = slideIdList.AppendChild(new SlideId());
                //SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), slideIdList.Last());

                string relationshipId = string.Format("cRelId{0}{1}", maxSlideId, new Random().Next(999));
                //Set the slide id and relationship id
                newSlideId.Id = maxSlideId;
                newSlideId.RelationshipId = relationshipId;
                presentationPart.AddPart(newSlidePart, relationshipId);

                presentationPart.Presentation.Save();
            }
        }

        public void InsertNewSlide(PresentationDocument presentationDocument, int position, string slideTitle)
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

            // Verify that the presentation is not empty.
            if (presentationPart == null)
            {
                throw new InvalidOperationException("The presentation document is empty.");
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
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

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
        }

        public void RemoveAllSlides(PresentationPart presentationPart)
        {
            Presentation presentation = presentationPart.Presentation;
            SlideIdList slideIdList = presentation.SlideIdList;

            foreach (SlideId slideId in slideIdList.ChildElements.ToList())
            {
                slideIdList.RemoveChild(slideId);
                string slideRelId = slideId.RelationshipId;

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
            }
        }
    }
}
