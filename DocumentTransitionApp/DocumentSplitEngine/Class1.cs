using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Linq;

public class Form1
{


    private void Button1_Click(object sender, EventArgs e)
    {
        string path = null;
        string sourceFileName = null;
        string destinationFileName = null;

        path = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\";
        sourceFileName = path + "Source.pptx";
        destinationFileName = path + "Destination.pptx";

        PresentationDocument pptDestinationDoc = PresentationDocument.Open(destinationFileName, true);
        PresentationPart PP = pptDestinationDoc.PresentationPart;
        Presentation P = PP.Presentation;
        SlideIdList SIL = P.SlideIdList;

        //Find the highest slide ID in the current list.
        uint maxSlideId = 1;
        SlideId prevSlideId = null;

        foreach (SlideId slideId in SIL.ChildElements)
        {
            if (slideId.Id.Value > maxSlideId)
            {
                maxSlideId = slideId.Id;
            }
        }

        maxSlideId += 1;

        //Get the ID of the first slide.
        SlidePart firstSlidePart = default(SlidePart);
        firstSlidePart = (SlidePart)PP.GetPartById((SIL.ChildElements[0] as SlideId).RelationshipId);

        //Create the slide part and copy the data from the first part
        SlidePart SP = PP.AddNewPart<SlidePart>();
        using (PresentationDocument pptSourceDoc = PresentationDocument.Open(sourceFileName, true))
        {
            SP.FeedData(pptSourceDoc.PresentationPart.SlideParts.First().GetStream());
            pptSourceDoc.Close();
        }

        //Use the same slide layout as that of the first slide.
        if (null != firstSlidePart.SlideLayoutPart)
        {
            SP.AddPart(firstSlidePart.SlideLayoutPart);
        }

        //Insert the new slide into the slide list.
        SlideId newSI = P.SlideIdList.InsertAfter(new SlideId(), P.SlideIdList.Last());

        //Set the slide id and relationship id
        newSI.Id = maxSlideId;
        newSI.RelationshipId = PP.GetIdOfPart(SP);

        //Save the modified presentation.
        P.Save();

        //Close the destination presentation
        pptDestinationDoc.Close();

    }
}

//=======================================================
//Service provided by Telerik (www.telerik.com)
//Conversion powered by NRefactory.
//Twitter: @telerik
//Facebook: facebook.com/telerik
//=======================================================
