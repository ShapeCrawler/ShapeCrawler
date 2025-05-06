using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Presentations;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

internal sealed class UpdatedSlideCollection(SlideCollection slideCollection, PresentationPart presPart) : ISlideCollection
{
    public int Count => slideCollection.Count;

    public ISlide this[int index] => slideCollection[index];

    public IEnumerator<ISlide> GetEnumerator() => slideCollection.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public void Add(int layoutNumber)
    {
        var rId = new SCOpenXmlPart(presPart).GetNextRelationshipId();
        var slidePart = presPart.AddNewPart<SlidePart>(rId);
        slidePart.Slide = new P.Slide(
            new P.CommonSlideData(
                new P.ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = (UInt32Value)1U, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties()),
                    new P.GroupShapeProperties(new A.TransformGroup()))),
            new P.ColorMapOverride(new A.MasterColorMapping()));
        var layout = new SlideMasterCollection(presPart.SlideMasterParts).SlideMaster(1).InternalSlideLayout(layoutNumber);
        slidePart.AddPart(layout.SlideLayoutPart, "rId1");

        // Check if we're using a blank layout - if so, don't copy any shapes
        if (layout.Name != "Blank" && 
            layout.SlideLayoutPart.SlideLayout.CommonSlideData is P.CommonSlideData commonSlideData &&
            commonSlideData.ShapeTree is P.ShapeTree shapeTree)
        {
            var placeholderShapes = shapeTree.ChildElements
                .OfType<P.Shape>()

                // Select all shapes with placeholder.
                .Where(shape => shape.NonVisualShapeProperties!
                    .OfType<P.ApplicationNonVisualDrawingProperties>()
                    .Any(anvdp => anvdp.PlaceholderShape is not null))
                
                // Different handling based on layout type
                .Where(shape =>
                {
                    var placeholderType = shape.NonVisualShapeProperties!
                        .OfType<P.ApplicationNonVisualDrawingProperties>()
                        .FirstOrDefault()?.PlaceholderShape?.Type?.Value;
                    
                    
                    return placeholderType != P.PlaceholderValues.Footer && 
                               placeholderType != P.PlaceholderValues.DateAndTime && 
                               placeholderType != P.PlaceholderValues.SlideNumber;
                })

                // And creates a new shape with the placeholder.
                .Select(shape => new P.Shape
                {
                    // Clone placeholder
                    NonVisualShapeProperties =
                        (P.NonVisualShapeProperties)shape.NonVisualShapeProperties!.CloneNode(true),

                    // Creates a new TextBody with no content.
                    TextBody = ResolveTextBody(shape),
                    ShapeProperties = new P.ShapeProperties()
                });

            slidePart.Slide.CommonSlideData = new P.CommonSlideData()
            {
                ShapeTree = new P.ShapeTree(placeholderShapes)
                {
                    GroupShapeProperties = (P.GroupShapeProperties)shapeTree.GroupShapeProperties!.CloneNode(true),
                    NonVisualGroupShapeProperties =
                        (P.NonVisualGroupShapeProperties)shapeTree.NonVisualGroupShapeProperties!.CloneNode(true)
                }
            };
        }

        var pSlideIdList = presPart.Presentation.SlideIdList!;
        var nextId = pSlideIdList.OfType<P.SlideId>().Any()
            ? pSlideIdList.OfType<P.SlideId>().Last().Id! + 1
            : 256; // according to the scheme, this id starts at 256
        var pSlideId = new P.SlideId { Id = nextId, RelationshipId = rId };
        pSlideIdList.Append(pSlideId);
    }

    public void Add(ISlide slide, int number)
    {
        if (number < 1 || number > this.Count + 1)
        {
            throw new ArgumentOutOfRangeException(nameof(number));
        }

        var presentationDocument = (PresentationDocument)presPart.OpenXmlPackage;
        var presentationPart = presentationDocument.PresentationPart!;
        var addingSlidePresDocument = slide.GetSDKPresentationDocument();
        var sourceSlidePresPart = addingSlidePresDocument.PresentationPart!;
        
        // Get the source slide's information
        var sourceSlideId = (P.SlideId)sourceSlidePresPart.Presentation.SlideIdList!.ChildElements[slide.Number - 1];
        var sourceSlidePart = (SlidePart)sourceSlidePresPart.GetPartById(sourceSlideId.RelationshipId!);
        
        // Clone the slide to avoid corrupting references
        SlidePart clonedSlidePart;
        
        // Create a clean new slide part in the target presentation with a predictable relationship ID
        string newSlideRelId = "rId" + new Random().Next(100000, 999999).ToString();
        clonedSlidePart = presentationPart.AddNewPart<SlidePart>(newSlideRelId);
        
        // Copy the content from source slide to new slide
        using (var sourceStream = sourceSlidePart.GetStream())
        {
            sourceStream.Position = 0;
            using (var destStream = clonedSlidePart.GetStream(FileMode.Create, FileAccess.Write))
            {
                sourceStream.CopyTo(destStream);
            }
        }
        
        // Ensure layout relationship is properly set up
        var sourceLayoutPart = sourceSlidePart.SlideLayoutPart;
        if (sourceLayoutPart != null)
        {
            // First see if there's a matching layout in the target presentation
            SlideLayoutPart? targetLayoutPart = null;
            
            // Try to find a matching layout by type or name
            foreach (var masterPart in presentationPart.SlideMasterParts)
            {
                foreach (var layoutPart in masterPart.SlideLayoutParts)
                {
                    if (LayoutsMatch(layoutPart, sourceLayoutPart))
                    {
                        targetLayoutPart = layoutPart;
                        break;
                    }
                }
                
                if (targetLayoutPart != null)
                    break;
            }
            
            // If no matching layout found, create one
            if (targetLayoutPart == null)
            {
                // Get or create a master part
                SlideMasterPart masterPart;
                if (presentationPart.SlideMasterParts.Any())
                {
                    masterPart = presentationPart.SlideMasterParts.First();
                }
                else
                {
                    masterPart = presentationPart.AddNewPart<SlideMasterPart>();
                    
                    // Copy the master content from source
                    var sourceMasterPart = sourceLayoutPart.SlideMasterPart;
                    if (sourceMasterPart != null)
                    {
                        using (var sourceStream = sourceMasterPart.GetStream())
                        {
                            sourceStream.Position = 0;
                            using (var destStream = masterPart.GetStream(FileMode.Create, FileAccess.Write))
                            {
                                sourceStream.CopyTo(destStream);
                            }
                        }
                    }
                }
                
                // Create a new layout part linked to the master
                targetLayoutPart = masterPart.AddNewPart<SlideLayoutPart>();
                
                // Copy the layout content
                using (var sourceStream = sourceLayoutPart.GetStream())
                {
                    sourceStream.Position = 0;
                    using (var destStream = targetLayoutPart.GetStream(FileMode.Create, FileAccess.Write))
                    {
                        sourceStream.CopyTo(destStream);
                    }
                }
            }
            
            // Link the new slide to the layout
            clonedSlidePart.AddPart(targetLayoutPart);
        }
        
        // Create a new slide ID with the correct position
        uint maxSlideId = 256; // Default starting ID
        if (presentationPart.Presentation.SlideIdList!.Elements<P.SlideId>().Any())
        {
            maxSlideId = presentationPart.Presentation.SlideIdList!.Elements<P.SlideId>()
                .Max(id => id.Id!.Value) + 1;
        }
        
        // Create the new slide ID
        var slideId = new P.SlideId 
        { 
            Id = maxSlideId, 
            RelationshipId = newSlideRelId 
        };
        
        // Insert at the specified position
        var slideIdList = presentationPart.Presentation.SlideIdList!;
        if (number > slideIdList.Elements<P.SlideId>().Count())
        {
            slideIdList.Append(slideId);
        }
        else
        {
            slideIdList.InsertAt(slideId, number - 1);
        }
        
        // Save changes
        presentationPart.Presentation.Save();
    }
    
    // Helper method to determine if two layouts match
    private bool LayoutsMatch(SlideLayoutPart layout1, SlideLayoutPart layout2)
    {
        // Compare by type if available
        if (layout1.SlideLayout.Type != null && layout2.SlideLayout.Type != null)
        {
            return layout1.SlideLayout.Type!.Value == layout2.SlideLayout.Type!.Value;
        }
        
        // Otherwise compare by name
        var name1 = layout1.SlideLayout.CommonSlideData?.Name?.Value;
        var name2 = layout2.SlideLayout.CommonSlideData?.Name?.Value;
        
        if (name1 != null && name2 != null)
        {
            return name1 == name2;
        }
        
        // If no reliable way to compare, just return false to be safe
        return false;
    }

    public void AddJSON(string jsonSlide)
    {
        throw new NotImplementedException();
    }

    public void Add(ISlide slide)
    {
        var targetPresDocument = (PresentationDocument)presPart.OpenXmlPackage;
        var addingSlidePresDocument = slide.GetSDKPresentationDocument();

        var sourceSlidePresPart = addingSlidePresDocument.PresentationPart!;
        var targetPresPart = targetPresDocument.PresentationPart!;
        var targetPres = targetPresPart.Presentation;
        var sourceSlideId = (P.SlideId)sourceSlidePresPart.Presentation.SlideIdList!.ChildElements[slide.Number - 1];
        var sourceSlidePart = (SlidePart)sourceSlidePresPart.GetPartById(sourceSlideId.RelationshipId!);

        new SCSlideMasterPart(sourceSlidePart.SlideLayoutPart!.SlideMasterPart!).RemoveLayoutsExcept(sourceSlidePart.SlideLayoutPart!);

        var wrappedPresentationPart = new SCPresentationPart(targetPresPart);
        wrappedPresentationPart.AddSlidePart(sourceSlidePart);
        var addedSlidePart = wrappedPresentationPart.Last<SlidePart>();
        var addedSlideMasterPart = wrappedPresentationPart.Last<SlideMasterPart>();

        AddNewSlideId(targetPresDocument, addedSlidePart);
        var masterId = AddNewSlideMasterId(targetPres, targetPresDocument, addedSlideMasterPart);
        AdjustLayoutIds(targetPresDocument, masterId);
    }

    private static P.TextBody ResolveTextBody(P.Shape shape)
    {
        if (shape.TextBody is null)
        {
            return new P.TextBody(new A.Paragraph(new A.EndParagraphRunProperties()))
            {
                BodyProperties = new A.BodyProperties(),
                ListStyle = new A.ListStyle(),
            };
        }

        return (P.TextBody)shape.TextBody.CloneNode(true);
    }

    private static void AdjustLayoutIds(PresentationDocument sdkPresDocDest, uint masterId)
    {
        foreach (var slideMasterPart in sdkPresDocDest.PresentationPart!.SlideMasterParts)
        {
            foreach (var pSlideLayoutId in slideMasterPart.SlideMaster.SlideLayoutIdList!
                         .OfType<P.SlideLayoutId>())
            {
                masterId++;
                pSlideLayoutId.Id = masterId;
            }

            slideMasterPart.SlideMaster.Save();
        }
    }

    private static uint AddNewSlideMasterId(
        P.Presentation sdkPresDest,
        PresentationDocument sdkPresDocDest,
        SlideMasterPart addedSlideMasterPart)
    {
        var masterId = CreateId(sdkPresDest.SlideMasterIdList!);
        P.SlideMasterId slideMaterId = new()
        {
            Id = masterId,
            RelationshipId = sdkPresDocDest.PresentationPart!.GetIdOfPart(addedSlideMasterPart!)
        };
        sdkPresDocDest.PresentationPart.Presentation.SlideMasterIdList!.Append(slideMaterId);
        sdkPresDocDest.PresentationPart.Presentation.Save();
        return masterId;
    }

    private static void AddNewSlideId(PresentationDocument targetSdkPresDoc, SlidePart addedSdkSlidePart)
    {
        P.SlideId slideId = new()
        {
            Id = CreateId(targetSdkPresDoc.PresentationPart!.Presentation.SlideIdList!),
            RelationshipId = targetSdkPresDoc.PresentationPart!.GetIdOfPart(addedSdkSlidePart)
        };
        targetSdkPresDoc.PresentationPart!.Presentation.SlideIdList!.Append(slideId);
    }

    private static uint CreateId(P.SlideIdList slideIdList)
    {
        uint currentId = 0;
        foreach (var slideId in slideIdList.OfType<P.SlideId>())
        {
            if (slideId.Id! > currentId)
            {
                currentId = slideId.Id!;
            }
        }

        return currentId + 1;
    }

    private static uint CreateId(P.SlideMasterIdList slideMasterIdList)
    {
        uint currentId = 0;
        foreach (var openXmlElement in slideMasterIdList)
        {
            var masterId = (P.SlideMasterId)openXmlElement;
            if (masterId.Id! > currentId)
            {
                currentId = masterId.Id!;
            }
        }

        return currentId + 1;
    }
}