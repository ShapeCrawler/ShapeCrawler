using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Presentations;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

internal sealed class UpdatedSlideCollection(SlideCollection slideCollection, PresentationPart presPart)
    : ISlideCollection
{
    public int Count => slideCollection.Count;

    public ISlide this[int index] => slideCollection[index];

    public IEnumerator<ISlide> GetEnumerator() => slideCollection.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => slideCollection.GetEnumerator();

    public void Add(int layoutNumber)
    {
        var rId = new SCOpenXmlPart(presPart).NextRelationshipId();
        var newSlidePart = presPart.AddNewPart<SlidePart>(rId);
        var layout = new SlideMasterCollection(presPart.SlideMasterParts).SlideMaster(1)
            .InternalSlideLayout(layoutNumber);
        newSlidePart.AddPart(layout.SlideLayoutPart, "rId1");

        newSlidePart.Slide = new P.Slide(layout.SlideLayoutPart.SlideLayout.CommonSlideData!.CloneNode(true));
        var removingShapes = newSlidePart.Slide.CommonSlideData!.ShapeTree!.OfType<P.Shape>()
            .Where(shape =>
            {
                var placeholderType = shape.NonVisualShapeProperties!
                    .OfType<P.ApplicationNonVisualDrawingProperties>()
                    .FirstOrDefault()?.PlaceholderShape?.Type?.Value;

                return placeholderType == P.PlaceholderValues.Footer ||
                       placeholderType == P.PlaceholderValues.DateAndTime ||
                       placeholderType == P.PlaceholderValues.SlideNumber;
            }).ToList();
        removingShapes.ForEach(shape => shape.Remove());

        // Older packages omit the SlideIdList element; create it to avoid null reference
        presPart.Presentation.SlideIdList ??= new P.SlideIdList();

        var pSlideIdList = presPart.Presentation.SlideIdList!;
        var nextId = pSlideIdList.OfType<P.SlideId>().Any()
            ? pSlideIdList.OfType<P.SlideId>().Last().Id! + 1
            : 256; // PowerPoint reserves IDs below 256 for built-in slides
        var pSlideId = new P.SlideId { Id = nextId, RelationshipId = rId };
        pSlideIdList.Append(pSlideId);
    }

    public void Add(ISlide slide, int slideNumber)
    {
        if (slideNumber < 1 || slideNumber > this.Count + 1)
        {
            throw new SCException(nameof(slideNumber));
        }

        var sourceSlidePresPart = slide.GetSDKPresentationPart();
        var sourceSlideId = (P.SlideId)sourceSlidePresPart.Presentation.SlideIdList!.ChildElements[slide.Number - 1];
        var sourceSlidePart = (SlidePart)sourceSlidePresPart.GetPartById(sourceSlideId.RelationshipId!);
        var presentationPart = ((PresentationDocument)presPart.OpenXmlPackage).PresentationPart!;
        string newSlideRelId = new SCOpenXmlPart(presentationPart).NextRelationshipId();
        var clonedSlidePart = new SCSlidePart(sourceSlidePart).CloneTo(presentationPart, newSlideRelId);

        SlideHyperlinkFix.FixSlideHyperlinks(sourceSlidePart, clonedSlidePart, presentationPart);
        InsertSlideAtPosition(presentationPart, newSlideRelId, slideNumber);

        presentationPart.Presentation.Save();
    }

    public void Add(int layoutNumber, int slideNumber)
    {
        if (slideNumber < 1 || slideNumber > this.Count + 1)
        {
            throw new ArgumentOutOfRangeException(nameof(slideNumber));
        }

        var newRelId = new SCOpenXmlPart(presPart).NextRelationshipId();
        var slidePart = presPart.AddNewPart<SlidePart>(newRelId);
        slidePart.Slide = new P.Slide(
            new P.CommonSlideData(
                new P.ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = (UInt32Value)1U, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties()),
                    new P.GroupShapeProperties(new A.TransformGroup()))),
            new P.ColorMapOverride(new A.MasterColorMapping()));

        var layout = new SlideMasterCollection(presPart.SlideMasterParts)
            .SlideMaster(1)
            .InternalSlideLayout(layoutNumber);
        slidePart.AddPart(layout.SlideLayoutPart, "rId1");

        if (layout.SlideLayoutPart.SlideLayout.CommonSlideData is { ShapeTree: { } shapeTree })
        {
            var placeholderShapes = shapeTree.ChildElements
                .OfType<P.Shape>()
                .Where(shape => shape.NonVisualShapeProperties!
                    .OfType<P.ApplicationNonVisualDrawingProperties>()
                    .Any(anvdp => anvdp.PlaceholderShape is not null))
                .Where(shape =>
                {
                    var placeholderType = shape.NonVisualShapeProperties!
                        .OfType<P.ApplicationNonVisualDrawingProperties>()
                        .FirstOrDefault()?.PlaceholderShape?.Type?.Value;

                    return placeholderType != P.PlaceholderValues.Footer &&
                           placeholderType != P.PlaceholderValues.DateAndTime &&
                           placeholderType != P.PlaceholderValues.SlideNumber;
                })
                .Select(shape => new P.Shape
                {
                    NonVisualShapeProperties =
                        (P.NonVisualShapeProperties)shape.NonVisualShapeProperties!.CloneNode(true),
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

        presPart.Presentation.SlideIdList ??= new P.SlideIdList();

        InsertSlideAtPosition(presPart, newRelId, slideNumber);

        presPart.Presentation.Save();
    }

    public void Add(ISlide slide)
    {
        var targetPresDocument = (PresentationDocument)presPart.OpenXmlPackage;
        var sourceSlidePresPart = slide.GetSDKPresentationPart();
        var targetPresPart = targetPresDocument.PresentationPart!;
        var targetPres = targetPresPart.Presentation;
        var sourceSlideId = (P.SlideId)sourceSlidePresPart.Presentation.SlideIdList!.ChildElements[slide.Number - 1];
        var sourceSlidePart = (SlidePart)sourceSlidePresPart.GetPartById(sourceSlideId.RelationshipId!);

        new SCSlideMasterPart(sourceSlidePart.SlideLayoutPart!.SlideMasterPart!).RemoveLayoutsExcept(sourceSlidePart
            .SlideLayoutPart!);

        var wrappedPresentationPart = new SCPresentationPart(targetPresPart);
        wrappedPresentationPart.AddSlidePart(sourceSlidePart);
        var addedSlidePart = wrappedPresentationPart.Last<SlidePart>();
        var addedSlideMasterPart = wrappedPresentationPart.Last<SlideMasterPart>();

        AddNewSlideId(targetPresDocument, addedSlidePart);
        var masterId = AddNewSlideMasterId(targetPres, targetPresDocument, addedSlideMasterPart);
        AdjustLayoutIds(targetPresDocument, masterId);

        var imageCatalog = new ImagePartCatalog();
        imageCatalog.SeedFrom(targetPresPart.SlideParts.SelectMany(sp => sp.ImageParts));
        imageCatalog.Deduplicate(addedSlidePart);

        SlideHyperlinkFix.FixSlideHyperlinks(sourceSlidePart, addedSlidePart, targetPresPart);
    }

    private static void InsertSlideAtPosition(PresentationPart presentationPart, string relationshipId, int position)
    {
        // Slide IDs must remain unique and monotonically increasing
        uint maxSlideId = 256; // Default starting ID
        if (presentationPart.Presentation.SlideIdList!.Elements<P.SlideId>().Any())
        {
            maxSlideId = presentationPart.Presentation.SlideIdList!.Elements<P.SlideId>()
                .Max(id => id.Id!.Value) + 1;
        }

        // Max ID increments each append; reuse latest for inserted slides
        var slideId = new P.SlideId { Id = maxSlideId, RelationshipId = relationshipId };

        // Maintain presentation order when inserting before the end
        var slideIdList = presentationPart.Presentation.SlideIdList!;
        if (position > slideIdList.Elements<P.SlideId>().Count())
        {
            slideIdList.Append(slideId);
        }
        else
        {
            slideIdList.InsertAt(slideId, position - 1);
        }
    }

    private static P.TextBody ResolveTextBody(P.Shape shape)
    {
        if (shape.TextBody is null)
        {
            return new P.TextBody(new A.Paragraph(new A.EndParagraphRunProperties()))
            {
                BodyProperties = new A.BodyProperties(), ListStyle = new A.ListStyle(),
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
            Id = masterId, RelationshipId = sdkPresDocDest.PresentationPart!.GetIdOfPart(addedSlideMasterPart!)
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

    private static bool LayoutsMatch(SlideLayoutPart layout1, SlideLayoutPart layout2)
    {
        // PowerPoint considers layout type when resolving placeholders
        if (layout1.SlideLayout.Type != null && layout2.SlideLayout.Type != null)
        {
            return layout1.SlideLayout.Type!.Value == layout2.SlideLayout.Type!.Value;
        }

        // Layouts fallback to name equality when type is missing
        var name1 = layout1.SlideLayout.CommonSlideData?.Name?.Value;
        var name2 = layout2.SlideLayout.CommonSlideData?.Name?.Value;

        if (name1 != null && name2 != null)
        {
            return name1 == name2;
        }

        // Unable to match reliably; require new layout to avoid corrupt links
        return false;
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