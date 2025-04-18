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

internal sealed class UpdatableSlideCollection : ISlideCollection
{
    private readonly SlideCollection slideCollection;
    private readonly PresentationPart presPart;

    internal UpdatableSlideCollection(PresentationPart presPart)
    {
        this.slideCollection = new SlideCollection(presPart.SlideParts);
        this.presPart = presPart;
    }

    public int Count => this.slideCollection.Count;

    public ISlide this[int index] => this.slideCollection[index];

    public IEnumerator<ISlide> GetEnumerator() => this.slideCollection.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public void AddEmptySlide(SlideLayoutType layoutType)
    {
        var presDocument = (PresentationDocument)this.presPart.OpenXmlPackage;
        var slideMasters = new SlideMasterCollection(presDocument.PresentationPart!.SlideMasterParts);
        var layout = slideMasters.SelectMany(m => m.SlideLayouts).First(l => l.Type == layoutType);

        this.AddEmptySlide(layout);
    }

    public void AddEmptySlide(ISlideLayout layout)
    {
        var rId = new SCOpenXmlPart(this.presPart).GetNextRelationshipId();
        var sdkSlidePart = this.presPart.AddNewPart<SlidePart>(rId);
        sdkSlidePart.Slide = new P.Slide(
            new P.CommonSlideData(
                new P.ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = (UInt32Value)1U, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties()),
                    new P.GroupShapeProperties(new A.TransformGroup()))),
            new P.ColorMapOverride(new A.MasterColorMapping()));
        var layoutInternal = (SlideLayout)layout;
        sdkSlidePart.AddPart(layoutInternal.SdkSlideLayoutPart(), "rId1");

        if (layoutInternal.SdkSlideLayoutPart().SlideLayout.CommonSlideData is P.CommonSlideData commonSlideData &&
            commonSlideData.ShapeTree is P.ShapeTree shapeTree)
        {
            var placeholderShapes = shapeTree.ChildElements
                .OfType<P.Shape>()

                // Select all shapes with placeholder.
                .Where(shape => shape.NonVisualShapeProperties!
                    .OfType<P.ApplicationNonVisualDrawingProperties>()
                    .Any(anvdp => anvdp.PlaceholderShape is not null))

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

            sdkSlidePart.Slide.CommonSlideData = new P.CommonSlideData()
            {
                ShapeTree = new P.ShapeTree(placeholderShapes)
                {
                    GroupShapeProperties = (P.GroupShapeProperties)shapeTree.GroupShapeProperties!.CloneNode(true),
                    NonVisualGroupShapeProperties =
                        (P.NonVisualGroupShapeProperties)shapeTree.NonVisualGroupShapeProperties!.CloneNode(true)
                }
            };
        }

        var pSlideIdList = this.presPart.Presentation.SlideIdList!;
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

        this.Add(slide);
        var addedSlideIndex = this.Count - 1;
        this.slideCollection[addedSlideIndex].Number = number;
    }

    public void AddJSON(string jsonSlide)
    {
        throw new NotImplementedException();
    }

    public void Add(ISlide slide)
    {
        var addingSlide = (Slide)slide;
        var addingSlidePresStream = new MemoryStream();
        var targetPresDocument = (PresentationDocument)this.presPart.OpenXmlPackage;
        var addingSlidePresDocument = addingSlide.SdkPresentationDocument().Clone(addingSlidePresStream);

        var sourceSlidePresPart = addingSlidePresDocument.PresentationPart!;
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
    }

    private static P.TextBody ResolveTextBody(P.Shape shape)
    {
        // Creates a new TextBody
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