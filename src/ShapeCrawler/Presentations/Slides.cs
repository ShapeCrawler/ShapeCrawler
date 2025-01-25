using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace ShapeCrawler.Presentations;

internal sealed class Slides : ISlides
{
    private readonly ReadOnlySlides readOnlySlides;
    private readonly PresentationPart presentationPart;

    internal Slides(PresentationPart presentationPart)
        : this(presentationPart, new ReadOnlySlides(presentationPart.SlideParts))
    {
    }

    private Slides(PresentationPart presentationPart, ReadOnlySlides readOnlySlides)
    {
        this.presentationPart = presentationPart;
        this.readOnlySlides = readOnlySlides;
    }

    public int Count => this.readOnlySlides.Count;

    public ISlide this[int index] => this.readOnlySlides[index];

    public IEnumerator<ISlide> GetEnumerator() => this.readOnlySlides.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public void Remove(ISlide slide)
    {
        // TODO: slide layout and master of removed slide also should be deleted if they are unused
        var presDocument = (PresentationDocument)this.presentationPart.OpenXmlPackage;
        var presPart = presDocument.PresentationPart!;
        var pPresentation = presPart.Presentation;
        var slideIdList = pPresentation.SlideIdList!;
        var removingPSlideId = (P.SlideId)slideIdList.ChildElements[slide.Number - 1];
        var sectionList = pPresentation.PresentationExtensionList?.Descendants<P14.SectionList>().FirstOrDefault();
        var removingSectionSlideIdListEntry = sectionList?.Descendants<P14.SectionSlideIdListEntry>()
            .FirstOrDefault(s => s.Id! == removingPSlideId.Id!);
        removingSectionSlideIdListEntry?.Remove();
        slideIdList.RemoveChild(removingPSlideId);
        pPresentation.Save();

        var removingSlideIdRelationshipId = removingPSlideId.RelationshipId!;
        new SPPresentation(pPresentation).RemoveSlideIdFromCustomShow(removingSlideIdRelationshipId.Value!);

        var removingSlidePart = (SlidePart)presPart.GetPartById(removingSlideIdRelationshipId!);
        presPart.DeletePart(removingSlidePart);

        presPart.Presentation.Save();
    }

    public void AddEmptySlide(SlideLayoutType layoutType)
    {
        var sdkPresDoc = (PresentationDocument)this.presentationPart.OpenXmlPackage;
        var slideMasters = new SlideMasterCollection(sdkPresDoc.PresentationPart!.SlideMasterParts);
        var layout = slideMasters.SelectMany(m => m.SlideLayouts).First(l => l.Type == layoutType);

        this.AddEmptySlide(layout);
    }

    public void AddEmptySlide(ISlideLayout layout)
    {
        var sdkPresDocument = (PresentationDocument)this.presentationPart.OpenXmlPackage;
        var sdkPresPart = sdkPresDocument.PresentationPart!;
        var rId = new SOpenXmlPart(sdkPresPart).NextRelationshipId();
        var sdkSlidePart = sdkPresPart.AddNewPart<SlidePart>(rId);
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

        var pSlideIdList = sdkPresPart.Presentation.SlideIdList!;
        var nextId = pSlideIdList.OfType<P.SlideId>().Any()
            ? pSlideIdList.OfType<P.SlideId>().Last().Id! + 1
            : 256; // according to the scheme, this id starts at 256
        var pSlideId = new P.SlideId { Id = nextId, RelationshipId = rId };
        pSlideIdList.Append(pSlideId);
    }

    public void Insert(int position, ISlide slide)
    {
        if (position < 1 || position > this.Count + 1)
        {
            throw new ArgumentOutOfRangeException(nameof(position));
        }

        this.Add(slide);
        int addedSlideIndex = this.Count - 1;
        this.readOnlySlides[addedSlideIndex].Number = position;
    }

    public void Add(ISlide slide)
    {
        var addingSlide = (Slide)slide;
        var addingSlidePresStream = new MemoryStream();
        var targetPresDocument = (PresentationDocument)this.presentationPart.OpenXmlPackage;
        var addingSlidePresDocument = addingSlide.SdkPresentationDocument().Clone(addingSlidePresStream);

        var sourceSlidePresPart = addingSlidePresDocument.PresentationPart!;
        var targetPresPart = targetPresDocument.PresentationPart!;
        var targetPres = targetPresPart.Presentation;
        var sourceSlideId = (P.SlideId)sourceSlidePresPart.Presentation.SlideIdList!.ChildElements[slide.Number - 1];
        var sourceSlidePart = (SlidePart)sourceSlidePresPart.GetPartById(sourceSlideId.RelationshipId!);

        new SSlideMasterPart(sourceSlidePart.SlideLayoutPart!.SlideMasterPart!).RemoveLayoutsExcept(sourceSlidePart
            .SlideLayoutPart!);

        var wrappedPresentationPart = new SPresentationPart(targetPresPart);
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
            return new P.TextBody(new DocumentFormat.OpenXml.Drawing.Paragraph([
                new DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties()
            ]))
            {
                BodyProperties = new DocumentFormat.OpenXml.Drawing.BodyProperties(),
                ListStyle = new DocumentFormat.OpenXml.Drawing.ListStyle(),
            };
        }

        return (P.TextBody)shape.TextBody.CloneNode(true);
    }

    private static void AdjustLayoutIds(PresentationDocument sdkPresDocDest, uint masterId)
    {
        foreach (var slideMasterPart in sdkPresDocDest.PresentationPart!.SlideMasterParts)
        {
            foreach (P.SlideLayoutId pSlideLayoutId in slideMasterPart.SlideMaster.SlideLayoutIdList!
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

    private static uint CreateId(P.SlideIdList slideIdList)
    {
        uint currentId = 0;
        foreach (P.SlideId slideId in slideIdList.OfType<P.SlideId>())
        {
            if (slideId.Id! > currentId)
            {
                currentId = slideId.Id!;
            }
        }

        return ++currentId;
    }

    private static uint CreateId(P.SlideMasterIdList slideMasterIdList)
    {
        uint currentId = 0;
        foreach (P.SlideMasterId masterId in slideMasterIdList)
        {
            if (masterId.Id! > currentId)
            {
                currentId = masterId.Id!;
            }
        }

        return ++currentId;
    }
}