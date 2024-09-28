using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace ShapeCrawler;

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
        var sdkPresentationDocument = (PresentationDocument)this.presentationPart.OpenXmlPackage;
        var sdkPresentationPart = sdkPresentationDocument.PresentationPart!;
        var pPresentation = sdkPresentationPart.Presentation;
        var slideIdList = pPresentation.SlideIdList!;
        var removingSlideIndex = slide.Number - 1;
        var removingSlideId = (P.SlideId)slideIdList.ChildElements[removingSlideIndex];
        var removingSlideRelId = removingSlideId.RelationshipId!;

        var sdkSectionList = pPresentation.PresentationExtensionList?.Descendants<P14.SectionList>().FirstOrDefault();
        var removing = sdkSectionList?.Descendants<P14.SectionSlideIdListEntry>()
            .FirstOrDefault(s => s.Id! == removingSlideId.Id!);
        removing?.Remove();
        pPresentation.Save();

        slideIdList.RemoveChild(removingSlideId);
        RemoveFromCustomShow(pPresentation, removingSlideRelId);

        var removingSlidePart = (SlidePart)sdkPresentationPart.GetPartById(removingSlideRelId!);
        sdkPresentationPart.DeletePart(removingSlidePart);

        sdkPresentationPart.Presentation.Save();
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
        var rId = sdkPresPart.NextRelationshipId();
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
        sdkSlidePart.AddPart(layoutInternal.SDKSlideLayoutPart(), "rId1");

        // Copy layout placeholders
        if (layoutInternal.SDKSlideLayoutPart().SlideLayout.CommonSlideData is P.CommonSlideData commonSlideData && commonSlideData.ShapeTree is P.ShapeTree shapeTree) // && layout.Type != SlideLayoutType.Blank) // if the layout is blank, no need to add placeholder
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
        var d = this.readOnlySlides[addedSlideIndex];
        this.readOnlySlides[addedSlideIndex].Number = position;
    }

    public void Add(ISlide slide)
    {
        var sourceSlide = (Slide)slide;
        var sourcePresStream = new MemoryStream();
        var targetPresDocument = (PresentationDocument)this.presentationPart.OpenXmlPackage;
        var sourceSlidePresDocument = sourceSlide.SDKPresentationDocument().Clone(sourcePresStream);

        var sourceSlidePresPart = sourceSlidePresDocument.PresentationPart!;
        var targetPresPart = targetPresDocument.PresentationPart!;
        var targetPres = targetPresPart.Presentation;
        var sourceSlideId = (P.SlideId)sourceSlidePresPart.Presentation.SlideIdList!.ChildElements[slide.Number - 1];
        var sourceSlidePart = (SlidePart)sourceSlidePresPart.GetPartById(sourceSlideId.RelationshipId!);

        NormalizeLayouts(sourceSlidePart);

        var wrappedPresentationPart = new WrappedPresentationPart(targetPresPart);
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
            return new P.TextBody(new OpenXmlElement[]
            {
                new DocumentFormat.OpenXml.Drawing.Paragraph([new DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties()])
            })
            {
                BodyProperties = new DocumentFormat.OpenXml.Drawing.BodyProperties(),
                ListStyle = new DocumentFormat.OpenXml.Drawing.ListStyle(),
            };
        }

        return (P.TextBody)shape.TextBody.CloneNode(true);
    }

    private static void NormalizeLayouts(SlidePart sourceSlidePart)
    {
        var sourceMasterPart = sourceSlidePart.SlideLayoutPart!.SlideMasterPart!;
        var layoutParts = sourceMasterPart.SlideLayoutParts.ToList();
        var layoutIdList = sourceMasterPart.SlideMaster.SlideLayoutIdList!.OfType<P.SlideLayoutId>();
        foreach (var layoutPart in layoutParts)
        {
            if (layoutPart == sourceSlidePart.SlideLayoutPart)
            {
                continue;
            }

            var id = sourceMasterPart.GetIdOfPart(layoutPart);
            var layoutId = layoutIdList.First(x => x.RelationshipId == id);
            layoutId.Remove();
            sourceMasterPart.DeletePart(layoutPart);
        }
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
        foreach (P.SlideId slideId in slideIdList.OfType<P.SlideId>())
        {
            if (slideId.Id! > currentId)
            {
                currentId = slideId.Id!;
            }
        }

        return ++currentId;
    }

    private static void RemoveFromCustomShow(P.Presentation sdkPresentation, StringValue? removingSlideRelId)
    {
        if (sdkPresentation.CustomShowList == null)
        {
            return;
        }

        // Iterate through the list of custom shows
        foreach (var customShow in sdkPresentation.CustomShowList.Elements<P.CustomShow>())
        {
            if (customShow.SlideList == null)
            {
                continue;
            }

            // declares a link list of slide list entries
            var slideListEntries = new LinkedList<P.SlideListEntry>();
            foreach (P.SlideListEntry pSlideListEntry in customShow.SlideList.OfType<P.SlideListEntry>())
            {
                // finds the slide reference to remove from the custom show
                if (pSlideListEntry.Id != null && pSlideListEntry.Id == removingSlideRelId)
                {
                    slideListEntries.AddLast(pSlideListEntry);
                }
            }

            // Removes all references to the slide from the custom show
            foreach (P.SlideListEntry slideListEntry in slideListEntries)
            {
                customShow.SlideList.RemoveChild(slideListEntry);
            }
        }
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