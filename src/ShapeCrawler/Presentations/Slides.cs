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
    private readonly IEnumerable<SlidePart> sdkSlideParts;
    private readonly ReadOnlySlides readOnlySlides;

    internal Slides(IEnumerable<SlidePart> sdkSlideParts)
        : this(sdkSlideParts, new ReadOnlySlides(sdkSlideParts))
    {
        this.sdkSlideParts = sdkSlideParts;
    }

    private Slides(IEnumerable<SlidePart> sdkSlideParts, ReadOnlySlides readOnlySlides)
    {
        this.sdkSlideParts = sdkSlideParts;
        this.readOnlySlides = readOnlySlides;
    }

    public int Count => this.readOnlySlides.Count;
    public ISlide this[int index] => this.readOnlySlides[index];
    public IEnumerator<ISlide> GetEnumerator()=>this.readOnlySlides.GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator()=>this.GetEnumerator();

    public void Remove(ISlide slide)
    {
        // TODO: slide layout and master of removed slide also should be deleted if they are unused
        var sdkPresentationDocument = (PresentationDocument)this.sdkSlideParts.First().OpenXmlPackage;
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
        var sdkPresentationDocument = (PresentationDocument)this.sdkSlideParts.First().OpenXmlPackage;
        var slideMaster = new SlideMasterCollection(sdkPresentationDocument.PresentationPart!.SlideMasterParts);
        var layout = slideMaster.SelectMany(m => m.SlideLayouts).First(l => l.Type == layoutType);

        this.AddEmptySlide(layout);
    }

    public void AddEmptySlide(ISlideLayout layout)
    {
        var sdkPresDocument = (PresentationDocument)this.sdkSlideParts.First().OpenXmlPackage;
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
        if (layoutInternal.SDKSlideLayoutPart().SlideLayout.CommonSlideData is P.CommonSlideData commonSlideData
            && commonSlideData.ShapeTree is P.ShapeTree shapeTree)
        {
            var placeholderShapes = shapeTree.ChildElements
                .OfType<P.Shape>()

                // Select all shapes with placeholder.
                .Where(shape => shape.NonVisualShapeProperties!
                    .OfType<P.ApplicationNonVisualDrawingProperties>()
                    .Any(anvdp => anvdp.PlaceholderShape is not null))

                // And creates a new shape with the placeholder.
                .Select(shape => new P.Shape()
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
        var nextId = pSlideIdList.OfType<P.SlideId>().Last().Id! + 1;
        var pSlideId = new P.SlideId { Id = nextId, RelationshipId = rId };
        pSlideIdList.Append(pSlideId);
    }

    private static P.TextBody ResolveTextBody(P.Shape shape)
    {
        // Creates a new TextBody
        if (shape.TextBody is null)
        {
            return new P.TextBody(new OpenXmlElement[]
            {
                new DocumentFormat.OpenXml.Drawing.Paragraph(new OpenXmlElement[]
                    { new DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties() })
            })
            {
                BodyProperties = new DocumentFormat.OpenXml.Drawing.BodyProperties(),
                ListStyle = new DocumentFormat.OpenXml.Drawing.ListStyle(),
            };
        }

        return (P.TextBody)shape.TextBody.CloneNode(true);
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
        var addingSlideInternal = (Slide)slide;
        PresentationDocument sourcePresDoc;
        var tempStream = new MemoryStream();
        var currentSdkPresDocument = (PresentationDocument)this.sdkSlideParts.First().OpenXmlPackage;
        var addingSlideSdkPresDocumentCopy =
            (PresentationDocument)addingSlideInternal.SDKPresentationDocument().Clone(tempStream);

        var addingSlideSdkPresPart = addingSlideSdkPresDocumentCopy.PresentationPart!;
        var destPresPart = currentSdkPresDocument.PresentationPart!;
        var destSdkPres = destPresPart.Presentation;
        var sourceSlideIndex = slide.Number - 1;
        var sourceSlideId = (P.SlideId)addingSlideSdkPresPart.Presentation.SlideIdList!.ChildElements[sourceSlideIndex];
        var sourceSlidePart = (SlidePart)addingSlideSdkPresPart.GetPartById(sourceSlideId.RelationshipId!);

        NormalizeLayouts(sourceSlidePart);

        var addedSlidePart = AddSlidePart(destPresPart, sourceSlidePart, out var addedSlideMasterPart);

        AddNewSlideId(destSdkPres, currentSdkPresDocument, addedSlidePart);
        var masterId = AddNewSlideMasterId(destSdkPres, currentSdkPresDocument, addedSlideMasterPart);
        AdjustLayoutIds(currentSdkPresDocument, masterId);
    }

    private static SlidePart AddSlidePart(
        PresentationPart destPresPart,
        SlidePart sourceSlidePart,
        out SlideMasterPart addedSlideMasterPart)
    {
        var rId = destPresPart.NextRelationshipId();
        var addedSlidePart = destPresPart.AddPart(sourceSlidePart, rId);
        var sdkNoticePart = addedSlidePart.GetPartsOfType<NotesSlidePart>().FirstOrDefault();
        if (sdkNoticePart != null)
        {
            addedSlidePart.DeletePart(sdkNoticePart);
        }

        rId = destPresPart.NextRelationshipId();
        addedSlideMasterPart = destPresPart.AddPart(addedSlidePart.SlideLayoutPart!.SlideMasterPart!, rId);
        var layoutIdList = addedSlideMasterPart.SlideMaster!.SlideLayoutIdList!.OfType<P.SlideLayoutId>();
        foreach (var lId in layoutIdList.ToList())
        {
            if (!addedSlideMasterPart.TryGetPartById(lId!.RelationshipId!, out _))
            {
                lId.Remove();
            }
        }

        return addedSlidePart;
    }

    private static void NormalizeLayouts(SlidePart sourceSlidePart)
    {
        var sourceMasterPart = sourceSlidePart.SlideLayoutPart!.SlideMasterPart!;
        var layoutParts = sourceMasterPart.SlideLayoutParts.ToList();
        var layoutIdList = sourceMasterPart.SlideMaster!.SlideLayoutIdList!.OfType<P.SlideLayoutId>();
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

    private static void AddNewSlideId(
        P.Presentation sdkPresDest,
        PresentationDocument sdkPresDocDest,
        SlidePart addedSdkSlidePart)
    {
        P.SlideId slideId = new()
        {
            Id = CreateId(sdkPresDest.SlideIdList!),
            RelationshipId = sdkPresDocDest.PresentationPart!.GetIdOfPart(addedSdkSlidePart)
        };
        sdkPresDest.SlideIdList!.Append(slideId);
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