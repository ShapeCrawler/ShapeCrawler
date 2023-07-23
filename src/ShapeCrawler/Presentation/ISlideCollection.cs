using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a collection of slides.
/// </summary>
public interface ISlideCollection : IReadOnlyList<ISlide>
{
    /// <summary>
    ///     Removes specified slide.
    /// </summary>
    void Remove(ISlide slide);

    /// <summary>
    ///     Adds a new slide based on the existing layout.
    /// </summary>
    /// <returns>A new slide.</returns>
    ISlide AddEmptySlide(ISlideLayout layout);

    /// <summary>
    ///     Adds a new slide based on the predefined layout type.
    /// </summary>
    /// <returns>A new slide.</returns>
    ISlide AddEmptySlide(SCSlideLayoutType layoutType);

    /// <summary>
    ///     Adds specified slide.
    /// </summary>
    void Add(ISlide slide);

    /// <summary>
    ///     Inserts slide at specified position.
    /// </summary>
    /// <param name="position">Position at which specified slide will be inserted.</param>
    /// <param name="slide">The slide to insert.</param>
    void Insert(int position, ISlide slide);
}

internal sealed class SCSlideCollection : ISlideCollection
{
    private readonly SCPresentation presentation;
    private readonly ResetAbleLazy<List<SCSlide>> slides;
    private PresentationPart presPart;

    internal SCSlideCollection(SCPresentation pres)
    {
        this.presentation = pres;
        this.presPart = pres.SDKPresentationInternal.PresentationPart!;
        this.slides = new ResetAbleLazy<List<SCSlide>>(this.GetSlides);
    }

    public int Count => this.slides.Value.Count;

    internal EventHandler? CollectionChanged { get; set; }

    public ISlide this[int index] => this.slides.Value[index];

    public IEnumerator<ISlide> GetEnumerator()
    {
        return this.slides.Value.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    public void Remove(ISlide slide)
    {
        // TODO: slide layout and master of removed slide also should be deleted if they are unused
        var sdkPresentation = this.presPart.Presentation;
        var slideIdList = sdkPresentation.SlideIdList!;
        var removingSlideIndex = slide.Number - 1;
        var removingSlideId = (P.SlideId)slideIdList.ChildElements[removingSlideIndex];
        var removingSlideRelId = removingSlideId.RelationshipId!;

        this.presentation.SectionsInternal.RemoveSldId(removingSlideId.Id!);

        slideIdList.RemoveChild(removingSlideId);
        RemoveFromCustomShow(sdkPresentation, removingSlideRelId);

        var removingSlidePart = (SlidePart)this.presPart.GetPartById(removingSlideRelId!);
        this.presPart.DeletePart(removingSlidePart);

        this.presPart.Presentation.Save();

        this.slides.Reset();

        this.OnCollectionChanged();
    }

    public ISlide AddEmptySlide(SCSlideLayoutType layoutType)
    {
        var masters = (SCSlideMasterCollection)this.presentation.SlideMasters;
        var layout = masters.SelectMany(m => m.SlideLayouts).First(l => l.Type == layoutType);

        return this.AddEmptySlide(layout);
    }

    public ISlide AddEmptySlide(ISlideLayout layout)
    {
        var rId = this.presPart.GetNextRelationshipId();
        SlidePart slidePart = this.presPart.AddNewSlidePart(rId);
        var layoutInternal = (SCSlideLayout)layout;
        slidePart.AddPart(layoutInternal.SlideLayoutPart, "rId1");

        // Copy layout placeholders
        if (layoutInternal.SlideLayoutPart.SlideLayout.CommonSlideData is P.CommonSlideData commonSlideData
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

        static P.TextBody ResolveTextBody(P.Shape shape)
        {
            // Creates a new TextBody
            if (shape.TextBody is null)
            {
                return new P.TextBody(new OpenXmlElement[]
                    { new A.Paragraph(new OpenXmlElement[] { new A.EndParagraphRunProperties() }) })
                {
                    BodyProperties = new A.BodyProperties(),
                    ListStyle = new A.ListStyle(),
                };
            }

            return (P.TextBody)shape.TextBody.CloneNode(true);
        }

        var pSlideIdList = this.presPart.Presentation.SlideIdList!;
        var nextId = pSlideIdList.OfType<P.SlideId>().Last().Id! + 1;
        var pSlideId = new P.SlideId { Id = nextId, RelationshipId = rId };
        pSlideIdList.Append(pSlideId);

        var newSlide = new SCSlide(this.presentation, slidePart, pSlideId);
        this.slides.Value.Add(newSlide);

        return newSlide;
    }

    public void Insert(int position, ISlide slide)
    {
        if (position < 1 || position > this.slides.Value.Count + 1)
        {
            throw new ArgumentOutOfRangeException(nameof(position));
        }

        this.Add(slide);
        int addedSlideIndex = this.slides.Value.Count - 1;
        this.slides.Value[addedSlideIndex].Number = position;

        this.slides.Reset();
        this.presentation.SlideMastersValue.Reset();
        this.OnCollectionChanged();
    }

    public void Add(ISlide slide)
    {
        var sourceSlideInternal = (SCSlide)slide;
        PresentationDocument sourcePresDoc;
        var tempStream = new MemoryStream();
        if (sourceSlideInternal.Presentation == this.presentation)
        {
            this.presentation.ChartWorkbooks.ForEach(c => c.Close());
            sourcePresDoc = (PresentationDocument)this.presentation.SDKPresentationInternal.Clone(tempStream);
        }
        else
        {
            sourcePresDoc =
                (PresentationDocument)sourceSlideInternal.PresentationInternal.SDKPresentationInternal
                    .Clone(tempStream);
        }

        var destPresDoc = this.presentation.SDKPresentationInternal;
        var sourcePresPart = sourcePresDoc.PresentationPart!;
        var destPresPart = destPresDoc.PresentationPart!;
        var destSdkPres = destPresPart.Presentation;
        var sourceSlideIndex = slide.Number - 1;
        var sourceSlideId = (P.SlideId)sourcePresPart.Presentation.SlideIdList!.ChildElements[sourceSlideIndex];
        var sourceSlidePart = (SlidePart)sourcePresPart.GetPartById(sourceSlideId.RelationshipId!);

        NormalizeLayouts(sourceSlidePart);

        var addedSlidePart = AddSlidePart(destPresPart, sourceSlidePart, out var addedSlideMasterPart);

        AddNewSlideId(destSdkPres, destPresDoc, addedSlidePart);
        var masterId = AddNewSlideMasterId(destSdkPres, destPresDoc, addedSlideMasterPart);
        AdjustLayoutIds(destPresDoc, masterId);

        this.slides.Reset();
        this.presentation.SlideMastersValue.Reset();

        this.CollectionChanged?.Invoke(this, EventArgs.Empty);
    }

    internal SCSlide GetBySlideId(string slideId)
    {
        return this.slides.Value.First(scSlide => scSlide.SlideId.Id == slideId);
    }

    private static SlidePart AddSlidePart(
        PresentationPart destPresPart,
        SlidePart sourceSlidePart,
        out SlideMasterPart addedSlideMasterPart)
    {
        var rId = destPresPart.GetNextRelationshipId();
        var addedSlidePart = destPresPart.AddPart(sourceSlidePart, rId);
        var sdkNoticePart = addedSlidePart.GetPartsOfType<NotesSlidePart>().FirstOrDefault();
        if (sdkNoticePart != null)
        {
            addedSlidePart.DeletePart(sdkNoticePart);
        }

        rId = destPresPart.GetNextRelationshipId();
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
            foreach (P.SlideLayoutId pSlideLayoutId in slideMasterPart.SlideMaster.SlideLayoutIdList!.OfType<P.SlideLayoutId>())
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

    private ISlide AddEmptySlide(Func<ISlideLayout, bool> query)
    {
        // Gets slide layoutName by type
        if (this.presentation.SlideMasters?[0] is not ISlideMaster slideMaster)
        {
            // TODO: add an exception.
            throw new Exception();
        }

        // Find layoutName of type.
        var layout = slideMaster.SlideLayouts.First(query);

        return this.AddEmptySlide(layout);
    }


    private List<SCSlide> GetSlides()
    {
        this.presPart = this.presentation.SDKPresentationInternal.PresentationPart!;
        int slidesCount = this.presPart.SlideParts.Count();
        var slides = new List<SCSlide>(slidesCount);
        var slideIds = this.presPart.Presentation.SlideIdList!.ChildElements.OfType<P.SlideId>().ToList();
        for (var slideIndex = 0; slideIndex < slidesCount; slideIndex++)
        {
            var slideId = slideIds[slideIndex];
            var slidePart = (SlidePart)this.presPart.GetPartById(slideId.RelationshipId!);
            var newSlide = new SCSlide(this.presentation, slidePart, slideId);
            slides.Add(newSlide);
        }

        return slides;
    }

    private void OnCollectionChanged()
    {
        this.CollectionChanged?.Invoke(this, EventArgs.Empty);
    }
}