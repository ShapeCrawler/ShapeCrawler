using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Collections;

internal class SCSlideCollection : ISlideCollection
{
    private readonly SCPresentation presentation;
    private readonly ResettableLazy<List<SCSlide>> slides;
    private PresentationPart presentationPart;

    internal SCSlideCollection(SCPresentation presentation)
    {
        this.presentation = presentation;
        this.presentationPart = presentation.SDKPresentationInternal.PresentationPart!;
        this.slides = new ResettableLazy<List<SCSlide>>(this.GetSlides);
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

    public void Remove(ISlide removingSlide)
    {
        // TODO: slide layout and master of removed slide also should be deleted if they are unused
        var sdkPresentation = this.presentationPart.Presentation;
        var slideIdList = sdkPresentation.SlideIdList!;
        var removingSlideIndex = removingSlide.Number - 1;
        var removingSlideId = (P.SlideId)slideIdList.ChildElements[removingSlideIndex];
        var removingSlideRelId = removingSlideId.RelationshipId!;

        this.presentation.SectionsInternal.RemoveSldId(removingSlideId.Id!);

        slideIdList.RemoveChild(removingSlideId);
        RemoveFromCustomShow(sdkPresentation, removingSlideRelId);

        var removingSlidePart = (SlidePart)this.presentationPart.GetPartById(removingSlideRelId!);
        this.presentationPart.DeletePart(removingSlidePart);

        this.presentationPart.Presentation.Save();

        this.slides.Reset();

        this.OnCollectionChanged();
    }

    public void Insert(int position, ISlide outerSlide)
    {
        if (position < 1 || position > this.slides.Value.Count + 1)
        {
            throw new ArgumentOutOfRangeException(nameof(position));
        }

        this.Add(outerSlide);
        int addedSlideIndex = this.slides.Value.Count - 1;
        this.slides.Value[addedSlideIndex].Number = position;

        this.slides.Reset();
        this.presentation.SlideMastersValue.Reset();
        this.OnCollectionChanged();
    }

    public void Add(ISlide sourceSlide)
    {
        var sourceSlideInternal = (SCSlide)sourceSlide;
        PresentationDocument sourcePresDoc;
        var tempStream = new MemoryStream();
        if (sourceSlideInternal.Presentation == this.presentation)
        {
            sourcePresDoc = (PresentationDocument)this.presentation.SDKPresentationInternal.Clone(tempStream);
        }
        else
        {
            sourcePresDoc = (PresentationDocument)sourceSlideInternal.PresentationInternal.SDKPresentationInternal.Clone(tempStream);
        }

        var destPresDoc = this.presentation.SDKPresentationInternal;
        var sourcePresPart = sourcePresDoc.PresentationPart!;
        var destPresPart = destPresDoc.PresentationPart!;
        var destSdkPres = destPresPart.Presentation;
        var sourceSlideIndex = sourceSlide.Number - 1;
        var sourceSlideId = (SlideId)sourcePresPart.Presentation.SlideIdList!.ChildElements[sourceSlideIndex];
        var sourceSlidePart = (SlidePart)sourcePresPart.GetPartById(sourceSlideId.RelationshipId!);

        NormalizeLayouts(sourceSlidePart);

        var addedSlidePart = AddSlidePart(destPresPart, sourceSlidePart, out var addedSlideMasterPart);

        AddNewSlideId(destSdkPres, destPresDoc, addedSlidePart);
        var masterId = AddNewSlideMasterId(destSdkPres, destPresDoc, addedSlideMasterPart);
        AdjustLayoutIds(destPresDoc, masterId);

        this.slides.Reset();
        this.presentation.SlideMastersValue.Reset();
        this.OnCollectionChanged();
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
        var addedSlidePart = destPresPart.AddPart(sourceSlidePart);
        var sdkNoticePart = addedSlidePart.GetPartsOfType<NotesSlidePart>().FirstOrDefault();
        if (sdkNoticePart != null)
        {
            addedSlidePart.DeletePart(sdkNoticePart);
        }

        addedSlideMasterPart = destPresPart.AddPart(addedSlidePart.SlideLayoutPart!.SlideMasterPart!);
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
            foreach (SlideLayoutId slideLayoutId in slideMasterPart.SlideMaster.SlideLayoutIdList!)
            {
                masterId++;
                slideLayoutId.Id = masterId;
            }

            slideMasterPart.SlideMaster.Save();
        }
    }

    private static uint AddNewSlideMasterId(
        Presentation sdkPresDest, 
        PresentationDocument sdkPresDocDest,
        SlideMasterPart addedSlideMasterPart)
    {
        var masterId = CreateId(sdkPresDest.SlideMasterIdList!);
        SlideMasterId slideMaterId = new ()
        {
            Id = masterId,
            RelationshipId = sdkPresDocDest.PresentationPart!.GetIdOfPart(addedSlideMasterPart!)
        };
        sdkPresDocDest.PresentationPart.Presentation.SlideMasterIdList!.Append(slideMaterId);
        sdkPresDocDest.PresentationPart.Presentation.Save();
        return masterId;
    }

    private static void AddNewSlideId(
        Presentation sdkPresDest, 
        PresentationDocument sdkPresDocDest,
        SlidePart addedSdkSlidePart)
    {
        SlideId slideId = new ()
        {
            Id = CreateId(sdkPresDest.SlideIdList!),
            RelationshipId = sdkPresDocDest.PresentationPart!.GetIdOfPart(addedSdkSlidePart)
        };
        sdkPresDest.SlideIdList!.Append(slideId);
    }

    private static uint CreateId(SlideIdList slideIdList)
    {
        uint currentId = 0;
        foreach (SlideId slideId in slideIdList)
        {
            if (slideId.Id! > currentId)
            {
                currentId = slideId.Id!;
            }
        }

        return ++currentId;
    }

    private static uint CreateId(SlideMasterIdList slideMasterIdList)
    {
        uint currentId = 0;
        foreach (SlideMasterId masterId in slideMasterIdList)
        {
            if (masterId.Id! > currentId)
            {
                currentId = masterId.Id!;
            }
        }

        return ++currentId;
    }

    private List<SCSlide> GetSlides()
    {
        this.presentationPart = this.presentation.SDKPresentationInternal.PresentationPart!;
        int slidesCount = this.presentationPart.SlideParts.Count();
        var slides = new List<SCSlide>(slidesCount);
        var slideIds = this.presentationPart.Presentation.SlideIdList!.ChildElements.OfType<SlideId>().ToList();
        for (var slideIndex = 0; slideIndex < slidesCount; slideIndex++)
        {
            var slideId = slideIds[slideIndex];
            var slidePart = (SlidePart)this.presentationPart.GetPartById(slideId.RelationshipId!);
            var newSlide = new SCSlide(this.presentation, slidePart, slideId);
            slides.Add(newSlide);
        }

        return slides;
    }

    private void OnCollectionChanged()
    {
        this.CollectionChanged?.Invoke(this, null);
    }

    private static void RemoveFromCustomShow(Presentation sdkPresentation, StringValue? removingSlideRelId)
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
            foreach (P.SlideListEntry slideListEntry in customShow.SlideList.Elements())
            {
                // finds the slide reference to remove from the custom show
                if (slideListEntry.Id != null && slideListEntry.Id == removingSlideRelId)
                {
                    slideListEntries.AddLast(slideListEntry);
                }
            }

            // Removes all references to the slide from the custom show
            foreach (P.SlideListEntry slideListEntry in slideListEntries)
            {
                customShow.SlideList.RemoveChild(slideListEntry);
            }
        }
    }
}