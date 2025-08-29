using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Presentations;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
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

        // Ensure SlideIdList exists for presentations that don't initialize it
        presPart.Presentation.SlideIdList ??= new P.SlideIdList();

        var pSlideIdList = presPart.Presentation.SlideIdList!;
        var nextId = pSlideIdList.OfType<P.SlideId>().Any()
            ? pSlideIdList.OfType<P.SlideId>().Last().Id! + 1
            : 256; // according to the scheme, this id starts at 256
        var pSlideId = new P.SlideId { Id = nextId, RelationshipId = rId };
        pSlideIdList.Append(pSlideId);
    }

    public void Add(ISlide slide, int slideNumber)
    {
        if (slideNumber < 1 || slideNumber > this.Count + 1)
        {
            throw new ArgumentOutOfRangeException(nameof(slideNumber));
        }

        var sourceSlidePresPart = slide.GetSDKPresentationPart();
        var sourceSlideId = (P.SlideId)sourceSlidePresPart.Presentation.SlideIdList!.ChildElements[slide.Number - 1];
        var sourceSlidePart = (SlidePart)sourceSlidePresPart.GetPartById(sourceSlideId.RelationshipId!);

        var presentationPart = ((PresentationDocument)presPart.OpenXmlPackage).PresentationPart!;
        string newSlideRelId = new SCOpenXmlPart(presentationPart).NextRelationshipId();
        var clonedSlidePart = presentationPart.AddNewPart<SlidePart>(newSlideRelId);
        CopySlideContent(sourceSlidePart, clonedSlidePart);
        CopyCustomXmlParts(sourceSlidePart, clonedSlidePart);

        // Ensure all image relationships used by the slide XML are present in the cloned slide
        // so that PowerPoint can resolve picture blips correctly.
        CopyImageParts(sourceSlidePart, clonedSlidePart);

        // Ensure any chart parts referenced by the slide are present and correctly linked
        CopyChartParts(sourceSlidePart, clonedSlidePart);
        FixHyperlinkRelationships(sourceSlidePart, clonedSlidePart, presentationPart);
        LinkToLayoutPart(sourceSlidePart, clonedSlidePart, presentationPart);
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

        // Build a map of existing image parts in the target presentation keyed by the SHA512 hash of their content.
        var existingImagePartsByHash = new Dictionary<string, ImagePart>();
        foreach (var existingImagePart in targetPresPart.SlideParts.SelectMany(sp => sp.ImageParts))
        {
            var hash = ComputeHash(existingImagePart);
            if (!existingImagePartsByHash.ContainsKey(hash))
            {
                existingImagePartsByHash.Add(hash, existingImagePart);
            }
        } // After the slide part has been added, deduplicate any image parts that are identical to existing ones.

        DeduplicateImageParts(addedSlidePart, existingImagePartsByHash);

        // Fix hyperlink relationships to point to the correct slides in the target presentation
        FixHyperlinkRelationships(sourceSlidePart, addedSlidePart, targetPresPart);
    }

    private static void CopySlideContent(SlidePart sourceSlidePart, SlidePart clonedSlidePart)
    {
        using var sourceStream = sourceSlidePart.GetStream();
        sourceStream.Position = 0;
        using var destStream = clonedSlidePart.GetStream(FileMode.Create, FileAccess.Write);
        sourceStream.CopyTo(destStream);
    }

    private static void CopyCustomXmlParts(SlidePart sourceSlidePart, SlidePart clonedSlidePart)
    {
        var sourceCustomXmlParts = sourceSlidePart.CustomXmlParts.ToList();
        if (!sourceCustomXmlParts.Any())
        {
            return;
        }

        foreach (var sourceCustomXmlPart in sourceCustomXmlParts)
        {
            var newCustomXmlPart = clonedSlidePart.AddCustomXmlPart(sourceCustomXmlPart.ContentType);
            using var sourceStream = sourceCustomXmlPart.GetStream();
            sourceStream.Position = 0;
            using var destStream = newCustomXmlPart.GetStream(FileMode.Create, FileAccess.Write);
            sourceStream.CopyTo(destStream);
        }
    }

    private static void CopyImageParts(SlidePart sourceSlidePart, SlidePart clonedSlidePart)
    {
        // The slide XML we copied contains a:blip/@r:embed IDs. We must make sure
        // the cloned slide part has relationships with the SAME IDs pointing to
        // valid image parts. Otherwise, PowerPoint shows "The picture can't be displayed".
        var blips = clonedSlidePart.Slide.CommonSlideData!
            .ShapeTree!
            .Descendants<A.Blip>()
            .ToList();

        foreach (var blip in blips)
        {
            var relId = blip.Embed?.Value;
            if (string.IsNullOrWhiteSpace(relId))
            {
                continue;
            }

            // If relationship already exists on the cloned slide, skip
            if (clonedSlidePart.Parts.Any(p => p.RelationshipId == relId))
            {
                continue;
            }

            // Get the source image part using the same relationship id
            if (sourceSlidePart.TryGetPartById(relId!, out var openXmlPart) && openXmlPart is ImagePart srcImage)
            {
                // If the source and target belong to the same package we can share the part.
                // Otherwise, create a new part and copy the bytes.
                if (ReferenceEquals(sourceSlidePart.OpenXmlPackage, clonedSlidePart.OpenXmlPackage))
                {
                    clonedSlidePart.AddPart(srcImage, relId!);
                }
                else
                {
                    var dstImage = clonedSlidePart.AddNewPart<ImagePart>(srcImage.ContentType, relId);
                    using var s = srcImage.GetStream();
                    s.Position = 0;
                    using var d = dstImage.GetStream(FileMode.Create, FileAccess.Write);
                    s.CopyTo(d);
                }
            }
        }
    }

    private static void CopyChartParts(SlidePart sourceSlidePart, SlidePart clonedSlidePart)
    {
        foreach (var relId in GetChartRelationshipIds(clonedSlidePart))
        {
            EnsureChartRelationship(relId, sourceSlidePart, clonedSlidePart);
        }
    }

    private static IEnumerable<string> GetChartRelationshipIds(SlidePart slidePart)
    {
        return
        [
            .. slidePart.Slide.CommonSlideData!
                .ShapeTree!
                .Descendants<A.GraphicData>()
                .Where(gd => gd.Uri?.Value == "http://schemas.openxmlformats.org/drawingml/2006/chart")
                .Select(gd => gd.GetFirstChild<C.ChartReference>())
                .Where(cr => cr?.Id?.Value != null)
                .Select(cr => cr!.Id!.Value!)
                .Distinct()
        ];
    }

    private static void EnsureChartRelationship(
        string relationshipId,
        SlidePart sourceSlidePart,
        SlidePart targetSlidePart)
    {
        if (RelationshipExists(targetSlidePart, relationshipId))
        {
            return;
        }

        if (!TryGetSourceChartPart(sourceSlidePart, relationshipId, out var sourceChartPart))
        {
            return;
        }

        if (ReferenceEquals(sourceSlidePart.OpenXmlPackage, targetSlidePart.OpenXmlPackage))
        {
            ShareChartPartWithinSamePackage(sourceChartPart!, targetSlidePart, relationshipId);
            return;
        }

        CloneChartPartAcrossPackages(sourceChartPart!, targetSlidePart, relationshipId);
    }

    private static bool RelationshipExists(SlidePart slidePart, string relationshipId)
    {
        return slidePart.Parts.Any(p => p.RelationshipId == relationshipId);
    }

    private static bool TryGetSourceChartPart(
        SlidePart sourceSlidePart,
        string relationshipId,
        out ChartPart? sourceChartPart)
    {
        sourceChartPart = null;
        if (sourceSlidePart.TryGetPartById(relationshipId, out var part) && part is ChartPart cp)
        {
            sourceChartPart = cp;
            return true;
        }

        return false;
    }

    private static void ShareChartPartWithinSamePackage(
        ChartPart sourceChartPart,
        SlidePart targetSlidePart,
        string relationshipId)
    {
        targetSlidePart.AddPart(sourceChartPart, relationshipId);
    }

    private static void CloneChartPartAcrossPackages(
        ChartPart sourceChartPart,
        SlidePart targetSlidePart,
        string relationshipId)
    {
        var targetChartPart = targetSlidePart.AddNewPart<ChartPart>(sourceChartPart.ContentType, relationshipId);
        CopyStream(sourceChartPart, targetChartPart);
        CopyChartChildParts(sourceChartPart, targetChartPart);
    }

    private static void CopyStream(OpenXmlPart sourcePart, OpenXmlPart targetPart)
    {
        using var s = sourcePart.GetStream();
        s.Position = 0;
        using var d = targetPart.GetStream(FileMode.Create, FileAccess.Write);
        s.CopyTo(d);
    }

    private static void CopyChartChildParts(ChartPart sourceChartPart, ChartPart targetChartPart)
    {
        foreach (var child in sourceChartPart.Parts)
        {
            var childRelId = child.RelationshipId;
            var childPart = child.OpenXmlPart;
            if (childPart is EmbeddedPackagePart embeddedPackagePart)
            {
                CopyEmbeddedPackagePart(embeddedPackagePart, targetChartPart, childRelId);
            }
            else
            {
                // Best-effort: link existing part into the new chart
                targetChartPart.AddPart(childPart, childRelId);
            }
        }
    }

    private static void CopyEmbeddedPackagePart(
        EmbeddedPackagePart sourceEmbeddedPackagePart,
        ChartPart targetChartPart,
        string relationshipId)
    {
        var destinationPart = targetChartPart.AddNewPart<EmbeddedPackagePart>(
            sourceEmbeddedPackagePart.ContentType,
            relationshipId);
        using var es = sourceEmbeddedPackagePart.GetStream();
        es.Position = 0;
        using var ed = destinationPart.GetStream(FileMode.Create, FileAccess.Write);
        es.CopyTo(ed);
    }

    private static SlideLayoutPart CreateNewLayout(PresentationPart presentationPart, SlideLayoutPart sourceLayoutPart)
    {
        var masterPart = GetOrCreateMasterPart(presentationPart, sourceLayoutPart);

        // Create a new layout part linked to the master
        var targetLayoutPart = masterPart.AddNewPart<SlideLayoutPart>();

        // Copy the layout content
        CopyPartContent(sourceLayoutPart, targetLayoutPart);

        return targetLayoutPart;
    }

    private static SlideMasterPart GetOrCreateMasterPart(
        PresentationPart presentationPart,
        SlideLayoutPart sourceLayoutPart)
    {
        if (presentationPart.SlideMasterParts.Any())
        {
            return presentationPart.SlideMasterParts.First();
        }

        var masterPart = presentationPart.AddNewPart<SlideMasterPart>();

        // Copy the master content from source
        var sourceMasterPart = sourceLayoutPart.SlideMasterPart;
        if (sourceMasterPart != null)
        {
            CopyPartContent(sourceMasterPart, masterPart);
        }

        return masterPart;
    }

    private static void CopyPartContent(OpenXmlPart sourcePart, OpenXmlPart targetPart)
    {
        using var sourceStream = sourcePart.GetStream();
        sourceStream.Position = 0;
        using var destStream = targetPart.GetStream(FileMode.Create, FileAccess.Write);
        sourceStream.CopyTo(destStream);
    }

    private static void InsertSlideAtPosition(PresentationPart presentationPart, string relationshipId, int position)
    {
        // Create a new slide ID with the correct position
        uint maxSlideId = 256; // Default starting ID
        if (presentationPart.Presentation.SlideIdList!.Elements<P.SlideId>().Any())
        {
            maxSlideId = presentationPart.Presentation.SlideIdList!.Elements<P.SlideId>()
                .Max(id => id.Id!.Value) + 1;
        }

        // Create the new slide ID
        var slideId = new P.SlideId { Id = maxSlideId, RelationshipId = relationshipId };

        // Insert at the specified position
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

    private static void LinkToLayoutPart(
        SlidePart sourceSlidePart,
        SlidePart clonedSlidePart,
        PresentationPart presentationPart)
    {
        var sourceLayoutPart = sourceSlidePart.SlideLayoutPart;
        if (sourceLayoutPart == null)
        {
            return;
        }

        var targetLayoutPart = FindMatchingLayout(presentationPart, sourceLayoutPart) ??
                               CreateNewLayout(presentationPart, sourceLayoutPart);

        // Link the new slide to the layout
        clonedSlidePart.AddPart(targetLayoutPart);
    }

    private static SlideLayoutPart? FindMatchingLayout(
        PresentationPart presentationPart,
        SlideLayoutPart sourceLayoutPart)
    {
        foreach (var masterPart in presentationPart.SlideMasterParts)
        {
            foreach (var layoutPart in masterPart.SlideLayoutParts)
            {
                if (LayoutsMatch(layoutPart, sourceLayoutPart))
                {
                    return layoutPart;
                }
            }
        }

        return null;
    }

    private static string ComputeHash(OpenXmlPart part)
    {
        using var sha512 = SHA512.Create();
        using var stream = part.GetStream();
        stream.Position = 0;
        var hash = sha512.ComputeHash(stream);
        stream.Position = 0;
        return Convert.ToBase64String(hash);
    }

    private static void DeduplicateImageParts(
        SlidePart addedSlidePart,
        IDictionary<string, ImagePart> existingImagePartsByHash)
    {
        foreach (var imagePart in addedSlidePart.ImageParts.ToList())
        {
            var hash = ComputeHash(imagePart);
            if (existingImagePartsByHash.TryGetValue(hash, out var existingPart) &&
                !ReferenceEquals(existingPart, imagePart))
            {
                var relId = addedSlidePart.GetIdOfPart(imagePart);
                addedSlidePart.DeletePart(imagePart);
                addedSlidePart.AddPart(existingPart, relId);
            }
            else
            {
                existingImagePartsByHash[hash] = imagePart;
            }
        }
    }

    private static void FixHyperlinkRelationships(
        SlidePart sourceSlidePart,
        SlidePart clonedSlidePart,
        PresentationPart targetPresentationPart)
    {
        var sourcePresentation = ((PresentationDocument)sourceSlidePart.OpenXmlPackage).PresentationPart!;

        // Find all hyperlinks in the cloned slide that reference slides
        var hyperlinks = clonedSlidePart.Slide.Descendants<A.HyperlinkOnClick>()
            .Where(h => h.Action?.Value == "ppaction://hlinksldjump" && !string.IsNullOrEmpty(h.Id?.Value));

        foreach (var hyperlink in hyperlinks)
        {
            try
            {
                // Get the original relationship from the source slide
                var sourceTargetSlidePart = (SlidePart)sourceSlidePart.GetPartById(hyperlink.Id!.Value!);

                // Find which slide number this was in the source presentation
                var sourceSlideIdList = sourcePresentation.Presentation.SlideIdList!.ChildElements.OfType<P.SlideId>();
                var sourceTargetSlideRelId = sourcePresentation.GetIdOfPart(sourceTargetSlidePart);

                var sourceSlideNumber = 0;
                foreach (var slideId in sourceSlideIdList)
                {
                    sourceSlideNumber++;
                    if (slideId.RelationshipId == sourceTargetSlideRelId)
                    {
                        break;
                    }
                }

                // Get the corresponding slide in the target presentation
                var targetSlideIdList =
                    targetPresentationPart.Presentation.SlideIdList!.ChildElements.OfType<P.SlideId>();
                var targetSlideId = targetSlideIdList.ElementAtOrDefault(sourceSlideNumber - 1);

                if (targetSlideId != null)
                {
                    var targetSlidePart = (SlidePart)targetPresentationPart.GetPartById(targetSlideId.RelationshipId!);

                    // Create a new relationship from the cloned slide to the target slide
                    var newRelationship = clonedSlidePart.AddPart(targetSlidePart);
                    var newRelId = clonedSlidePart.GetIdOfPart(newRelationship);

                    // Update the hyperlink to use the new relationship ID
                    hyperlink.Id = newRelId;
                }
            }
            catch
            {
                // If we can't fix the hyperlink, remove it to prevent validation errors
                hyperlink.Id = null;
                hyperlink.Action = null;
            }
        }
    }
}