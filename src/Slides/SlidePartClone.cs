using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Slides;

/// <summary>
/// Provides helpers for cloning slide parts and reconnecting their dependent relationships.
/// </summary>
internal sealed class SlidePartClone
{
    private readonly SlidePart sourceSlidePart;
    
    /// <summary>
    /// Initializes a new instance of the <see cref="SlidePartClone"/> class.
    /// </summary>
    /// <param name="sourceSlidePart">Slide part that serves as the cloning origin.</param>
    internal SlidePartClone(SlidePart sourceSlidePart)
    {
        this.sourceSlidePart = sourceSlidePart;
    }
    
    private static void CopyStream(OpenXmlPart sourcePart, OpenXmlPart targetPart)
    {
        using var sourceStream = sourcePart.GetStream();
        sourceStream.Position = 0;
        using var destinationStream = targetPart.GetStream(FileMode.Create, FileAccess.Write);
        sourceStream.CopyTo(destinationStream);
    }

    /// <summary>
    /// Clones the wrapped slide part to the specified presentation part using the provided relationship id.
    /// </summary>
    /// <param name="targetPresentationPart">Destination presentation part.</param>
    /// <param name="relationshipId">Relationship identifier to use for the new slide.</param>
    /// <returns>Cloned slide part instance.</returns>
    internal SlidePart CloneTo(PresentationPart targetPresentationPart, string relationshipId)
    {
        var clonedSlidePart = targetPresentationPart.AddNewPart<SlidePart>(relationshipId);
        this.CopySlideContent(clonedSlidePart);
        this.CopyCustomXmlParts(clonedSlidePart);
        this.CopyImageParts(clonedSlidePart);
        this.CopyChartParts(clonedSlidePart);
        this.LinkToLayoutPart(targetPresentationPart, clonedSlidePart);

        return clonedSlidePart;
    }
    
    
    private static IEnumerable<string> GetChartRelationshipIds(SlidePart slidePart)
    {
        return slidePart.Slide.CommonSlideData!
            .ShapeTree!
            .Descendants<GraphicData>()
            .Where(graphicData => graphicData.Uri?.Value == "http://schemas.openxmlformats.org/drawingml/2006/chart")
            .Select(graphicData => graphicData.GetFirstChild<ChartReference>())
            .Where(chartReference => chartReference?.Id?.Value != null)
            .Select(chartReference => chartReference!.Id!.Value!)
            .Distinct();
    }

    private static bool RelationshipExists(SlidePart slidePart, string relationshipId)
    {
        return slidePart.Parts.Any(part => part.RelationshipId == relationshipId);
    }

    private static void ShareChartPartWithinSamePackage(ChartPart sourceChartPart, SlidePart targetSlidePart, string relationshipId)
    {
        targetSlidePart.AddPart(sourceChartPart, relationshipId);
    }

    private static void CloneChartPartAcrossPackages(ChartPart sourceChartPart, SlidePart targetSlidePart, string relationshipId)
    {
        var destinationChartPart = targetSlidePart.AddNewPart<ChartPart>(sourceChartPart.ContentType, relationshipId);
        CopyStream(sourceChartPart, destinationChartPart);
        CopyChartChildParts(sourceChartPart, destinationChartPart);
    }

    private static void CopyChartChildParts(ChartPart sourceChartPart, ChartPart targetChartPart)
    {
        foreach (var child in sourceChartPart.Parts)
        {
            var childRelationshipId = child.RelationshipId;
            var childPart = child.OpenXmlPart;
            if (childPart is EmbeddedPackagePart embeddedPackagePart)
            {
                CopyEmbeddedPackagePart(embeddedPackagePart, targetChartPart, childRelationshipId);
            }
            else
            {
                targetChartPart.AddPart(childPart, childRelationshipId);
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
        using var sourceStream = sourceEmbeddedPackagePart.GetStream();
        sourceStream.Position = 0;
        using var destinationStream = destinationPart.GetStream(FileMode.Create, FileAccess.Write);
        sourceStream.CopyTo(destinationStream);
    }

    private static SlideLayoutPart? FindMatchingLayout(PresentationPart presentationPart, SlideLayoutPart sourceLayoutPart)
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

    private static bool LayoutsMatch(SlideLayoutPart layout1, SlideLayoutPart layout2)
    {
        if (layout1.SlideLayout.Type != null && layout2.SlideLayout.Type != null)
        {
            return layout1.SlideLayout.Type!.Value == layout2.SlideLayout.Type!.Value;
        }

        var name1 = layout1.SlideLayout.CommonSlideData?.Name?.Value;
        var name2 = layout2.SlideLayout.CommonSlideData?.Name?.Value;

        if (name1 != null && name2 != null)
        {
            return string.Equals(name1, name2, StringComparison.Ordinal);
        }

        return false;
    }

    private static SlideLayoutPart CreateNewLayout(PresentationPart presentationPart, SlideLayoutPart sourceLayoutPart)
    {
        var masterPart = GetOrCreateMasterPart(presentationPart, sourceLayoutPart);
        var targetLayoutPart = masterPart.AddNewPart<SlideLayoutPart>();
        CopyPartContent(sourceLayoutPart, targetLayoutPart);
        return targetLayoutPart;
    }

    private static SlideMasterPart GetOrCreateMasterPart(PresentationPart presentationPart, SlideLayoutPart sourceLayoutPart)
    {
        if (presentationPart.SlideMasterParts.Any())
        {
            return presentationPart.SlideMasterParts.First();
        }

        var masterPart = presentationPart.AddNewPart<SlideMasterPart>();
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
        using var destinationStream = targetPart.GetStream(FileMode.Create, FileAccess.Write);
        sourceStream.CopyTo(destinationStream);
    }

    private bool TryGetSourceChartPart(string relationshipId, out ChartPart? sourceChartPart)
    {
        sourceChartPart = null;
        if (this.sourceSlidePart.TryGetPartById(relationshipId, out var part) && part is ChartPart chartPart)
        {
            sourceChartPart = chartPart;
            return true;
        }

        return false;
    }

    private void LinkToLayoutPart(PresentationPart presentationPart, SlidePart clonedSlidePart)
    {
        var sourceLayoutPart = this.sourceSlidePart.SlideLayoutPart;
        if (sourceLayoutPart == null)
        {
            return;
        }

        var targetLayoutPart = FindMatchingLayout(presentationPart, sourceLayoutPart) ??
                               CreateNewLayout(presentationPart, sourceLayoutPart);

        clonedSlidePart.AddPart(targetLayoutPart);
    }

    private void CopySlideContent(SlidePart clonedSlidePart)
    {
        using var sourceStream = this.sourceSlidePart.GetStream();
        sourceStream.Position = 0;
        using var destinationStream = clonedSlidePart.GetStream(FileMode.Create, FileAccess.Write);
        sourceStream.CopyTo(destinationStream);
    }

    private void CopyCustomXmlParts(SlidePart clonedSlidePart)
    {
        var sourceCustomXmlParts = this.sourceSlidePart.CustomXmlParts.ToList();
        if (!sourceCustomXmlParts.Any())
        {
            return;
        }

        foreach (var sourceCustomXmlPart in sourceCustomXmlParts)
        {
            var newCustomXmlPart = clonedSlidePart.AddCustomXmlPart(sourceCustomXmlPart.ContentType);
            using var sourceStream = sourceCustomXmlPart.GetStream();
            sourceStream.Position = 0;
            using var destinationStream = newCustomXmlPart.GetStream(FileMode.Create, FileAccess.Write);
            sourceStream.CopyTo(destinationStream);
        }
    }

    private void CopyImageParts(SlidePart clonedSlidePart)
    {
        var blips = clonedSlidePart.Slide.CommonSlideData!
            .ShapeTree!
            .Descendants<Blip>()
            .ToList();

        foreach (var blip in blips)
        {
            var relId = blip.Embed?.Value;
            if (string.IsNullOrWhiteSpace(relId))
            {
                continue;
            }

            if (clonedSlidePart.Parts.Any(part => part.RelationshipId == relId))
            {
                continue;
            }

            if (this.sourceSlidePart.TryGetPartById(relId!, out var openXmlPart) &&
                openXmlPart is ImagePart sourceImage)
            {
                if (ReferenceEquals(this.sourceSlidePart.OpenXmlPackage, clonedSlidePart.OpenXmlPackage))
                {
                    clonedSlidePart.AddPart(sourceImage, relId!);
                }
                else
                {
                    var destinationImage = clonedSlidePart.AddNewPart<ImagePart>(sourceImage.ContentType, relId);
                    using var sourceStream = sourceImage.GetStream();
                    sourceStream.Position = 0;
                    using var destinationStream = destinationImage.GetStream(FileMode.Create, FileAccess.Write);
                    sourceStream.CopyTo(destinationStream);
                }
            }
        }
    }

    private void CopyChartParts(SlidePart clonedSlidePart)
    {
        foreach (var relationshipId in GetChartRelationshipIds(clonedSlidePart))
        {
            this.EnsureChartRelationship(relationshipId, clonedSlidePart);
        }
    }

    private void EnsureChartRelationship(string relationshipId, SlidePart targetSlidePart)
    {
        if (RelationshipExists(targetSlidePart, relationshipId))
        {
            return;
        }

        if (!this.TryGetSourceChartPart(relationshipId, out var sourceChartPart))
        {
            return;
        }

        if (ReferenceEquals(this.sourceSlidePart.OpenXmlPackage, targetSlidePart.OpenXmlPackage))
        {
            ShareChartPartWithinSamePackage(sourceChartPart!, targetSlidePart, relationshipId);
            return;
        }

        CloneChartPartAcrossPackages(sourceChartPart!, targetSlidePart, relationshipId);
    }
}