using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Assets;
using ShapeCrawler.Presentations;
using ShapeCrawler.Shapes;
using ShapeCrawler.SmartArts;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;
using Position = ShapeCrawler.Positions.Position;

#pragma warning disable SA1204 // Static members should appear before non-static members

namespace ShapeCrawler.Slides;

internal readonly ref struct SCSlidePart(SlidePart slidePart)
{
    private const string SmartArtDiagramUri = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
    private const string DiagramDrawingRelationshipType = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";
    private const string DiagramDrawingContentType = "application/vnd.ms-office.drawingml.diagramDrawing+xml";
    private const string BasicBlockListDataAsset = "smartart-basicblocklist-data.xml";
    private const string BasicBlockListLayoutAsset = "smartart-basicblocklist-layout.xml";
    private const string BasicBlockListColorsAsset = "smartart-basicblocklist-colors.xml";
    private const string BasicBlockListQuickStyleAsset = "smartart-basicblocklist-quickStyle.xml";
    private const string BasicBlockListDrawingAsset = "smartart-basicblocklist-drawing.xml";

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
        this.CopyNotesSlidePart(clonedSlidePart);
        this.CopyImageParts(clonedSlidePart);
        this.CopyChartParts(clonedSlidePart);
        this.LinkToLayoutPart(targetPresentationPart, clonedSlidePart);

        return clonedSlidePart;
    }

    internal IShape AddSmartArt(int x, int y, int width, int height, SmartArtType smartArtType)
    {
        if (smartArtType != SmartArtType.BasicBlockList)
        {
            throw new NotSupportedException($"SmartArt type '{smartArtType}' is not supported.");
        }

        var diagramPartIds = this.CreateBasicBlockListDiagramParts();
        var pGraphicFrame = new GraphicFrame();

        // Add ID and name properties
        var nvGraphicFrameProperties = new NonVisualGraphicFrameProperties();
        var nonVisualDrawingProperties = new NonVisualDrawingProperties
        {
            Id = this.GetNextShapeId(),
            Name = $"SmartArt {smartArtType}"
        };
        var nonVisualGraphicFrameDrawingProperties = new NonVisualGraphicFrameDrawingProperties();
        var applicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

        nvGraphicFrameProperties.Append(nonVisualDrawingProperties);
        nvGraphicFrameProperties.Append(nonVisualGraphicFrameDrawingProperties);
        nvGraphicFrameProperties.Append(applicationNonVisualDrawingProperties);
        pGraphicFrame.Append(nvGraphicFrameProperties);

        // Add transform properties
        var transform = new Transform();
        transform.Append(new A.Offset { X = new Points(x).AsEmus(), Y = new Points(y).AsEmus() });
        transform.Append(new A.Extents { Cx = new Points(width).AsEmus(), Cy = new Points(height).AsEmus() });
        pGraphicFrame.Append(transform);

        // Create the diagram graphic
        var graphic = new A.Graphic();
        var graphicData = new A.GraphicData { Uri = SmartArtDiagramUri, InnerXml = "<dgm:relIds xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" " +
                                                                                   "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                                                                                   $"r:dm=\"{diagramPartIds.DataId}\" " +
                                                                                   $"r:lo=\"{diagramPartIds.LayoutId}\" " +
                                                                                   $"r:qs=\"{diagramPartIds.QuickStyleId}\" " +
                                                                                   $"r:cs=\"{diagramPartIds.ColorsId}\" />"
        };
        graphic.Append(graphicData);
        pGraphicFrame.Append(graphic);

        slidePart.Slide!.CommonSlideData!.ShapeTree!.Append(pGraphicFrame);

        return
            new SmartArtShape(
                new Position(pGraphicFrame),
                new ShapeSize(pGraphicFrame),
                new ShapeId(pGraphicFrame),
                pGraphicFrame);
    }

    private static void WriteSmartArtPart(
        OpenXmlPart targetPart,
        AssetCollection assets,
        string assetName,
        string? drawingRelationshipId = null)
    {
        using var destinationStream = targetPart.GetStream(FileMode.Create, FileAccess.Write);
        if (drawingRelationshipId == null)
        {
            using var sourceStream = assets.StreamOf(assetName);
            sourceStream.CopyTo(destinationStream);
            return;
        }

        var template = assets.StringOf(assetName).Replace("{{DRAWING_REL_ID}}", drawingRelationshipId);
        using var writer = new StreamWriter(destinationStream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(template);
    }

    private DiagramPartIds CreateBasicBlockListDiagramParts()
    {
        var assets = new AssetCollection(Assembly.GetExecutingAssembly());
        var relationshipGenerator = new SCOpenXmlPart(slidePart);

        var dataRelId = relationshipGenerator.NextRelationshipId();
        var dataPart = slidePart.AddNewPart<DiagramDataPart>(dataRelId);

        var layoutRelId = relationshipGenerator.NextRelationshipId();
        var layoutPart = slidePart.AddNewPart<DiagramLayoutDefinitionPart>(layoutRelId);
        WriteSmartArtPart(layoutPart, assets, BasicBlockListLayoutAsset);

        var quickStyleRelId = relationshipGenerator.NextRelationshipId();
        var quickStylePart = slidePart.AddNewPart<DiagramStylePart>(quickStyleRelId);
        WriteSmartArtPart(quickStylePart, assets, BasicBlockListQuickStyleAsset);

        var colorsRelId = relationshipGenerator.NextRelationshipId();
        var colorsPart = slidePart.AddNewPart<DiagramColorsPart>(colorsRelId);
        WriteSmartArtPart(colorsPart, assets, BasicBlockListColorsAsset);

        var drawingRelId = relationshipGenerator.NextRelationshipId();
        var drawingPart = slidePart.AddExtendedPart(
            DiagramDrawingRelationshipType,
            DiagramDrawingContentType,
            drawingRelId);
        WriteSmartArtPart(drawingPart, assets, BasicBlockListDrawingAsset);
        WriteSmartArtPart(dataPart, assets, BasicBlockListDataAsset, drawingRelId);

        return new DiagramPartIds(dataRelId, layoutRelId, quickStyleRelId, colorsRelId);
    }

    private static void CopyStream(OpenXmlPart sourcePart, OpenXmlPart targetPart)
    {
        using var sourceStream = sourcePart.GetStream();
        sourceStream.Position = 0;
        using var destinationStream = targetPart.GetStream(FileMode.Create, FileAccess.Write);
        sourceStream.CopyTo(destinationStream);
    }

    private static IEnumerable<string> GetChartRelationshipIds(SlidePart slidePart)
    {
        var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
        if (shapeTree == null)
        {
            return [];
        }

        return shapeTree.Descendants<A.GraphicData>()
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

    private static bool LayoutsMatch(SlideLayoutPart layout1, SlideLayoutPart layout2)
    {
        if (layout1.SlideLayout!.Type != null && layout2.SlideLayout!.Type != null)
        {
            return layout1.SlideLayout!.Type!.Value == layout2.SlideLayout!.Type!.Value;
        }

        var name1 = layout1.SlideLayout!.CommonSlideData?.Name?.Value;
        var name2 = layout2.SlideLayout!.CommonSlideData?.Name?.Value;

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
        CopyStream(sourceLayoutPart, targetLayoutPart);
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
        var sourceMasterPart = sourceLayoutPart.SlideMasterPart;
        if (sourceMasterPart != null)
        {
            CopyStream(sourceMasterPart, masterPart);
        }

        return masterPart;
    }

    private bool TryGetSourceChartPart(string relationshipId, out ChartPart? sourceChartPart)
    {
        sourceChartPart = null;
        if (slidePart.TryGetPartById(relationshipId, out var part) && part is ChartPart chartPart)
        {
            sourceChartPart = chartPart;
            return true;
        }

        return false;
    }

    private void LinkToLayoutPart(PresentationPart presentationPart, SlidePart clonedSlidePart)
    {
        var sourceLayoutPart = slidePart.SlideLayoutPart;
        if (sourceLayoutPart == null)
        {
            return;
        }

        if (ReferenceEquals(slidePart.OpenXmlPackage, presentationPart.OpenXmlPackage))
        {
            clonedSlidePart.AddPart(sourceLayoutPart);
            return;
        }

        var targetLayoutPart = FindMatchingLayout(presentationPart, sourceLayoutPart) ??
                               CreateNewLayout(presentationPart, sourceLayoutPart);

        clonedSlidePart.AddPart(targetLayoutPart);
    }

    private void CopySlideContent(SlidePart clonedSlidePart)
    {
        slidePart.Slide?.Save();
        using var sourceStream = slidePart.GetStream();
        sourceStream.Position = 0;
        using var destinationStream = clonedSlidePart.GetStream(FileMode.Create, FileAccess.Write);
        sourceStream.CopyTo(destinationStream);
    }

    private void CopyCustomXmlParts(SlidePart clonedSlidePart)
    {
        var sourceCustomXmlParts = slidePart.CustomXmlParts.ToList();
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
        var shapeTree = clonedSlidePart.Slide?.CommonSlideData?.ShapeTree;
        if (shapeTree == null)
        {
            return;
        }

        var blips = shapeTree.Descendants<A.Blip>().ToList();

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

            if (!slidePart.TryGetPartById(relId!, out var openXmlPart) ||
                openXmlPart is not ImagePart sourceImage)
            {
                continue;
            }

            if (ReferenceEquals(slidePart.OpenXmlPackage, clonedSlidePart.OpenXmlPackage))
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

    private void CopyChartParts(SlidePart clonedSlidePart)
    {
        foreach (var relationshipId in GetChartRelationshipIds(clonedSlidePart))
        {
            this.EnsureChartRelationship(relationshipId, clonedSlidePart);
        }
    }

    private void CopyNotesSlidePart(SlidePart clonedSlidePart)
    {
        var sourceNotesPart = slidePart.NotesSlidePart;
        if (sourceNotesPart == null)
        {
            return;
        }

        var targetNotesPart = clonedSlidePart.AddNewPart<NotesSlidePart>(sourceNotesPart.ContentType);
        CopyStream(sourceNotesPart, targetNotesPart);

        foreach (var childPart in sourceNotesPart.Parts)
        {
            targetNotesPart.AddPart(childPart.OpenXmlPart, childPart.RelationshipId);
        }

        if (targetNotesPart.NotesMasterPart != null)
        {
            targetNotesPart.DeletePart(targetNotesPart.NotesMasterPart);
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

        if (ReferenceEquals(slidePart.OpenXmlPackage, targetSlidePart.OpenXmlPackage))
        {
            ShareChartPartWithinSamePackage(sourceChartPart!, targetSlidePart, relationshipId);
            return;
        }

        CloneChartPartAcrossPackages(sourceChartPart!, targetSlidePart, relationshipId);
    }

    private uint GetNextShapeId()
    {
        // Get all existing shape IDs from the slide
        var shapeIds = slidePart.Slide!.Descendants<NonVisualDrawingProperties>()
            .Select(p => p.Id?.Value ?? 0)
            .ToList();

        // Find the maximum ID and add 1, or start with 1 if no shapes exist
        return shapeIds.Count > 0 ? shapeIds.Max() + 1 : 1;
    }

    private readonly record struct DiagramPartIds(
        string DataId,
        string LayoutId,
        string QuickStyleId,
        string ColorsId);
}