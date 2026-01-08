using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using A16 = DocumentFormat.OpenXml.Office2016.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

/// <summary>
///     Represents an Open XML picture element.
/// </summary>
internal sealed class XmlPicture(SlidePart slidePart, uint shapeId, string shapeName)
{
    /// <summary>
    ///     Sets the transform (position and size) of a picture element.
    /// </summary>
    internal static void SetTransform(P.Picture pPicture, uint width, uint height)
    {
        var transform2D = pPicture.ShapeProperties!.Transform2D!;
        transform2D.Offset!.X = transform2D.Offset!.Y = 952500;
        transform2D.Extents!.Cx = new Pixels(width).AsHorizontalEmus();
        transform2D.Extents!.Cy = new Pixels(height).AsVerticalEmus();
    }

    /// <summary>
    ///     Creates a standard P.Picture element.
    /// </summary>
    internal P.Picture CreatePPicture(string imagePartRId)
    {
        var nonVisualPictureProperties = new P.NonVisualPictureProperties();
        var nonVisualDrawingProperties = new P.NonVisualDrawingProperties
        {
            Id = shapeId,
            Name = $"{shapeName} {shapeId}"
        };
        var nonVisualPictureDrawingProperties = new P.NonVisualPictureDrawingProperties();
        var appNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties();

        nonVisualPictureProperties.Append(nonVisualDrawingProperties);
        nonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);
        nonVisualPictureProperties.Append(appNonVisualDrawingProperties);

        var blipFill = new P.BlipFill();
        var blip = new A.Blip { Embed = imagePartRId };
        var stretch = new A.Stretch();
        blipFill.Append(blip);
        blipFill.Append(stretch);

        var transform2D = new A.Transform2D(
            new A.Offset { X = 0, Y = 0 },
            new A.Extents { Cx = 0, Cy = 0 });

        var presetGeometry = new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle };
        var shapeProperties = new P.ShapeProperties();
        shapeProperties.Append(transform2D);
        shapeProperties.Append(presetGeometry);

        var pPicture = new P.Picture();
        pPicture.Append(nonVisualPictureProperties);
        pPicture.Append(blipFill);
        pPicture.Append(shapeProperties);

        slidePart.Slide!.CommonSlideData!.ShapeTree!.Append(pPicture);

        return pPicture;
    }

    /// <summary>
    ///     Creates an SVG P.Picture element with both vector and raster representations.
    /// </summary>
    internal P.Picture CreateSvgPPicture(string imagePartRId, string svgPartRId)
    {
        var nonVisualPictureProperties = new P.NonVisualPictureProperties();
        var nonVisualDrawingProperties = new P.NonVisualDrawingProperties
        {
            Id = shapeId,
            Name = $"{shapeName} {shapeId}"
        };
        var nonVisualPictureDrawingProperties = new P.NonVisualPictureDrawingProperties();
        var appNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties();

        var aNonVisualDrawingPropertiesExtensionList =
            new A.NonVisualDrawingPropertiesExtensionList();

        var aNonVisualDrawingPropertiesExtension =
            new A.NonVisualDrawingPropertiesExtension { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

        var a16CreationId = new A16.CreationId();

        // "http://schemas.microsoft.com/office/drawing/2014/main"
        var a16 = DocumentFormat.OpenXml.Linq.A16.a16;
        a16CreationId.AddNamespaceDeclaration(nameof(a16), a16.NamespaceName);

        a16CreationId.Id = "{2BEA8DB4-11C1-B7BA-06ED-DC504E2BBEBE}";

        aNonVisualDrawingPropertiesExtension.AppendChild(a16CreationId);

        aNonVisualDrawingPropertiesExtensionList.AppendChild(aNonVisualDrawingPropertiesExtension);

        nonVisualDrawingProperties.AppendChild(aNonVisualDrawingPropertiesExtensionList);
        nonVisualPictureProperties.AppendChild(nonVisualDrawingProperties);
        nonVisualPictureProperties.AppendChild(nonVisualPictureDrawingProperties);
        nonVisualPictureProperties.AppendChild(appNonVisualDrawingProperties);

        var blipFill = new P.BlipFill();
        var aBlip = new A.Blip { Embed = imagePartRId };
        var aBlipExtensionList = new A.BlipExtensionList();
        var aBlipExtension = new A.BlipExtension { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };
        var a14UseLocalDpi = new A14.UseLocalDpi();

        // "http://schemas.microsoft.com/office/drawing/2010/main"
        var a14 = DocumentFormat.OpenXml.Linq.A14.a14;

        a14UseLocalDpi.AddNamespaceDeclaration(nameof(a14), a14.NamespaceName);
        a14UseLocalDpi.Val = false;
        aBlipExtension.AppendChild(a14UseLocalDpi);
        aBlipExtensionList.AppendChild(aBlipExtension);
        aBlipExtension = new A.BlipExtension { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };
        var svgBlip = new DocumentFormat.OpenXml.Office2019.Drawing.SVG.SVGBlip() { Embed = svgPartRId };

        // "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
        var asvg = DocumentFormat.OpenXml.Linq.ASVG.asvg;

        svgBlip.AddNamespaceDeclaration(nameof(asvg), asvg.NamespaceName);
        aBlipExtension.AppendChild(svgBlip);
        aBlipExtensionList.AppendChild(aBlipExtension);
        aBlip.AppendChild(aBlipExtensionList);
        blipFill.AppendChild(aBlip);
        var aStretch = new A.Stretch();
        var aFillRectangle = new A.FillRectangle();
        aStretch.AppendChild(aFillRectangle);

        blipFill.AppendChild(aStretch);

        var transform2D = new A.Transform2D(
            new A.Offset { X = 0, Y = 0 },
            new A.Extents { Cx = 0, Cy = 0 });

        var presetGeometry = new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle };

        var aAdjustValueList = new A.AdjustValueList();

        presetGeometry.AppendChild(aAdjustValueList);

        var shapeProperties = new P.ShapeProperties();
        shapeProperties.AppendChild(transform2D);
        shapeProperties.AppendChild(presetGeometry);

        var pPicture = new P.Picture();
        pPicture.AppendChild(nonVisualPictureProperties);
        pPicture.AppendChild(blipFill);
        pPicture.AppendChild(shapeProperties);

        slidePart.Slide!.CommonSlideData!.ShapeTree!.AppendChild(pPicture);

        return pPicture;
    }
}