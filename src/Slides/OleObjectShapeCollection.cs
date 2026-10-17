using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ImageMagick;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

internal sealed class OleObjectShapeCollection(SlidePart slidePart)
{
    internal void Add(int x, int y, int width, int height, Stream content, OleObjectType type, Stream previewImage)
    {
        ArgumentNullException.ThrowIfNull(content);
        ArgumentNullException.ThrowIfNull(previewImage);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(height);
        var (progId, contentType) = DocumentType(type);
        var cx = new Points(width).AsEmus();
        var cy = new Points(height).AsEmus();
        var imageWidth = checked((int)cx);
        var imageHeight = checked((int)cy);

        // Read both inputs before modifying the slide.
        using var document = new MemoryStream();
        content.CopyTo(document);
        document.Position = 0;
        using var image = new MagickImage(previewImage);
        using var preview = new MemoryStream(image.ToByteArray(MagickFormat.Png));

        var tree = slidePart.Slide!.CommonSlideData!.ShapeTree!;
        var id = tree.Descendants<P.NonVisualDrawingProperties>().Select(p => p.Id!.Value).DefaultIfEmpty(0U).Max() + 1;
        var package = slidePart.AddNewPart<EmbeddedPackagePart>(contentType);
        package.FeedData(document);
        var imagePart = slidePart.AddImagePart(ImagePartType.Png);
        imagePart.FeedData(preview);

        var picture = PreviewPicture(id + 1, slidePart.GetIdOfPart(imagePart), x, y, cx, cy);
        var ole = new P.OleObject(new P.OleObjectEmbed(), picture)
        {
            Id = slidePart.GetIdOfPart(package),
            ProgId = progId,
            Name = type.ToString(),
            ImageWidth = imageWidth,
            ImageHeight = imageHeight
        };
        var frame = new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = $"Object {id}" },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.Transform(
                new A.Offset { X = new Points(x).AsEmus(), Y = new Points(y).AsEmus() },
                new A.Extents { Cx = cx, Cy = cy }),
            new A.Graphic(new A.GraphicData(ole)
            {
                Uri = "http://schemas.openxmlformats.org/presentationml/2006/ole"
            }));
        tree.Append(frame);
    }

    private static (string ProgId, string ContentType) DocumentType(OleObjectType type) => type switch
    {
        OleObjectType.WordDocument =>
            ("Word.Document.12", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
        OleObjectType.ExcelWorkbook =>
            ("Excel.Sheet.12", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        OleObjectType.PowerPointPresentation =>
            ("PowerPoint.Show.12", "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported OLE document type.")
    };

    private static P.Picture PreviewPicture(uint id, string imageId, int x, int y, long width, long height) =>
        new(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = $"Object preview {id}" },
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(new A.Blip { Embed = imageId }, new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = new Points(x).AsEmus(), Y = new Points(y).AsEmus() },
                    new A.Extents { Cx = width, Cy = height }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
}
