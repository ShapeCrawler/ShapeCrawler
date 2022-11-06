using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OneOf;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories;
using ShapeCrawler.Media;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace ShapeCrawler.Collections;

/// <summary>
///     Represents a collection of a slide shapes.
/// </summary>
internal class ShapeCollection : LibraryCollection<IShape>, IShapeCollection
{
    private readonly P.ShapeTree shapeTree;
    private readonly OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject;

    private ShapeCollection(List<IShape> shapes, P.ShapeTree shapeTree, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject)
    : base(shapes)
    {
        this.slideObject = slideObject;
        this.shapeTree = shapeTree;
    }

    public IAudioShape AddNewAudio(int xPixels, int yPixels, Stream mp3Stream)
    {
        long xEmu = UnitConverter.HorizontalPixelToEmu(xPixels);
        long yEmu = UnitConverter.VerticalPixelToEmu(yPixels);

        var slideBase = this.slideObject.Match(slide => slide as SlideObject, layout => layout, master => master);
        var mediaDataPart = slideBase.PresentationInternal.SDKPresentationInternal.CreateMediaDataPart("audio/mpeg", ".mp3");

        mp3Stream.Position = 0;
        mediaDataPart.FeedData(mp3Stream);
        string imgPartRId = $"rId{Guid.NewGuid().ToString().Replace("-", string.Empty).Substring(0, 5)}";
        var slidePart = slideBase.TypedOpenXmlPart as SlidePart;
        var imagePart = slidePart!.AddNewPart<ImagePart>("image/png", imgPartRId);
        var imgStream = Assembly.GetExecutingAssembly().GetStream("audio-image.png");
        imgStream.Position = 0;
        imagePart.FeedData(imgStream);

        var audioRr = slidePart.AddAudioReferenceRelationship(mediaDataPart);
        var mediaRr = slidePart.AddMediaReferenceRelationship(mediaDataPart);

        P.Picture picture1 = new ();

        P.NonVisualPictureProperties nonVisualPictureProperties1 = new ();

        uint shapeId = (uint)this.CollectionItems.Max(sp => sp.Id) + 1;
        P.NonVisualDrawingProperties nonVisualDrawingProperties2 = new () { Id = shapeId, Name = $"Audio{shapeId}" };
        A.HyperlinkOnClick hyperlinkOnClick1 = new A.HyperlinkOnClick() { Id = "", Action = "ppaction://media" };

        A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new ();

        A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = new () { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

        OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{2FF36D28-5328-4DA3-BF85-A2B65D7EE127}\" />");

        nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement1);

        nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

        nonVisualDrawingProperties2.Append(hyperlinkOnClick1);
        nonVisualDrawingProperties2.Append(nonVisualDrawingPropertiesExtensionList1);

        P.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new ();
        A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

        nonVisualPictureDrawingProperties1.Append(pictureLocks1);

        P.ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new ();
        A.AudioFromFile audioFromFile1 = new A.AudioFromFile() { Link = audioRr.Id };

        P.ApplicationNonVisualDrawingPropertiesExtensionList applicationNonVisualDrawingPropertiesExtensionList1 = new ();

        P.ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension1 = new () { Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}" };

        P14.Media media1 = new P14.Media() { Embed = mediaRr.Id };
        media1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

        applicationNonVisualDrawingPropertiesExtension1.Append(media1);

        applicationNonVisualDrawingPropertiesExtensionList1.Append(applicationNonVisualDrawingPropertiesExtension1);

        applicationNonVisualDrawingProperties2.Append(audioFromFile1);
        applicationNonVisualDrawingProperties2.Append(applicationNonVisualDrawingPropertiesExtensionList1);

        nonVisualPictureProperties1.Append(nonVisualDrawingProperties2);
        nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);
        nonVisualPictureProperties1.Append(applicationNonVisualDrawingProperties2);

        P.BlipFill blipFill1 = new ();
        A.Blip blip1 = new () { Embed = imgPartRId };

        A.Stretch stretch1 = new ();
        A.FillRectangle fillRectangle1 = new ();

        stretch1.Append(fillRectangle1);

        blipFill1.Append(blip1);
        blipFill1.Append(stretch1);

        P.ShapeProperties shapeProperties1 = new ();

        A.Transform2D transform2D1 = new ();
        A.Offset offset2 = new () { X = xEmu, Y = yEmu };
        A.Extents extents2 = new () { Cx = 609600L, Cy = 609600L };

        transform2D1.Append(offset2);
        transform2D1.Append(extents2);

        A.PresetGeometry presetGeometry1 = new () { Preset = A.ShapeTypeValues.Rectangle };
        A.AdjustValueList adjustValueList1 = new ();

        presetGeometry1.Append(adjustValueList1);

        shapeProperties1.Append(transform2D1);
        shapeProperties1.Append(presetGeometry1);

        picture1.Append(nonVisualPictureProperties1);
        picture1.Append(blipFill1);
        picture1.Append(shapeProperties1);

        this.shapeTree.Append(picture1);

        P14.CreationId creationId1 = new () { Val = (UInt32Value)3972997422U };
        creationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

        return new AudioShape(this.shapeTree, this.slideObject);
    }

    public IVideoShape AddNewVideo(int xPixels, int yPixels, Stream videoStream)
    {
        long xEmu = UnitConverter.HorizontalPixelToEmu(xPixels);
        long yEmu = UnitConverter.VerticalPixelToEmu(yPixels);

        var slideBase = this.slideObject.Match(slide => slide as SlideObject, layout => layout, master => master);
        MediaDataPart mediaDataPart = slideBase.PresentationInternal.SDKPresentationInternal.CreateMediaDataPart("video/mp4", ".mp4");

        videoStream.Position = 0;
        mediaDataPart.FeedData(videoStream);
        string imgPartRId = $"rId{Guid.NewGuid().ToString().Replace("-", string.Empty).Substring(0, 5)}";
        var slidePart = slideBase.TypedOpenXmlPart as SlidePart;
        var imagePart = slidePart!.AddNewPart<ImagePart>("image/png", imgPartRId);
        var imageStream = Assembly.GetExecutingAssembly().GetStream("video-image.bmp");
        imagePart.FeedData(imageStream);

        var videoRr = slidePart.AddVideoReferenceRelationship(mediaDataPart);
        var mediaRr = slidePart.AddMediaReferenceRelationship(mediaDataPart);

        P.Picture picture1 = new ();

        P.NonVisualPictureProperties nonVisualPictureProperties1 = new ();

        uint shapeId = (uint)this.CollectionItems.Max(sp => sp.Id) + 1;
        P.NonVisualDrawingProperties nonVisualDrawingProperties2 = new () { Id = shapeId, Name = $"Video{shapeId}" };
        A.HyperlinkOnClick hyperlinkOnClick1 = new A.HyperlinkOnClick() { Id = "", Action = "ppaction://media" };

        A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new ();

        A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = new () { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

        OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{2FF36D28-5328-4DA3-BF85-A2B65D7EE127}\" />");

        nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement1);

        nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

        nonVisualDrawingProperties2.Append(hyperlinkOnClick1);
        nonVisualDrawingProperties2.Append(nonVisualDrawingPropertiesExtensionList1);

        P.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new ();
        A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

        nonVisualPictureDrawingProperties1.Append(pictureLocks1);

        P.ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new ();
        A.VideoFromFile videoFromFile1 = new A.VideoFromFile() { Link = videoRr.Id };

        P.ApplicationNonVisualDrawingPropertiesExtensionList applicationNonVisualDrawingPropertiesExtensionList1 = new ();

        P.ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension1 = new () { Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}" };

        P14.Media media1 = new P14.Media() { Embed = mediaRr.Id };
        media1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

        applicationNonVisualDrawingPropertiesExtension1.Append(media1);

        applicationNonVisualDrawingPropertiesExtensionList1.Append(applicationNonVisualDrawingPropertiesExtension1);

        applicationNonVisualDrawingProperties2.Append(videoFromFile1);
        applicationNonVisualDrawingProperties2.Append(applicationNonVisualDrawingPropertiesExtensionList1);

        nonVisualPictureProperties1.Append(nonVisualDrawingProperties2);
        nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);
        nonVisualPictureProperties1.Append(applicationNonVisualDrawingProperties2);

        P.BlipFill blipFill1 = new ();
        A.Blip blip1 = new () { Embed = imgPartRId };

        A.Stretch stretch1 = new ();
        A.FillRectangle fillRectangle1 = new ();

        stretch1.Append(fillRectangle1);

        blipFill1.Append(blip1);
        blipFill1.Append(stretch1);

        P.ShapeProperties shapeProperties1 = new ();

        A.Transform2D transform2D1 = new ();
        A.Offset offset2 = new () { X = xEmu, Y = yEmu };
        A.Extents extents2 = new () { Cx = 609600L, Cy = 609600L };

        transform2D1.Append(offset2);
        transform2D1.Append(extents2);

        A.PresetGeometry presetGeometry1 = new () { Preset = A.ShapeTypeValues.Rectangle };
        A.AdjustValueList adjustValueList1 = new ();

        presetGeometry1.Append(adjustValueList1);

        shapeProperties1.Append(transform2D1);
        shapeProperties1.Append(presetGeometry1);

        picture1.Append(nonVisualPictureProperties1);
        picture1.Append(blipFill1);
        picture1.Append(shapeProperties1);

        this.shapeTree.Append(picture1);

        P14.CreationId creationId1 = new () { Val = (UInt32Value)3972997422U };
        creationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

        return new VideoShape(this.slideObject, this.shapeTree);
    }

    public T GetById<T>(int shapeId)
        where T : IShape
    {
        var shape = this.CollectionItems.First(shape => shape.Id == shapeId);
        return (T)shape;
    }

    public T GetByName<T>(string shapeName)
        where T : IShape
    {
        var shape = this.CollectionItems.First(shape => shape.Name == shapeName);
        return (T)shape;
    }

    public Shape? GetReferencedShapeOrDefault(P.PlaceholderShape inputPPlaceholderShape)
    {
        var collectionShapes = this.CollectionItems.Where(sp => sp.Placeholder != null).OfType<Shape>();
        var mappedShape = collectionShapes.FirstOrDefault(IsEqual);

        bool IsEqual(Shape collectionShape)
        {
            var placeholder = (Placeholder)collectionShape.Placeholder!;
            var colPPlaceholderShape = placeholder.PPlaceholderShape;

            if (inputPPlaceholderShape.Index is not null && colPPlaceholderShape.Index is not null &&
                inputPPlaceholderShape.Index == colPPlaceholderShape.Index)
            {
                return true;
            }

            if (inputPPlaceholderShape.Type != null && colPPlaceholderShape.Type != null)
            {
                if (inputPPlaceholderShape.Type == P.PlaceholderValues.Body &&
                    inputPPlaceholderShape.Index is not null && colPPlaceholderShape.Index is not null )
                {
                    return inputPPlaceholderShape.Index == colPPlaceholderShape.Index;
                }

                var left = inputPPlaceholderShape.Type;
                if (inputPPlaceholderShape.Type == PlaceholderValues.CenteredTitle)
                {
                    left = PlaceholderValues.Title;
                }

                var right = colPPlaceholderShape.Type;
                if (colPPlaceholderShape.Type == PlaceholderValues.CenteredTitle)
                {
                    right = PlaceholderValues.Title;
                }

                return left.Equals(right);
            }

            return false;
        }

        return mappedShape;
    }

    internal static ShapeCollection Create(OneOf<SlidePart, SlideLayoutPart, SlideMasterPart> oneOfSlidePart, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> oneOfSlide)
    {
        var chartGrFrameHandler = new ChartGraphicFrameHandler();
        var tableGrFrameHandler = new TableGraphicFrameHandler();
        var oleGrFrameHandler = new OleGraphicFrameHandler();
        var autoShapeCreator = new AutoShapeCreator();
        var pictureHandler = new PictureHandler();

        autoShapeCreator.Successor = oleGrFrameHandler;
        oleGrFrameHandler.Successor = pictureHandler;
        pictureHandler.Successor = chartGrFrameHandler;
        chartGrFrameHandler.Successor = tableGrFrameHandler;

        var pShapeTree = oneOfSlidePart.Match(
            slidePart => slidePart.Slide.CommonSlideData!.ShapeTree!,
            layoutPart => layoutPart.SlideLayout.CommonSlideData!.ShapeTree!,
            masterPart => masterPart.SlideMaster.CommonSlideData!.ShapeTree!);
        var shapes = new List<IShape>(pShapeTree.Count());
        foreach (var childElementOfShapeTree in pShapeTree.OfType<OpenXmlCompositeElement>())
        {
            IShape shape;
            if (childElementOfShapeTree is P.GroupShape pGroupShape)
            {
                shape = new SCGroupShape(pGroupShape, oneOfSlide, null);
            }
            else if (childElementOfShapeTree is P.ConnectionShape)
            {
                shape = new SCConnectionShape(childElementOfShapeTree, oneOfSlide);
            }
            else
            {
                shape = autoShapeCreator.Create(childElementOfShapeTree, oneOfSlide, null);
            }

            if (shape != null)
            {
                shapes.Add(shape);
            }
        }

        return new ShapeCollection(shapes, pShapeTree, oneOfSlide);
    }
}