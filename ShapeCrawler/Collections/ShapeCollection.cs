using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Charts;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories;
using ShapeCrawler.Media;
using ShapeCrawler.OLEObjects;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using ShapeCrawler.Statics;
using ShapeCrawler.Tables;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represents a collection of a slide shapes.
    /// </summary>
    internal class ShapeCollection : LibraryCollection<IShape>, IShapeCollection
    {
        private readonly P.ShapeTree shapeTree;
        private readonly SCSlide slide;

        private ShapeCollection(List<IShape> shapes)
        {
            this.CollectionItems = shapes;
        }

        private ShapeCollection(List<IShape> shapes, P.ShapeTree shapeTree, SCSlide slide)
        {
            this.slide = slide;
            this.CollectionItems = shapes;
            this.shapeTree = shapeTree;
        }

        public IAudioShape AddNewAudio(int xPixels, int yPixels, Stream mp3Stream)
        {
            long xEmu = PixelConverter.HorizontalPixelToEmu(xPixels);
            long yEmu = PixelConverter.VerticalPixelToEmu(yPixels);

            MediaDataPart mediaDataPart = this.slide.PresentationInternal.SdkPresentation.CreateMediaDataPart("audio/mpeg", ".mp3");

            mp3Stream.Position = 0;
            mediaDataPart.FeedData(mp3Stream);
            string imgPartRId = $"rId{Guid.NewGuid().ToString().Replace("-", string.Empty).Substring(0, 5)}";
            var slidePart = this.slide.SDKSlidePart;
            var imagePart = slidePart.AddNewPart<ImagePart>("image/png", imgPartRId);
            var imgStream = Assembly.GetExecutingAssembly().GetStream("audio-image.png");
            imgStream.Position = 0;
            imagePart.FeedData(imgStream);

            AudioReferenceRelationship audioRr = slidePart.AddAudioReferenceRelationship(mediaDataPart);
            MediaReferenceRelationship mediaRr = slidePart.AddMediaReferenceRelationship(mediaDataPart);

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

            return new AudioShape(this.shapeTree, this.slide);
        }

        public IVideoShape AddNewVideo(int xPixels, int yPixels, Stream videoStream)
        {
            long xEmu = PixelConverter.HorizontalPixelToEmu(xPixels);
            long yEmu = PixelConverter.VerticalPixelToEmu(yPixels);

            MediaDataPart mediaDataPart = this.slide.PresentationInternal.SdkPresentation.CreateMediaDataPart("video/mp4", ".mp4");

            videoStream.Position = 0;
            mediaDataPart.FeedData(videoStream);
            string imgPartRId = $"rId{Guid.NewGuid().ToString().Replace("-", string.Empty).Substring(0, 5)}";
            var slidePart = this.slide.SDKSlidePart;
            var imagePart = slidePart.AddNewPart<ImagePart>("image/png", imgPartRId);
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

            return new VideoShape(this.slide, this.shapeTree);
        }

        public T GetById<T>(int shapeId)
            where T : IShape
        {
            var shape = this.CollectionItems.First(shape => shape.Id == shapeId);
            return (T)shape;
        }

        public T GetByName<T>(string shapeName)
        {
            var shape = this.CollectionItems.First(shape => shape.Name == shapeName);
            return (T)shape;
        }

        public Shape? GetReferencedShapeOrDefault(P.PlaceholderShape inpPPlaceholderShape)
        {
            var collectionShapes = this.CollectionItems.Where(sp => sp.Placeholder != null).OfType<Shape>();
            Shape mappedShape = collectionShapes.FirstOrDefault(IsEqual);

            bool IsEqual(Shape collectionShape)
            {
                var placeholder = (Placeholder)collectionShape.Placeholder;
                var colPPlaceholderShape = placeholder.PPlaceholderShape;

                if (inpPPlaceholderShape.Index != null && colPPlaceholderShape.Index != null &&
                    inpPPlaceholderShape.Index == colPPlaceholderShape.Index)
                {
                    return true;
                }

                if (inpPPlaceholderShape.Type != null && colPPlaceholderShape.Type != null)
                {
                    if (inpPPlaceholderShape.Type == P.PlaceholderValues.Body &&
                        inpPPlaceholderShape.Index != null && colPPlaceholderShape.Index != null)
                    {
                        return inpPPlaceholderShape.Index == colPPlaceholderShape.Index;
                    }

                    var left = inpPPlaceholderShape.Type;
                    if (inpPPlaceholderShape.Type == PlaceholderValues.CenteredTitle)
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
        
        internal static ShapeCollection ForSlideLayout(P.ShapeTree pShapeTree, SlideBase baseSlide)
        {
            var shapeList = new List<IShape>();
            var layout = baseSlide as SCSlideLayout;
            var master = baseSlide as SCSlideMaster;
            foreach (var childOfPShapeTree in pShapeTree.OfType<OpenXmlCompositeElement>())
            {
                switch (childOfPShapeTree)
                {
                    case P.Shape pShape:
                        if (layout != null)
                        {
                            shapeList.Add(new LayoutAutoShape(layout, pShape));
                        }
                        else
                        {
                            shapeList.Add(new MasterAutoShape(master!, pShape));
                        }

                        continue;
                    case P.GraphicFrame pGraphicFrame:
                    {
                        A.GraphicData aGraphicData =
                            pGraphicFrame.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                        if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/presentationml/2006/ole",
                            StringComparison.Ordinal))
                        {
                            if (layout != null)
                            {
                                shapeList.Add(new LayoutOLEObject(layout, pGraphicFrame));
                            }
                            else
                            {
                                shapeList.Add(new MasterOLEObject(master!, pGraphicFrame));
                            }

                            continue;
                        }

                        if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/chart",
                            StringComparison.Ordinal))
                        {
                            if (layout != null)
                            {
                                shapeList.Add(new LayoutChart(layout, pGraphicFrame));
                            }
                            else
                            {
                                shapeList.Add(new MasterChart(master!, pGraphicFrame));
                            }

                            continue;
                        }

                        if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/table",
                            StringComparison.Ordinal))
                        {
                            if (layout != null)
                            {
                                shapeList.Add(new LayoutTable(layout, pGraphicFrame));
                            }
                            else
                            {
                                shapeList.Add(new MasterTable(master!, pGraphicFrame));
                            }

                            continue;
                        }

                        break;
                    }
                }

                P.Picture? pPicture;
                
                if (childOfPShapeTree is P.Picture treePic)
                {
                    pPicture = treePic;
                }
                else
                {
                    pPicture = childOfPShapeTree.Descendants<P.Picture>().FirstOrDefault();
                }

                if (pPicture != null)
                {
                    var embeddedPicReference = pPicture.GetFirstChild<P.BlipFill>()?.Blip?.Embed;
                    if (embeddedPicReference != null)
                    {
                        if (layout != null)
                        {
                            shapeList.Add(new LayoutPicture(pPicture, layout, embeddedPicReference));
                        }
                        else
                        {
                            shapeList.Add(new MasterPicture(pPicture, master, embeddedPicReference));
                        }
                    }
                }
            }

            return new ShapeCollection(shapeList);
        }
        
        internal static ShapeCollection ForSlide(SlidePart slidePart, SCSlide slide)
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

            var pShapeTree = slidePart.Slide.CommonSlideData!.ShapeTree!;
            var shapes = new List<IShape>(pShapeTree.Count());
            foreach (var childElementOfShapeTree in pShapeTree.OfType<OpenXmlCompositeElement>())
            {
                IShape shape;
                if (childElementOfShapeTree is P.GroupShape pGroupShape)
                {
                    shape = new SlideGroupShape(pGroupShape, slide, null);
                }
                else if (childElementOfShapeTree is P.ConnectionShape)
                {
                    shape = new SCConnectionShape(childElementOfShapeTree, slide);
                }
                else
                {
                    shape = autoShapeCreator.Create(childElementOfShapeTree, slide, null);
                }

                if (shape != null)
                {
                    shapes.Add(shape);
                }
            }

            return new ShapeCollection(shapes, pShapeTree, slide);
        }
    }
}