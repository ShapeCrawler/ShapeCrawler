using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Audio;
using ShapeCrawler.Charts;
using ShapeCrawler.Drawing;
using ShapeCrawler.Factories;
using ShapeCrawler.OLEObjects;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Settings;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using ShapeCrawler.Statics;
using ShapeCrawler.Tables;
using ShapeCrawler.Video;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represents a collection of a slide shapes.
    /// </summary>
    public class ShapeCollection : LibraryCollection<IShape>, IShapeCollection
    {
        private readonly P.ShapeTree shapeTree;
        private readonly SCSlide slide;

        #region Constructors

        internal ShapeCollection(List<IShape> shapes)
        {
            this.CollectionItems = shapes;
        }

        internal ShapeCollection(List<IShape> shapes, P.ShapeTree shapeTree, SCSlide slide)
        {
            this.slide = slide;
            this.CollectionItems = shapes;
            this.shapeTree = shapeTree;
        }

        #endregion Constructors

        internal static ShapeCollection CreateForSlide(SlidePart slidePart, SCSlide slide)
        {
            var shapeContextBuilder = new ShapeContext.Builder(slidePart);

            var chartGrFrameHandler = new ChartGraphicFrameHandler();
            var tableGrFrameHandler = new TableGraphicFrameHandler(shapeContextBuilder);
            var oleGrFrameHandler = new OleGraphicFrameHandler(shapeContextBuilder);
            var autoShapeCreator = new AutoShapeCreator();
            var pictureHandler = new PictureHandler();

            autoShapeCreator.Successor = oleGrFrameHandler;
            oleGrFrameHandler.Successor = pictureHandler;
            pictureHandler.Successor = chartGrFrameHandler;
            chartGrFrameHandler.Successor = tableGrFrameHandler;

            P.ShapeTree shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;
            var shapes = new List<IShape>(shapeTree.Count());
            foreach (OpenXmlCompositeElement shapeTreesChildElement in shapeTree.OfType<OpenXmlCompositeElement>())
            {
                IShape shape;
                if (shapeTreesChildElement is P.GroupShape pGroupShape)
                {
                    shape = new SlideGroupShape(pGroupShape, slide, null);
                }
                else
                {
                    shape = autoShapeCreator.Create(shapeTreesChildElement, slide, null);
                }

                if (shape != null)
                {
                    shapes.Add(shape);
                }
            }

            return new ShapeCollection(shapes, shapeTree, slide);
        }

        internal static ShapeCollection CreateForSlideLayout(P.ShapeTree pShapeTree, SCSlideLayout slideLayout)
        {
            var shapeList = new List<IShape>();
            foreach (OpenXmlCompositeElement compositeElement in pShapeTree.OfType<OpenXmlCompositeElement>())
            {
                switch (compositeElement)
                {
                    case P.Shape pShape:
                        shapeList.Add(new LayoutAutoShape(slideLayout, pShape));
                        continue;
                    case P.GraphicFrame pGraphicFrame:
                    {
                        A.GraphicData aGraphicData =
                            pGraphicFrame.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                        if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/presentationml/2006/ole",
                            StringComparison.Ordinal))
                        {
                            shapeList.Add(new LayoutOLEObject(slideLayout, pGraphicFrame));
                            continue;
                        }

                        if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/chart",
                            StringComparison.Ordinal))
                        {
                            shapeList.Add(new LayoutChart(slideLayout, pGraphicFrame));
                            continue;
                        }

                        if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/table",
                            StringComparison.Ordinal))
                        {
                            shapeList.Add(new LayoutTable(slideLayout, pGraphicFrame));
                            continue;
                        }

                        break;
                    }
                }

                // OLE Objects should be parsed before pictures, since OLE containers can contain p:pic elements,
                // thus OLE objects can be parsed as a picture by mistake.
                P.Picture pPicture;
                if (compositeElement is P.Picture treePicture)
                {
                    pPicture = treePicture;
                }
                else
                {
                    P.Picture framePicture = compositeElement.Descendants<P.Picture>().FirstOrDefault();
                    pPicture = framePicture;
                }

                if (pPicture != null)
                {
                    shapeList.Add(new LayoutPicture(slideLayout, pPicture));
                }
            }

            return new ShapeCollection(shapeList);
        }

        internal static ShapeCollection CreateForSlideMaster(SCSlideMaster slideMaster)
        {
            P.ShapeTree pShapeTree = slideMaster.PSlideMaster.CommonSlideData.ShapeTree;
            var shapeList = new List<IShape>();
            foreach (OpenXmlCompositeElement compositeElement in pShapeTree.OfType<OpenXmlCompositeElement>())
            {
                switch (compositeElement)
                {
                    case P.Shape pShape:
                        shapeList.Add(new MasterAutoShape(pShape, slideMaster));
                        continue;
                    case P.GraphicFrame pGraphicFrame:
                    {
                        A.GraphicData aGraphicData =
                            pGraphicFrame.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                        if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/presentationml/2006/ole",
                            StringComparison.Ordinal))
                        {
                            shapeList.Add(new MasterOLEObject(pGraphicFrame, slideMaster));
                            continue;
                        }

                        if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/chart",
                            StringComparison.Ordinal))
                        {
                            shapeList.Add(new MasterChart(slideMaster, pGraphicFrame));
                            continue;
                        }

                        if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/table",
                            StringComparison.Ordinal))
                        {
                            shapeList.Add(new MasterTable(pGraphicFrame, slideMaster));
                            continue;
                        }

                        break;
                    }
                }

                // OLE Objects should be parsed before pictures, since OLE containers can contain p:pic elements,
                // thus OLE objects can be parsed as a picture by mistake.
                P.Picture pPicture;
                if (compositeElement is P.Picture treePicture)
                {
                    pPicture = treePicture;
                }
                else
                {
                    P.Picture framePicture = compositeElement.Descendants<P.Picture>().FirstOrDefault();
                    pPicture = framePicture;
                }

                if (pPicture != null)
                {
                    shapeList.Add(new MasterPicture(pPicture, slideMaster));
                }
            }

            return new ShapeCollection(shapeList);
        }

        public IAudioShape AddNewAudio(int xPixels, int yPixels, Stream mp3Stream)
        {
            long xEmu = PixelConverter.HorizontalPixelToEmu(xPixels);
            long yEmu = PixelConverter.VerticalPixelToEmu(yPixels);

            MediaDataPart mediaDataPart = this.slide.ParentPresentation.PresentationDocument.CreateMediaDataPart("audio/mpeg", ".mp3");

            mp3Stream.Position = 0;
            mediaDataPart.FeedData(mp3Stream);
            string imgPartRId = $"rId{Guid.NewGuid().ToString().Replace("-", string.Empty).Substring(0, 5)}";
            ImagePart imagePart = this.slide.SlidePart.AddNewPart<ImagePart>("image/png", imgPartRId);
            MemoryStream ms = new ();
            Properties.Resources.audio_image.Save(ms, ImageFormat.Png);
            ms.Position = 0;
            imagePart.FeedData(ms);

            AudioReferenceRelationship audioRr = this.slide.SlidePart.AddAudioReferenceRelationship(mediaDataPart);
            MediaReferenceRelationship mediaRr = this.slide.SlidePart.AddMediaReferenceRelationship(mediaDataPart);

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
            A.Blip blip1 = new() { Embed = imgPartRId };

            A.Stretch stretch1 = new();
            A.FillRectangle fillRectangle1 = new();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            P.ShapeProperties shapeProperties1 = new ();

            A.Transform2D transform2D1 = new();
            A.Offset offset2 = new() { X = xEmu, Y = yEmu };
            A.Extents extents2 = new() { Cx = 609600L, Cy = 609600L };

            transform2D1.Append(offset2);
            transform2D1.Append(extents2);

            A.PresetGeometry presetGeometry1 = new() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);

            this.shapeTree.Append(picture1);

            P14.CreationId creationId1 = new() { Val = (UInt32Value)3972997422U };
            creationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            return new AudioShape(this.shapeTree, this.slide);
        }

        internal Shape GetShapeByPPlaceholderShape(P.PlaceholderShape inpPPlaceholderShape)
        {
            IEnumerable<Shape> placeholderShapes = CollectionItems.Where(sp => sp.Placeholder != null).OfType<Shape>();
            Shape mappedShape = placeholderShapes.FirstOrDefault(IsEqual);

            bool IsEqual(Shape collectionShape)
            {
                Placeholder placeholder = (Placeholder) collectionShape.Placeholder;
                P.PlaceholderShape colPPlaceholderShape = placeholder.SdkPPlaceholderShape;

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

                    return inpPPlaceholderShape.Type.Equals(colPPlaceholderShape.Type);
                }

                return false;
            }

            return mappedShape;
        }

        public IVideoShape AddNewVideo(int xPixels, int yPixels, Stream videoStream)
        {
            long xEmu = PixelConverter.HorizontalPixelToEmu(xPixels);
            long yEmu = PixelConverter.VerticalPixelToEmu(yPixels);

            MediaDataPart mediaDataPart = this.slide.ParentPresentation.PresentationDocument.CreateMediaDataPart("video/mp4", ".mp4");

            videoStream.Position = 0;
            mediaDataPart.FeedData(videoStream);
            string imgPartRId = $"rId{Guid.NewGuid().ToString().Replace("-", string.Empty).Substring(0, 5)}";
            ImagePart imagePart = this.slide.SlidePart.AddNewPart<ImagePart>("image/png", imgPartRId);
            MemoryStream ms = new();
            Properties.Resources.video_image.Save(ms, ImageFormat.Png);
            ms.Position = 0;
            imagePart.FeedData(ms);

            VideoReferenceRelationship videoRr = this.slide.SlidePart.AddVideoReferenceRelationship(mediaDataPart);
            MediaReferenceRelationship mediaRr = this.slide.SlidePart.AddMediaReferenceRelationship(mediaDataPart);

            P.Picture picture1 = new();

            P.NonVisualPictureProperties nonVisualPictureProperties1 = new();

            uint shapeId = (uint)this.CollectionItems.Max(sp => sp.Id) + 1;
            P.NonVisualDrawingProperties nonVisualDrawingProperties2 = new() { Id = shapeId, Name = $"Video{shapeId}" };
            A.HyperlinkOnClick hyperlinkOnClick1 = new A.HyperlinkOnClick() { Id = "", Action = "ppaction://media" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = new() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{2FF36D28-5328-4DA3-BF85-A2B65D7EE127}\" />");

            nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement1);

            nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

            nonVisualDrawingProperties2.Append(hyperlinkOnClick1);
            nonVisualDrawingProperties2.Append(nonVisualDrawingPropertiesExtensionList1);

            P.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new();
            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);

            P.ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new();
            A.VideoFromFile videoFromFile1 = new A.VideoFromFile() { Link = videoRr.Id };

            P.ApplicationNonVisualDrawingPropertiesExtensionList applicationNonVisualDrawingPropertiesExtensionList1 = new();

            P.ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension1 = new() { Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}" };

            P14.Media media1 = new P14.Media() { Embed = mediaRr.Id };
            media1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            applicationNonVisualDrawingPropertiesExtension1.Append(media1);

            applicationNonVisualDrawingPropertiesExtensionList1.Append(applicationNonVisualDrawingPropertiesExtension1);

            applicationNonVisualDrawingProperties2.Append(videoFromFile1);
            applicationNonVisualDrawingProperties2.Append(applicationNonVisualDrawingPropertiesExtensionList1);

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties2);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);
            nonVisualPictureProperties1.Append(applicationNonVisualDrawingProperties2);

            P.BlipFill blipFill1 = new();
            A.Blip blip1 = new() { Embed = imgPartRId };

            A.Stretch stretch1 = new();
            A.FillRectangle fillRectangle1 = new();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            P.ShapeProperties shapeProperties1 = new();

            A.Transform2D transform2D1 = new();
            A.Offset offset2 = new() { X = xEmu, Y = yEmu };
            A.Extents extents2 = new() { Cx = 609600L, Cy = 609600L };

            transform2D1.Append(offset2);
            transform2D1.Append(extents2);

            A.PresetGeometry presetGeometry1 = new() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);

            this.shapeTree.Append(picture1);

            P14.CreationId creationId1 = new() { Val = (UInt32Value)3972997422U };
            creationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            return new VideoShape(this.slide, this.shapeTree);
        }
    }
}